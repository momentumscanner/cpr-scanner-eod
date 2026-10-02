import pandas as pd
import zipfile
import os
import re
import datetime
from openpyxl.styles import Alignment, Font, PatternFill

class CPRScanner:
    def __init__(self):
        pass

    def load_bhav_copy(self, zip_path):
        """Loads the CSV from the ZIP file into a DataFrame."""
        try:
            with zipfile.ZipFile(zip_path, 'r') as z:
                # Find the first CSV file
                csv_files = [f for f in z.namelist() if f.lower().endswith('.csv')]
                if not csv_files:
                    raise ValueError(f"No CSV found in {zip_path}")
                
                # Read CSV
                df = pd.read_csv(z.open(csv_files[0]))
                
                # Standardize columns (strip whitespace)
                df.columns = df.columns.str.strip()
                
                # Strip string columns
                str_cols = ['TckrSymb', 'FinInstrmTp', 'XpryDt', 'OptnTp']
                for c in str_cols:
                    if c in df.columns:
                        df[c] = df[c].astype(str).str.strip()

                # Convert Expiry to datetime for sorting
                df['XpryDt_Date'] = pd.to_datetime(df['XpryDt'], errors='coerce')
                
                return df
        except Exception as e:
            print(f"Error loading {zip_path}: {e}")
            return None

    def calculate_cpr(self, high, low, close):
        """Calculates Central Pivot Range (CPR) and standard Pivot levels."""
        p = (high + low + close) / 3.0
        bc = (high + low) / 2.0
        tc = (2 * p) - bc
        
        top_cpr = max(tc, bc)
        bottom_cpr = min(tc, bc)
        cpr_width = top_cpr - bottom_cpr
        cpr_width_pct = (cpr_width / p * 100.0) if p > 0 else 0.0

        r1 = (2 * p) - low
        s1 = (2 * p) - high
        r2 = p + (high - low)
        s2 = p - (high - low)
        r3 = high + 2 * (p - low)
        s3 = low - 2 * (high - p)

        return {
            'P': round(p, 2),
            'TC': round(tc, 2),
            'BC': round(bc, 2),
            'Top_CPR': round(top_cpr, 2),
            'Bottom_CPR': round(bottom_cpr, 2),
            'CPR_Width': round(cpr_width, 2),
            'CPR_Width_Pct': round(cpr_width_pct, 4),
            'R1': round(r1, 2),
            'S1': round(s1, 2),
            'R2': round(r2, 2),
            'S2': round(s2, 2),
            'R3': round(r3, 2),
            'S3': round(s3, 2)
        }

    def get_atm_strike(self, spot_price, available_strikes):
        """Returns the strike price closest to the spot price."""
        if len(available_strikes) == 0:
            return None
        return min(available_strikes, key=lambda x: abs(x - spot_price))

    def process_data(self, today_file, yesterday_file):
        print(f"Processing Today for CPR: {today_file}")
        print(f"Processing Yesterday for CPR: {yesterday_file}")

        df_today = self.load_bhav_copy(today_file)
        df_yest = self.load_bhav_copy(yesterday_file)

        if df_today is None or df_yest is None:
            return None

        # 1. Processing Yesterday's Data for Lookups
        # Support both Stock Options (STO) and Index Options (IDO)
        yest_opts = df_yest[df_yest['FinInstrmTp'].isin(['STO', 'IDO'])].copy()
        
        print("Indexing Yesterday's option data for CPR...")
        yest_lookup = {}
        for idx, row in yest_opts.iterrows():
            key = (row['TckrSymb'], float(row['StrkPric']), row['OptnTp'], row['XpryDt'])
            yest_lookup[key] = {
                'Open': row['OpnPric'],
                'High': row['HghPric'],
                'Low': row['LwPric'],
                'Close': row['ClsPric']
            }

        # 2. Process Today's Data
        today_futs = df_today[df_today['FinInstrmTp'].isin(['STF', 'IDF'])].copy()
        today_opts = df_today[df_today['FinInstrmTp'].isin(['STO', 'IDO'])].copy()

        results = []

        # Get unique symbols from futures
        symbols = today_futs['TckrSymb'].unique()
        print(f"Found {len(symbols)} underlying symbols in Futures.")

        for symbol in symbols:
            futs_sym = today_futs[today_futs['TckrSymb'] == symbol]
            if futs_sym.empty:
                continue
            
            # Find nearest expiry
            min_expiry = futs_sym['XpryDt_Date'].min()
            nearest_fut = futs_sym[futs_sym['XpryDt_Date'] == min_expiry].iloc[0]
            
            spot_close = nearest_fut['ClsPric']
            expiry_str = nearest_fut['XpryDt']
            
            # Options for this symbol and nearest expiry
            opts_sym = today_opts[
                (today_opts['TckrSymb'] == symbol) & 
                (today_opts['XpryDt'] == expiry_str)
            ]
            
            if opts_sym.empty:
                continue

            available_strikes = opts_sym['StrkPric'].astype(float).unique()
            atm_strike = self.get_atm_strike(spot_close, available_strikes)
            
            if atm_strike is None:
                continue

            for opt_type in ['CE', 'PE']:
                opt_row = opts_sym[
                    (opts_sym['StrkPric'].astype(float) == atm_strike) & 
                    (opts_sym['OptnTp'] == opt_type)
                ]
                
                if opt_row.empty:
                    continue
                
                row = opt_row.iloc[0]
                
                # Lookup Yesterday
                yest_key = (symbol, float(atm_strike), opt_type, expiry_str)
                yest_data = yest_lookup.get(yest_key)
                
                yest_levels = {}
                if yest_data:
                    yest_levels = self.calculate_cpr(
                        yest_data['High'], yest_data['Low'], yest_data['Close']
                    )
                
                today_levels = self.calculate_cpr(
                    row['HghPric'], row['LwPric'], row['ClsPric']
                )

                # CPR Logic Calculations
                is_narrow = False
                is_wide = False
                is_inside = False
                is_outside = False
                is_higher_value = False
                is_lower_value = False

                if yest_levels:
                    # Narrow CPR: Today CPR Width < Yest CPR Width
                    if today_levels['CPR_Width'] < yest_levels['CPR_Width']:
                        is_narrow = True
                    
                    # Wide CPR: Today CPR Width > Yest CPR Width
                    if today_levels['CPR_Width'] > yest_levels['CPR_Width']:
                        is_wide = True

                    # Inside CPR: Today Top < Yest Top AND Today Bottom > Yest Bottom
                    if (today_levels['Top_CPR'] < yest_levels['Top_CPR']) and (today_levels['Bottom_CPR'] > yest_levels['Bottom_CPR']):
                        is_inside = True

                    # Outside CPR: Today Top > Yest Top AND Today Bottom < Yest Bottom
                    if (today_levels['Top_CPR'] > yest_levels['Top_CPR']) and (today_levels['Bottom_CPR'] < yest_levels['Bottom_CPR']):
                        is_outside = True

                    # Higher Value CPR: Today Bottom > Yest Top
                    if today_levels['Bottom_CPR'] > yest_levels['Top_CPR']:
                        is_higher_value = True

                    # Lower Value CPR: Today Top < Yest Bottom
                    if today_levels['Top_CPR'] < yest_levels['Bottom_CPR']:
                        is_lower_value = True

                res = {
                    'Symbol': symbol,
                    'Expiry': expiry_str,
                    'Spot_Close': spot_close,
                    'ATM_Strike': atm_strike,
                    'Option_Type': opt_type,
                    'Today_Open': row['OpnPric'],
                    'Today_High': row['HghPric'],
                    'Today_Low': row['LwPric'],
                    'Today_Close': row['ClsPric'],
                    'Is_Narrow_CPR': is_narrow,
                    'Is_Wide_CPR': is_wide,
                    'Is_Inside_CPR': is_inside,
                    'Is_Outside_CPR': is_outside,
                    'Is_Higher_Value_CPR': is_higher_value,
                    'Is_Lower_Value_CPR': is_lower_value,
                    'OpnIntrst': row['OpnIntrst'],
                    'ChngInOpnIntrst': row['ChngInOpnIntrst'],
                    'TtlTradgVol': row['TtlTradgVol'],
                    'TtlNbOfTxsExctd': row['TtlNbOfTxsExctd']
                }

                # Add Today CPR levels
                for k, v in today_levels.items():
                    res[f'Today_{k}'] = v
                
                # Add Yesterday CPR levels
                if yest_levels:
                    for k, v in yest_levels.items():
                        res[f'Yest_{k}'] = v

                results.append(res)

        return pd.DataFrame(results)

    def generate_excel_writer(self, df, output_writer_or_path, top_n=5):
        """
        Generates formatted Excel sheets for CPR Scanner.
        Accepts either file path string or ExcelWriter / BytesIO.
        """
        if isinstance(output_writer_or_path, str):
            writer = pd.ExcelWriter(output_writer_or_path, engine='openpyxl')
            should_close = True
        else:
            writer = output_writer_or_path
            should_close = False

        cols = list(df.columns)
        priority = ['Symbol', 'Expiry', 'ATM_Strike', 'Option_Type', 'Spot_Close']
        final_cols = priority + [c for c in cols if c not in priority]
        df_final = df[final_cols]

        cols_to_show = ['Symbol', 'Spot_Close', 'ATM_Strike']

        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # Dark Navy Blue for CPR
        center = Alignment(horizontal='center', vertical='center')

        def build_side_by_side_sheet(filter_col, sheet_title, ce_header_text, pe_header_text):
            df_filtered = df[df[filter_col] == True].copy()
            df_ce = df_filtered[df_filtered['Option_Type'] == 'CE'][cols_to_show].copy().reset_index(drop=True)
            df_pe = df_filtered[df_filtered['Option_Type'] == 'PE'][cols_to_show].copy().reset_index(drop=True)
            
            df_combined = pd.concat([df_ce, df_pe], axis=1)
            df_combined.to_excel(writer, sheet_name=sheet_title, index=False, startrow=1)
            
            worksheet = writer.sheets[sheet_title]
            
            # Merged headers
            worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
            c_ce = worksheet.cell(row=1, column=1)
            c_ce.value = ce_header_text
            c_ce.font = header_font
            c_ce.fill = header_fill
            c_ce.alignment = center
            
            worksheet.merge_cells(start_row=1, start_column=4, end_row=1, end_column=6)
            c_pe = worksheet.cell(row=1, column=4)
            c_pe.value = pe_header_text
            c_pe.font = header_font
            c_pe.fill = header_fill
            c_pe.alignment = center

            for i, col in enumerate(df_combined.columns):
                col_idx = i + 1
                series = df_combined.iloc[:, i]
                max_len = max((series.apply(str).map(len).max() if not series.empty else 0), len(str(col))) + 2
                column_letter = worksheet.cell(row=2, column=col_idx).column_letter
                worksheet.column_dimensions[column_letter].width = max_len

        # Sheet 1: Main Data
        df_final.to_excel(writer, sheet_name='Main Data', index=False)

        # Sheet 2: Narrow CPR
        build_side_by_side_sheet('Is_Narrow_CPR', 'Narrow CPR', 'Narrow CPR CE', 'Narrow CPR PE')

        # Sheet 3: Wide CPR
        build_side_by_side_sheet('Is_Wide_CPR', 'Wide CPR', 'Wide CPR CE', 'Wide CPR PE')

        # Sheet 4: Inside CPR
        build_side_by_side_sheet('Is_Inside_CPR', 'Inside CPR', 'Inside CPR CE', 'Inside CPR PE')

        # Sheet 5: Outside CPR
        build_side_by_side_sheet('Is_Outside_CPR', 'Outside CPR', 'Outside CPR CE', 'Outside CPR PE')

        # Sheet 6: Higher Value CPR
        build_side_by_side_sheet('Is_Higher_Value_CPR', 'Higher Value CPR', 'Higher Value CPR CE', 'Higher Value CPR PE')

        # Sheet 7: Lower Value CPR
        build_side_by_side_sheet('Is_Lower_Value_CPR', 'Lower Value CPR', 'Lower Value CPR CE', 'Lower Value CPR PE')

        # Sheet 7: Top N Output
        metrics = [
            ('OpnIntrst', f'Top {top_n} Open Interest'),
            ('ChngInOpnIntrst', f'Top {top_n} Change in OI'),
            ('TtlTradgVol', f'Top {top_n} Volume'),
            ('TtlNbOfTxsExctd', f'Top {top_n} Transactions')
        ]
        
        sheet_name_top = f'Top {top_n} Output'
        workbook = writer.book
        workbook.create_sheet(sheet_name_top)
        worksheet_top = workbook[sheet_name_top]
        
        start_row = 1
        current_col = 0
        
        for metric, title in metrics:
            df[metric] = pd.to_numeric(df[metric], errors='coerce').fillna(0)
            top_df = df.sort_values(by=metric, ascending=False).head(top_n).copy()
            cols_top = ['Symbol', 'Option_Type', 'ATM_Strike', 'Spot_Close', metric]
            top_display = top_df[cols_top].copy()

            top_display.to_excel(writer, sheet_name=sheet_name_top, index=False, startrow=start_row, startcol=current_col)
            
            op_col_start = current_col + 1
            op_col_end = current_col + len(cols_top)
            
            worksheet_top.merge_cells(start_row=1, start_column=op_col_start, end_row=1, end_column=op_col_end)
            cell_title = worksheet_top.cell(row=1, column=op_col_start)
            cell_title.value = title
            cell_title.font = header_font
            cell_title.fill = header_fill
            cell_title.alignment = center
            
            for i, col in enumerate(top_display.columns):
                col_idx = op_col_start + i
                series = top_display.iloc[:, i]
                max_len = max((series.apply(str).map(len).max() if not series.empty else 0), len(str(col))) + 2
                column_letter = worksheet_top.cell(row=2, column=col_idx).column_letter
                worksheet_top.column_dimensions[column_letter].width = max_len

            current_col += len(cols_top) + 1

        if should_close:
            writer.close()

if __name__ == "__main__":
    import glob
    files = sorted(glob.glob("BhavCopy*.zip"))
    if len(files) >= 2:
        scanner = CPRScanner()
        df = scanner.process_data(files[-1], files[-2])
        if df is not None:
            os.makedirs("CPR EOD SCANNER", exist_ok=True)
            output_file = os.path.join("CPR EOD SCANNER", "CPR_Test_Output.xlsx")
            scanner.generate_excel_writer(df, output_file)
            print(f"Saved CPR Test Output to {output_file}")

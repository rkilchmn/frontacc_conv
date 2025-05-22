"""
Financial data conversion utilities module.
Supports converting between various financial data formats including:
- Excel GL transactions to QIF format
"""

import pandas as pd
import argparse

class ConversionError(Exception):
    """Custom exception for conversion errors"""
    pass

class FrontaccConverter:
    """Class handling various financial data format conversions"""
    
    @staticmethod
    def gl2qif(excel_path: str, qif_path: str, payee: str, account_type: str = "Bank",) -> None:
        """
        Convert Excel GL transactions to QIF format
        
        Args:
            excel_path: Path to source Excel file
            qif_path: Path to output QIF file
            payee: Payee for the transaction
            account_type: QIF account type (default: Cash)
            
        Raises:
            ConversionError: If conversion fails
            FileNotFoundError: If input file doesn't exist
        """
        try:
            # Read the Excel file to find the 'Type' header row
            xls = pd.ExcelFile(excel_path)
            df_headers = pd.read_excel(xls, header=None)
            
            # Find the row index where the first column contains 'Type'
            type_row_idx = df_headers[df_headers[0].astype(str).str.strip().str.upper() == 'TYPE'].index[0]
            ROW_OPENING_BALANCE = type_row_idx + 1  # Row with opening balance is right after the header
            
            # Read period and opening balance using the dynamic row number
            period = pd.read_excel(xls, usecols="B", skiprows=2, nrows=1).iloc[0, 0]
            
            opening_balance_debit = pd.read_excel(xls, usecols="H", skiprows=ROW_OPENING_BALANCE, nrows=1).iloc[0, 0]
            opening_balance_credit = pd.read_excel(xls, usecols="I", skiprows=ROW_OPENING_BALANCE, nrows=1).iloc[0, 0]
            opening_balance = opening_balance_credit if pd.notna(opening_balance_credit) else - opening_balance_debit
            
            last_valid_index = 0
            # Read transactions starting from row 9
            df = pd.read_excel(
                xls, 
                skiprows=7, 
                usecols="A:J", 
                names=["Type", "Ref", "#", "Date", "Dimension", "Unused", "Person/Item", "Debit", "Credit", "Balance"],
                dtype={
                    "Type": str, 
                    "Ref": str, 
                    "#": str, 
                    "Date": "datetime64[ns]", 
                    "Dimension": str, 
                    "Unused": str, 
                    "Person/Item": str, 
                },
                converters={
                    "Debit": lambda x: float(str(x).replace(',', '').replace('.', '.', 1)) if pd.notnull(x) else 0.0,
                    "Credit": lambda x: float(str(x).replace(',', '').replace('.', '.', 1)) if pd.notnull(x) else 0.0,
                    "Balance": lambda x: float(str(x).replace(',', '').replace('.', '.', 1)) if pd.notnull(x) else 0.0
                }
            )
            # Start QIF file with account type header
            with open(qif_path, 'w') as qif_file:
                qif_file.write(f"!Type:{account_type}\n")
                
                # Iterate over each row in the dataframe and format into QIF
                for _, row in df.iterrows():
                    if pd.isna(row['Date']):  # Stop processing if 'Date' is NaN
                        empty_line_index = _
                        break
                    trans_amount = - row['Debit'] if pd.notna(row['Debit']) else row['Credit']
                    qif_file.write(f"D{row['Date'].strftime('%m/%d/%Y')}\n")
                    qif_file.write(f"T{trans_amount}\n")
                    qif_file.write(f"N{row['Ref']}\n")
                    qif_file.write(f"M{row['Type']}: {row['Person/Item']}\n")
                    qif_file.write(f"P{payee}\n")
                    qif_file.write("^\n")

                # Read closing balance similar to opening balance
                closing_balance_debit = pd.read_excel(xls, usecols="H", skiprows=6 + empty_line_index, nrows=1).iloc[0, 0]
                closing_balance_credit = pd.read_excel(xls, usecols="I", skiprows=6 + empty_line_index, nrows=1).iloc[0, 0]
                closing_balance = closing_balance_credit if pd.notna(closing_balance_credit) else - closing_balance_debit

            print(f"Successfully converted {excel_path} to {qif_path}")
            print(f"Period: {period}, Opening Balance: {opening_balance}, Closing Balance: {closing_balance}")
            
        except Exception as e:
            raise ConversionError(f"Failed to convert {excel_path} to QIF: {str(e)}")

def main(): 
    """CLI entry point"""
    parser = argparse.ArgumentParser(description='Convert financial data formats.')
    parser.add_argument('conversion_type', type=str, choices=['gl2qif'], help='Type of conversion to perform')
    parser.add_argument('input_file', type=str, help='Path to the input Excel file')
    parser.add_argument('output_file', type=str, help='Path to the output QIF file')
    parser.add_argument('payee', type=str, help='Payee for the transactions')
    
    args = parser.parse_args()
    
    if args.conversion_type == 'gl2qif':
        FrontaccConverter.gl2qif(args.input_file, args.output_file, args.payee)  # Pass payee to gl2qif method

if __name__ == "__main__":
    main()

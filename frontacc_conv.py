"""
Financial data conversion utilities module.
Supports converting between various financial data formats including:
- Excel GL transactions to QIF format
"""

import pandas as pd
import argparse
import csv
import os  # Add this import at the top of the file
import re  # Ensure this import is at the top of the file
from decimal import Decimal

class ConversionError(Exception):
    """Custom exception for conversion errors"""
    pass

class FrontaccConverter:
    """Class handling various financial data format conversions"""
    
    @staticmethod
    def gl2qif(input_file: str, output_file: str, account_type: str) -> None:
        """
        Convert Excel GL transactions to QIF format
        
        Args:
            input_file: Path to source Excel file
            output_file: Path to output QIF file
            payee: Payee for the transaction
            account_type: QIF account type (default: Cash)
            
        Raises:
            ConversionError: If conversion fails
            FileNotFoundError: If input file doesn't exist
        """
        try:
            # Read specific cells for period and opening balance
            xls = pd.ExcelFile(input_file)
            period = pd.read_excel(xls, usecols="B", skiprows=2, nrows=1).iloc[0, 0]

            # convention from frontaccounting GL report for bank accounts: credits are -, debits are +
            opening_balance_debit  = pd.read_excel(xls, usecols="I", skiprows=5, nrows=1).iloc[0, 0]
            opening_balance_credit = pd.read_excel(xls, usecols="J", skiprows=5, nrows=1).iloc[0, 0]
            opening_balance = - Decimal(f"{float(opening_balance_credit):.2f}") if pd.notna(opening_balance_credit) else Decimal(f"{float(opening_balance_debit):.2f}")
            
            last_valid_index = 0
            # Initialize running total for balance
            running_balance = opening_balance
            
            # Read transactions starting from row 9
            df = pd.read_excel(
                xls, 
                skiprows=7, 
                usecols="A:K", 
                names=["Type", "Ref", "#", "Date", "Dimension", "Unused", "Person/Item", "Memo", "Debit", "Credit", "Balance"],
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
                    "Debit": lambda x: Decimal(str(float(str(x).replace(',', '').replace('.', '.', 1)))) if pd.notnull(x) else Decimal('0.00'),
                    "Credit": lambda x: Decimal(str(float(str(x).replace(',', '').replace('.', '.', 1)))) if pd.notnull(x) else Decimal('0.00'),
                    "Balance": lambda x: Decimal(str(float(str(x).replace(',', '').replace('.', '.', 1)))) if pd.notnull(x) else Decimal('0.00')
                }
            )
            # Start QIF file with account type header
            with open(output_file, 'w') as qif_file:
                qif_file.write(f"!Type:{account_type}\n")
                
                # Iterate over each row in the dataframe and format into QIF
                for _, row in df.iterrows():
                    if pd.isna(row['Date']):  # Stop processing if 'Date' is NaN
                        empty_line_index = _
                        break

                    trans_amount = -row['Credit'] if pd.notna(row['Credit']) else row['Debit']
                    running_balance += trans_amount  # Update running balance
                    qif_file.write(f"D{row['Date'].strftime('%m/%d/%Y')}\n")
                    qif_file.write(f"T{trans_amount}\n")  # This will now be a Decimal with 2 digits
                    qif_file.write(f"N{row['Ref']}\n")
                    descr = row['Memo'].replace("\n", "")
                    qif_file.write(f"M{row['Type']}: {descr}\n")
                    # extract payee info
                    if isinstance(row['Person/Item'], str) and row['Person/Item']:
                        payee_match = re.search(r'(?<=\] )([^/|\n]+)', row['Person/Item'])
                        payee = payee_match.group(1).strip() if payee_match else None
                        if payee:  
                            qif_file.write(f"P{payee}\n")
                    # end record
                    qif_file.write("^\n")

                # Read closing balance similar to opening balance
                closing_balance_debit = pd.read_excel(xls, usecols="I", skiprows=8 + empty_line_index, nrows=1).iloc[0, 0]
                closing_balance_credit = pd.read_excel(xls, usecols="J", skiprows=8 + empty_line_index, nrows=1).iloc[0, 0]
                closing_balance = - Decimal(f"{float(closing_balance_credit):.2f}") if pd.notna(closing_balance_credit) else Decimal(f"{float(closing_balance_debit):.2f}")

                # Compare final running balance with closing balance
                if running_balance != closing_balance:
                    raise ValueError(f"Final calculated balance {running_balance} does not match closing balance {closing_balance}")

            print(f"Successfully converted {input_file} to {output_file}")
            print(f"Period: {period}, Opening Balance: {opening_balance}, Closing Balance: {closing_balance}")
            
        except Exception as e:
            print(f"Error details: {str(e)}")  # Print the error details for debugging
            raise ConversionError(f"Failed to convert {input_file} to QIF: {str(e)}")

def main(): 
    """CLI entry point"""
    parser = argparse.ArgumentParser(description='Convert financial data formats.')
    parser.add_argument('conversion_type', type=str, choices=['gl2qif'], help='Type of conversion to perform')
    parser.add_argument('input_file', type=str, help='Path to the input file')
    parser.add_argument('output_file', type=str, help='Path to the output file')
    parser.add_argument('account_type', type=str, nargs='?', default='Bank', help='Account Type (default: "BANK")')
     
    args = parser.parse_args()
    
    if args.conversion_type == 'gl2qif':
        FrontaccConverter.gl2qif(args.input_file, args.output_file, args.account_type) 
    
if __name__ == "__main__":
    main()

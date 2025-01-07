
import pandas as pd
import re
from PyPDF2 import PdfReader

def extractData():
    try:
        # Load the Excel file provided by the user
        input_path = 'inputs/Vendor_Information_Input.xlsx'
        df = pd.read_excel(input_path)

        # Create new columns 'Vendor Name' and 'Vendor Number' derived from 'Partner Name'
        df[['Vendor Name', 'Vendor Number']] = df['Partner Name'].str.split('-', expand=True)

        # Display the updated DataFrame
        df.head()

        # Save the DataFrame to an Excel file
        output_path = 'output/Vendor_Information_Output.xlsx'
        df.to_excel(output_path, index=False)

        print(f"Data has been saved to {output_path}")
    except Exception as e:
        print(f"An error occurred: {e}")



if __name__ == "__main__":
    print("Hello World")
    extractData()
import tabula
import pandas as pd

def convert_pdf_to_excel(pdf_path, excel_path):
    # Read the PDF file
    temp_df=pd.DataFrame()
    for i in range(1,980):
        df = tabula.read_pdf(pdf_path, pages=f'{i}')
        temp_df = pd.concat([temp_df,df[0]],ignore_index=True)
        print(i)
    # Convert the extracted data to Excel
    temp_df.to_excel(excel_path, index=False)

if __name__ == "__main__":
    # Specify the path to your PDF file and the desired Excel file
    pdf_file_path = "1457bdd4.pdf"
    excel_file_path = "new1.xlsx"

    # Convert PDF to Excel
    convert_pdf_to_excel(pdf_file_path, excel_file_path)

    print(f"Conversion completed. Excel file saved at: {excel_file_path}") 
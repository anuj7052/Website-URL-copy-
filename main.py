import pandas as pd
from googlesearch import search
import openpyxl

def fetch_company_websites(file_path):
    # Load the Excel file
    df = pd.read_excel(file_path, sheet_name='Sheet1')
    
    # Ensure Column B exists and add a placeholder for Column F
    if 'Company Name' not in df.columns:
        raise ValueError("Column B ('Company Name') is missing.")
    if 'Website' not in df.columns:
        df['Website'] = None  # Add a new column for storing websites

    # Loop through each company name in Column B
    for index, company_name in enumerate(df['Company Name']):
        if pd.notnull(company_name):  # Check if the company name is not empty
            print(f"Searching for: {company_name}")
            try:
                # Perform Google Search to get the company website
                query = f"{company_name} official site"
                # Extract the first result from the search generator
                website = next(search(query, num_results=1), None)
                df.at[index, 'Website'] = website  # Save the website URL in Column F
            except Exception as e:
                print(f"Error fetching website for {company_name}: {e}")
                df.at[index, 'Website'] = "Error"

    # Save the updated DataFrame back to Excel
    output_file = "updated_" + file_path
    df.to_excel(output_file, index=False)
    print(f"Updated file saved as: {output_file}")

# Call the function
fetch_company_websites("FCSPL_Dec Data Website.xlsx")

import pdfplumber
import pandas as pd

# Function to extract table data from PDF using pdfplumber
def extract_table_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        all_tables = []
        
        # Loop through all pages and extract tables
        for page in pdf.pages:
            # Extract tables on the current page
            table = page.extract_table()
            if table:
                all_tables.append(table)
        
        return all_tables

# Function to process the extracted table data into a structured DataFrame
def process_table_to_dataframe(tables):
    # Initialize an empty list to hold the rows of data
    rows = []
    
    # Initialize a variable to hold the current dimension group
    current_dim = None

    # Loop through all extracted tables
    for table in tables:
        header = table[0]  # Get the first row which is the header
        
        # Check if the table header contains dimension info
        if "Dim" in header[0]:  # If it's a dimension-related table
            for row in table[1:]:  # Start processing from the second row onwards
                if row[0]:  # If there's a dimension identifier
                    current_dim = row[0]  # Set the current dimension (e.g., "1B1")
                    
                    # Create a new row with the current dimension and the rest of the values
                    new_row = [current_dim] + row[1:]  # Merge dimension with other columns
                    rows.append(new_row)  # Append to the rows list

        elif "M" in header[0]:  # If this table contains size data (M, L, XL)
            size_data = table[1:]  # Extract size-related rows
            
            # For each size data, align it under the proper columns
            for i, row in enumerate(size_data):
                # Update the appropriate row with size data
                if len(rows) > i:
                    rows[i] = rows[i] + row  # Merge the size columns (M, L, XL)
    
    # Dynamically set column headers based on the number of columns in the data
    num_columns = len(rows[0]) if rows else 0
    columns = ["Dim", "Description", "Comment", "Tol (-)", "Tol (+)", "XS", "", "", "S", "", "", "M", "", "", "L", "", "", "XL", "", ""]
    
    # Ensure that the number of columns matches the data
    while len(columns) < num_columns:
        columns.append(f"Extra Column {len(columns) - 10 + 1}")
    
    # Create a DataFrame from the rows
    final_df = pd.DataFrame(rows, columns=columns)
    return final_df

# Function to save the DataFrame to a CSV file
def save_to_csv(df, output_filename):
    df.to_csv(output_filename, index=False)

# Main function to extract data from PDF, process it, and save to CSV
def convert_pdf_to_csv(pdf_path, output_filename):
    # Step 1: Extract table data from PDF
    tables = extract_table_from_pdf(pdf_path)
    
    # Step 2: Process the extracted table into a structured DataFrame
    df = process_table_to_dataframe(tables)
    
    # Step 3: Save the DataFrame to CSV
    if not df.empty:
        save_to_csv(df, output_filename)
        print(f"Data has been successfully saved to {output_filename}")
    else:
        print("No valid table found in the PDF.")

# Path to the PDF file
pdf_path = "input_doc.pdf"  # Replace with the path to your PDF file

# Output CSV file name
output_filename = "output_merged.csv"  # Replace with the desired CSV file name

# Run the conversion process
convert_pdf_to_csv(pdf_path, output_filename)

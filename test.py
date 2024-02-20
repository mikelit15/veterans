import openpyxl

def fix_hyperlinks_in_excel(file_path, column_letter):
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(file_path)
    # Assuming the hyperlinks are in the first sheet, or adjust as necessary
    sheet = workbook.active
    
    # Iterate through each row in the specified column
    for row in sheet.iter_rows(min_col=openpyxl.utils.column_index_from_string(column_letter),
                               max_col=openpyxl.utils.column_index_from_string(column_letter),
                               min_row=884,  # Assuming the first row is headers
                               max_row=1154):
        cell = row[0]  # Since we're iterating by columns, each 'row' is actually a single cell
        if cell.hyperlink:  # Check if the cell has a hyperlink
            url = cell.hyperlink.target
            # Check and fix the incorrect pattern in the URL
            new_url = url.replace('EvergreenDD', 'EvergreenD', 1)
            if new_url != url:  # If a change was made
                cell.hyperlink.target = new_url
                print(f"Fixed hyperlink: {url} to {new_url}")

    # Save the workbook (consider saving to a new file during testing to avoid data loss)
    workbook.save(file_path)
    print("Hyperlinks fixed and workbook saved.")

# Replace with the actual Excel file path and the column containing the hyperlinks
excel_file_path = r"\\ucclerk\pgmdoc\Veterans\Veterans.xlsx"
column_with_hyperlinks = 'O'  # Adjust the column as needed
fix_hyperlinks_in_excel(excel_file_path, column_with_hyperlinks)
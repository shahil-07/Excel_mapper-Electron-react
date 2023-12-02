import xlwings as xw
import re
import time
import sys
import os

def main():
    try:
        def open_open_primary_secondary_excels(file1, file2):
            app = xw.App(visible=False)
            # Load the first Excel file
            wb1 = app.books.open(file1) 

            # Load the second Excel file
            wb2 = app.books.open(file2) 

            return [app, wb1, wb2]

        [app, wb1, wb2] = open_open_primary_secondary_excels(sys.argv[1],sys.argv[2])

        def save_and_close_excel(app, wb1, wb2, file1): 
            input_file_name = os.path.splitext(os.path.basename(file1))[0]   
            # Specify the document directory
            document_dir = os.path.expanduser("~\\Documents")

            # Create the directory if it doesn't exist
            excel_mapper_dir = os.path.join(document_dir, "ExcelMapper")
            if not os.path.exists(excel_mapper_dir):
                os.makedirs(excel_mapper_dir)

            #Clear the fiels inside the Excel mapper folder
            for filename in os.listdir(excel_mapper_dir):
                file_path = os.path.join(excel_mapper_dir, filename)
                if os.path.isfile(file_path):
                    os.remove(file_path)

            # Build the full path for the output file in the ExcelMapper directory
            output_file_path = os.path.join(excel_mapper_dir, f"{input_file_name}.xlsx")

            # Save the modified first Excel file to the document directory
            wb1.save(output_file_path)
            wb1.close()
            wb2.close()
            # Close the Excel application
            app.quit()
            # print("Output file path:", output_file_path) 
            return output_file_path
            

        def open_sheet_with_name(ws1_name, ws2_name):
            ws1 = wb1.sheets[ws1_name]
            ws2 = wb2.sheets[ws2_name]
            return [ws1, ws2]

        def cell_mapping_Dimensionen(file1, file2):
            # [ws1, ws2] = open_sheet_with_name('Dimensionen1', 'Sheet1')
            # Determine which sheet name to use based on a condition
            ws1_name = None
            # print("Sheet names in file1:")
            for sheet in wb1.sheets:
                # print(sheet.name)
                if re.search(r'Dimensionen1', sheet.name, re.IGNORECASE):
                    ws1_name = 'Dimensionen1'
                    break
                elif re.search(r'Dimension1', sheet.name, re.IGNORECASE):
                    ws1_name = 'Dimension1'
                    break
                elif re.search(r'Dimensions1', sheet.name, re.IGNORECASE):
                    ws1_name = 'Dimensions1'
                    break
            if ws1_name is None:
                print(f"Sheet name not found in file1: {file1}")
                return
            # print(f"Selected ws1_name: {ws1_name}")

            [ws1, ws2] = open_sheet_with_name(ws1_name, 'Sheet1')
            # for row in ws2.range('A1').expand('table').value:
            #     print(row)

            # Define a list of valid starting values
            header_values = ['No.', 'Nr.', 'Pos', 'Ref. No.']
            header_row = None
            
            # Loop through all cells in column A of ws1
            for cell in ws1.range('A1:A40'):
                # Check if the cell value is a string and starts with a valid starting value
                if isinstance(cell.value, str) and any(cell.value.startswith(prefix) for prefix in header_values):
                    # Get the entire row where a valid starting value is present
                    if cell.api.MergeArea:
                        header_row = cell.api.MergeArea.Rows.Count + cell.row - 1
                    else:
                        header_row = cell.row
                    break

            # If a valid header row was found, clear only the cell values in rows from  1 to 40
            if header_row is not None:
                for row_num in range(header_row + 1, header_row + 41):
                    ws1.range(f'A{row_num}:Z{row_num}').value = [''] * len(ws1.range(f'A{row_num}:Z{row_num}').value)
                # print("header row:", header_row)
            else:
                print("Header not found.")

            match = re.match(r'(\D+)(\d+)', ws1_name)

            if match:
                # Extract the two parts
                First_name = match.group(1)
                Layer = int(match.group(2))
                # print(f"First Part: {First_name}")
                # print(f"Second Part: {Layer}")
            else:
                return
            

            # Check if ws2 has more than 40 rows
            # ws2_data = ws2.range('A:A').value
            # non_empty_rows = [row for row in ws2_data if row is not None and any(cell.strip() for cell in row)]
        
            font_name = 'Arial'
            font_size = 10

            ws1_current = ws1      #track the current sheet
            num_new_sheets_created = 0    #number of new sheets created 
            ws2_row_num = 0      #track the row number in ws2
            current_sheet_row_count = 0    #count the number of rows in the current sheet
            next_dimension_sheet = None

            # Copy all rows from ws2 and paste them in ws1 starting from header_row + 1
            used_range_ws2 = ws2.used_range
            for row_num, row in enumerate(used_range_ws2.rows, start=1):
                ws2_row_num += 1  # Increment the row number in ws2

                if current_sheet_row_count == 0:
                    # Clear the content of the current sheet from header_row + 1 to header_row + 41
                    for clear_row in range(header_row + 1, header_row + 41):
                        ws1_current.range(f'A{clear_row}:Z{clear_row}').value = [''] * 26

                # Calculate the destination row in the current sheet
                destination_row = header_row + 1 + current_sheet_row_count
                # print(f"ws2_row_num: {ws2_row_num}, destination_row: {destination_row}")

                data_from_file2 = [cell.value for cell in row]
                # print("data_from_file2", data_from_file2)
                destination_range = ws1_current.range(f'A{destination_row}:Z{destination_row}')
                # print("destination_range.value", destination_range.value)
                merged_cell_names = []

                # Initialize a counter for merged cells
                merged_cell_count = 0

                # Loop through cells in the destination range and check for merged cells
                for row in destination_range.rows:
                    for cell in row:
                        if cell.api.MergeArea.Cells.Count > 1:
                            # Add the address (name) of the merged cell to the list
                            merged_cell_names.append(cell.api.MergeArea.Address)
                            if merged_cell_count < len(data_from_file2):
                                # Check if the merged cell has already been filled
                                if cell.api.MergeArea.Address not in merged_cell_names[:-1]:
                                    cell.value = data_from_file2[merged_cell_count]
                                    merged_cell_count += 1
                                    cell.api.Font.Name = font_name
                                    cell.api.Font.Size = font_size
                        else:
                            # If it's not a merged cell, add data directly
                            if merged_cell_count < len(data_from_file2):
                                cell.value = data_from_file2[merged_cell_count]
                                merged_cell_count += 1
                                cell.api.Font.Name = font_name
                                cell.api.Font.Size = font_size
                # print(f"Merged Cell Names: {merged_cell_names}")
                # print(f"Number of Merged Cells: {merged_cell_count}"
                
                # destination_range.value = data_from_file2

                current_sheet_row_count += 1

                if current_sheet_row_count >= 40:
                    # Check if the next_dimension_sheet already exists
                    Layer += 1
                    next_sheet_name = f'{First_name}{Layer}'
                    existing_next_sheet = None
                    for sheet in ws1.book.sheets:
                        if sheet.name == next_sheet_name:
                            existing_next_sheet = sheet
                            break

                    if existing_next_sheet:
                        # Update the existing sheet
                        ws1_current = existing_next_sheet
                    else:
                        # Create a new sheet after every 40 rows
                        if next_dimension_sheet:
                            ws1.api.Copy(After=next_dimension_sheet.api)
                        else:
                            ws1.api.Copy(After=ws1.api)
                        ws1_current = ws1.book.sheets.active
                        next_dimension_sheet = ws1_current
                        Layer += 1
                        new_sheet_name = f'{First_name}{Layer}'
                        ws1_current.name = new_sheet_name
                        num_new_sheets_created += 1

                    # Reset the row count for the new sheet
                    current_sheet_row_count = 0

                # Check if any cell value is a decimal number
                for col_num, cell in enumerate(destination_range.columns, start=1):
                    if col_num != 1 and isinstance(cell.value, (float, int)) and isinstance(cell.value, float) and '.' in str(cell.value):
                        # Format the cell to display exactly three decimal places
                        cell.number_format = '0.000'

            # print(f"Total new sheets created: {num_new_sheets_created}")
                

        cell_mapping_Dimensionen(sys.argv[1],sys.argv[2])                   

        # save after updation 
        save_and_close_excel(app, wb1, wb2, sys.argv[1])
        print("DONE", flush=True)
        sys.stdout.flush()
    
    except Exception as e:
        sys.stderr.write("An error occurred: " + str(e) + "\n")
        if 'app' in locals():
            app.quit()
        # Send a status code of -1 to Electron indicating an error
        sys.stderr.write("-1\n")
        sys.stderr.flush()

if __name__ == "__main__":
    main()
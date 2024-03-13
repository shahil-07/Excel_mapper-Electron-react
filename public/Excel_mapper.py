import xlwings as xw
import re
import time
import sys
import os

# python .\siemens.py 
# "C:\Users\Shahil\Desktop\app\public\uploads\template\Siemens_Amberg_Rev-16_F01.xlsm" 
# "C:\Users\Shahil\Desktop\app\public\uploads\script-data\empb_data.xlsx" 
# "C:\Users\Shahil\Desktop\project 4\cell_mapping_siemens.xlsx"
# -arg 0 sccript name -arg1 is template loc -arg2 is scriptdata loc -arg3 is cellmapping 
def main():
    # try:
        app = xw.App(visible=False)
        app.display_alerts = False
        app.screen_updating = False
        Job_Name = None
        cell_mapping_for_Laufzettel = {}
        cell_mapping_for_Urwerte = {}
        cell_mapping_for_Prufplan = {}
        cell_mapping_for_Schliffbilder = {}
        cell_mapping_for_SchliffbilderViaHolefilling = {}
        cell_mapping_for_MasterData = {}
        cell_mapping_for_ManualData = {}

        def populate_cell_mappings():
            # print('Args', sys.argv)
            file_name = sys.argv[3]
            wb1 = app.books.open(file_name)

            try:
                Laufzettel_sheet = wb1.sheets["Laufzettel"]
                for row_number, cell in enumerate(Laufzettel_sheet.range('A:A'), start=2):
                    key = Laufzettel_sheet.range(f"A{row_number}").value
                    if row_number > 50:
                        break
                    if key:
                        cell_mapping_for_Laufzettel[key] = (Laufzettel_sheet.range(f"B{row_number}").value,Laufzettel_sheet.range(f"C{row_number}").value)
                # print(cell_mapping_for_Laufzettel)
            except Exception as e:
                print(f"An error occurred while processing the 'Laufzettel' sheet: {e}")      
                
            try:
                Urwerte_sheet = wb1.sheets["Urwerte"]
                for row_number, cell in enumerate(Urwerte_sheet.range('A:A'), start=2):
                    key = Urwerte_sheet.range(f"A{row_number}").value
                    if row_number > 50:
                        break
                    # print(cell.value,Urwerte_sheet.range(f"B{row_number}").value,Urwerte_sheet.range(f"C{row_number}").value)
                    if key:
                        cell_mapping_for_Urwerte[key] = (Urwerte_sheet.range(f"B{row_number}").value,Urwerte_sheet.range(f"C{row_number}").value)
                # print(cell_mapping_for_Urwerte)
            except Exception as e:
                print(f"An error occurred while processing the 'Urwerte' sheet: {e}")   

            try:
                Prufplan_sheet = wb1.sheets["Prufplan"]
                for row_number, cell in enumerate(Prufplan_sheet.range('A:A'), start=2):
                    key = Prufplan_sheet.range(f"A{row_number}").value
                    if row_number > 50:
                        break
                    if key:
                        cell_mapping_for_Prufplan[key] = (Prufplan_sheet.range(f"B{row_number}").value,Prufplan_sheet.range(f"C{row_number}").value)
                # print(cell_mapping_for_Prufplan)
            except Exception as e:
                print(f"An error occurred while processing the 'Prufplan' sheet: {e}")    

            try:
                Schliffbilder_sheet = wb1.sheets["Schliffbilder"]
                for row_number, cell in enumerate(Schliffbilder_sheet.range('A:A'), start=2):
                    key = Schliffbilder_sheet.range(f"A{row_number}").value
                    if row_number > 50:
                        break
                    if key:
                        cell_mapping_for_Schliffbilder[key] = (Schliffbilder_sheet.range(f"B{row_number}").value,Schliffbilder_sheet.range(f"C{row_number}").value)
                # print(cell_mapping_for_Schliffbilder)
            except Exception as e:
                print(f"An error occurred while processing the 'Schliffbilder' sheet: {e}")         
            
            try:
                SchliffbilderViaHolefilling_sheet = wb1.sheets["Schliffbilder Via Hole filling"]
                for row_number, cell in enumerate(SchliffbilderViaHolefilling_sheet.range('A:A'), start=2):
                    key = SchliffbilderViaHolefilling_sheet.range(f"A{row_number}").value
                    if row_number > 50:
                        break
                    if key:
                        cell_mapping_for_SchliffbilderViaHolefilling[key] = (SchliffbilderViaHolefilling_sheet.range(f"B{row_number}").value,SchliffbilderViaHolefilling_sheet.range(f"C{row_number}").value)
                # print(cell_mapping_for_SchliffbilderViaHolefilling)
            except Exception as e:
                print(f"An error occurred while processing the 'Schliffbilder Via Hole filling' sheet: {e}")   

            try:
                MasterData_sheet = wb1.sheets["MasterData"]
                for row_number, cell in enumerate(MasterData_sheet.range('A:A'), start=2):
                    key = MasterData_sheet.range(f"A{row_number}").value
                    if row_number > 80:
                        break
                    if key:
                        cell_mapping_for_MasterData[key] = (MasterData_sheet.range(f"B{row_number}").value,MasterData_sheet.range(f"C{row_number}").value,MasterData_sheet.range(f"D{row_number}").value)
                # print(cell_mapping_for_MasterData)
            except Exception as e:
                print(f"An error occurred while processing the 'Master Data' sheet: {e}")  

            try:
                ManualData_sheet = wb1.sheets["ManualData"]
                for row_number, cell in enumerate(ManualData_sheet.range('A:A'), start=2):
                    key = ManualData_sheet.range(f"A{row_number}").value
                    if row_number > 80:
                        break
                    if key:
                        cell_mapping_for_ManualData[key] = (ManualData_sheet.range(f"B{row_number}").value,ManualData_sheet.range(f"C{row_number}").value,ManualData_sheet.range(f"D{row_number}").value)
                # print(cell_mapping_for_ManualData)
            except Exception as e:
                print(f"An error occurred while processing the 'Manual Data' sheet: {e}")  
            wb1.close()
        populate_cell_mappings() 

        def open_open_primary_secondary_excels(file1, file2):
            # Load the first Excel file
            wb1 = app.books.open(file1)

            # Load the second Excel file
            wb2 = app.books.open(file2)


            return [app, wb1, wb2]

        # open file descriptions
        [app, wb1, wb2] = open_open_primary_secondary_excels(sys.argv[1],sys.argv[2])


        def save_and_close_excel(app, wb1, wb2, new_filename):
            # Build the full path for the output file in the ExcelMapper directory
            output_file_path = os.path.join(sys.argv[4], sys.argv[5])

            # Save the modified first Excel file to the document directory
            wb1.save(output_file_path)
            wb1.close()
            wb2.close()
            # Close the Excel application
            app.quit()
            # print("Output file path:", output_file_path) 
            return output_file_path


        def open_sheet_with_name(ws1_name, ws2_name):
            ws1 = None
            ws2 = None

            for sheet in wb1.sheets:
                if sheet.name.strip() == ws1_name.strip():
                    ws1 = sheet
                    break

            for sheet in wb2.sheets:
                if sheet.name.strip() == ws2_name.strip():
                    ws2 = sheet
                    break
            return [ws1, ws2]

        def cell_mapping_Laufzettel(file1, file2):
            try:
                [ws1, ws2] = open_sheet_with_name('Laufzettel', 'Laufzettel')

                # Update the cells using the mappings
                for cell_ref, (source_cell, prefix) in cell_mapping_for_Laufzettel.items():
                    value = ws2.range(source_cell).value
                    ws1.range(cell_ref).value = value

                    # if source_cell == 'B5':
                    # # Assuming the checkbox value is either 'Yes' or 'No'
                    # # Modify the condition according to your actual checkbox values
                    #     if value == 'Yes':
                    #         try:
                    #             checkbox_shape = ws1.shapes
                    #             # checkbox_shape[18].text = "dummy"
                    #             # print(checkbox_shape[18].top)
                    #             # Fill the checkbox with a black color
                    #             checkbox_shape.fill.solid()
                    #             checkbox_shape.fill.fore_color.rgb = (0, 0, 0)  # Set color to black
                    #         except Exception as e:
                    #             print(f"Error occurred while changing checkbox color: {e}")
                    #     else:
                    #         try:
                    #             # If the checkbox value is 'No' or any other value, do not change the checkbox color
                    #             pass
                    #         except Exception as e:
                    #             print(f"Error occurred while changing checkbox color: {e}")
                    # else:
                    #     ws1.range(cell_ref).value = value
            except Exception as e:
                print(f"An error occurred while processing the 'Laufzettel' sheet: {e}")        

        def cell_mapping_Urwerte(file1, file2):
            try:
                [ws1, ws2] = open_sheet_with_name('Urwerte', 'Uwerte')
                # print(cell_mapping_Urwerte)
                
                # Update the cells using the mappings
                for cell_ref, (source_cell, prefix) in cell_mapping_for_Urwerte.items():
                    value = ws2.range(source_cell).value
                    if prefix is not None:
                        if value is None:
                            value = '---'
                        else:
                            value = f"{prefix} {value}"
                

                    if value is not None:    
                        if value.lower() == 'no':
                            ws1.range(cell_ref).value = '---'
                        elif value.lower() == 'yes':
                            ws1_range = ws1.range(cell_ref)
                            data_validation = ws1_range.api.Validation
                            if data_validation.Type == 3:  # Check if data validation is a dropdown list
                                formula_text = data_validation.Formula1
                                formula_range = ws1.range(formula_text)  # Convert formula to xlwings range
                                # Get the value from the second cell in the formula_range
                                second_cell_value = formula_range[1].value
                                # print(f"Value at second cell: {second_cell_value}")

                                ws1.range(cell_ref).value = second_cell_value
                                # Loop through each cell in the formula_range and print its value
                                # for row in formula_range.rows:
                                #     for cell in row:
                                #         cell_value = cell.value
                                #         print(f"Value at {cell.address}: {cell_value}")
                        else:    
                            ws1.range(cell_ref).value = value
                    else:
                        ws1.range(cell_ref).value = '---'
            except Exception as e:
                print(f"An error occurred while processing the 'Urwerte' sheet: {e}")            

        def cell_mapping_Prufplan(file1, file2):
            try:
                [ws1, ws2] = open_sheet_with_name('Prüfplan', 'Prufplan')
                

                for cell_ref, (source_cell, prefix) in cell_mapping_for_Prufplan.items():
                    if not cell_ref.startswith('A'):
                        source_value = ws2.range(source_cell).value
                        if source_value is not None:
                            if ',' in cell_ref:
                                cell_refs = cell_ref.split(',')
                                cell_refs = [ref.strip() for ref in cell_refs]
                                if len(cell_refs) == 2:
                                    source_values = source_value.split('x')
                                    source_values = [val.strip() for val in source_values]
                                    if len(source_values) == 2:
                                        ws1.range(cell_refs[0]).value = source_values[0]
                                        ws1.range(cell_refs[1]).value = source_values[1]
                            else:
                                ws1.range(cell_ref).value = source_value
                        else:
                            # Source value is empty
                            ws1_range = ws1.range(cell_ref)
                            try:
                                data_validation = ws1_range.api.Validation
                                if data_validation.Type == 3:  # corresponds to dropdown list
                                    formula_text = data_validation.Formula1
                                    formula_range = ws1.range(formula_text)  # Convert formula to xlwings range
                                    # Get the value from the second cell in the formula_range
                                    first_cell_value = formula_range[0].value
                                    # print(f"Value at first cell: {first_cell_value}")

                                    ws1.range(cell_ref).value = first_cell_value
                                    # ws1_range.value = 'kein UL'  # Set to 'kein UL' for dropdown cells
                                else:
                                    ws1_range.value = None  # Set to None for non-dropdown cells
                            except:
                                ws1_range.value = None  

                    #for hide and unhide                  
                    elif cell_ref.startswith('A'):
                        if source_cell is not None:
                            value = str(ws2.range(source_cell).value).strip()

                            # Check the value of the source cell (assuming the cell contains either 'yes' , 'no' or ' ')
                            if value.lower() == 'yes':
                                ws1.range(cell_ref).api.EntireRow.Hidden = False
                            else:
                                ws1.range(cell_ref).api.EntireRow.Hidden = True
                        else:
                            ws1.range(cell_ref).api.EntireRow.Hidden = True
            except Exception as e:
                print(f"An error occurred while processing the 'Prufplan' sheet: {e}")                


        def cell_mapping_Lagenaufbau(file1, file2):
            try:
                [ws1, ws2] = open_sheet_with_name('Lagenaufbau', 'Lagenaufbau')

                # Find the starting and ending rows based on the presence of green color in column A
                green_color = (204, 255, 204)  # RGB color of the green cells

                start_row = None
                end_row = None
                urwete_mapping_formula_outer_layer = None
                urwete_mapping_formula_inner_layer = None
                smt =None
                smb =None
                l01 = None
                l_end = None

                for row_number, cell in enumerate(ws1.range('A:A'), start=1):
                    if row_number > 100:
                        break

                    if cell.color == green_color:
                        if start_row is None:
                            start_row = row_number
                        else:
                            end_row = row_number
                    
                    if type (cell.value) == str and 'Kupfer Innenlage 2' in cell.value:
                        urwete_mapping_formula_inner_layer =  ws1.range(f"E{row_number}").formula 



                if start_row is None or end_row is None:
                    raise ValueError("Start or end row not found in the worksheet.")

                # Find Urwete mapping formula for outer and inner layers
                urwete_mapping_formula_outer_layer = ws1.range(f"E{start_row + 1}").formula 
                Lagenauf_formula_layer = ws1.range(f"F{start_row + 1}").formula
                # print(urwete_mapping_formula_outer_layer, urwete_mapping_formula_inner_layer) 

                for row_number, cell in enumerate(ws2.range('A:A'), start=1):
                    if row_number > 40:
                        break

                    if type (cell.value) == str:
                        if 'smt' in cell.value:
                            smt = row_number
                        elif 'smb' in cell.value:
                            smb = row_number
                        elif 'l01' in cell.value:
                            l01 = row_number
                        elif l01 is not None and cell.value.startswith('l') and cell.value[1:].isdigit():
                            l_end = row_number

                # Substitute the green row values
                # ws1.range(f"B{start_row}").value = ws2.range(f"B{smt}").value
                # ws1.range(f"B{end_row}").value = ws2.range(f"B{smb}").value

                
                # Delete cells in the rows between the last updated row and the end_row
                ws1.range(f"A{start_row + 1}:A{end_row - 1}").api.EntireRow.Delete()
                
                ws1_write_pointer = start_row + 1
                layer = 1
                if smt is not None and smb is not None:
                    start_cell = smt + 1
                    end_cell = smb - 1
                else:
                    start_cell = l01
                    end_cell = l_end
                # Update the cells using the mappings for the first worksheet
                for row_number, cell in enumerate(ws2.range(f'A{start_cell}:A{end_cell}'), start=start_cell):
                    # print(ws1_write_pointer, row_number, cell.value)
                    ws1.api.Rows(ws1_write_pointer).Insert()
                    if 'Base Material' in cell.value:
                        base_material_value = ws2.range(f"B{row_number}").value
                        if 'Base Material' in ws2.range(f'A{row_number -1}').value:
                            if 'X' in base_material_value:
                                [times, v] = base_material_value.split('X')
                                v = re.sub(r'(?<=\d)\D.*', '', v)
                                ws1.range(f"B{ws1_write_pointer - 2}").value += f"\n{times}x Prepreg FR4-{v}"
                            else:
                                base_material_value = base_material_value.replace("um", "").strip()
                                ws1.range(f"B{ws1_write_pointer - 2}").value += f"\nCore FR4 {base_material_value}"
                            ws1.api.Rows(ws1_write_pointer).Delete()
                            continue

                        text_1 = f"Basis Material"
                        text_2 = f"base material"
                        ws1.range(f"A{ws1_write_pointer}").value = text_1 + "\n" + text_2
                        ws1.range(f"A{ws1_write_pointer}").api.GetCharacters(1, len(text_1)).Font.Size = 10
                        ws1.range(f"A{ws1_write_pointer}").api.GetCharacters(len(text_1) + 1, len(text_2)).Font.Size = 8
                        ws1.range(f"A{ws1_write_pointer}").api.Font.Name = "Arial"

                        ws1.range(f"B{ws1_write_pointer}:D{ws1_write_pointer}").merge()
                        ws1.range(f"A{ws1_write_pointer}:E{ws1_write_pointer}").color = (255,255,255)

                        if 'X' in base_material_value:
                            [times, v] = base_material_value.split('X')
                            v = re.sub(r'(?<=\d)\D.*', '', v)
                            ws1.range(f"B{ws1_write_pointer}").value = f"{times}x Prepreg FR4-{v}"
                        else:
                            base_material_value = base_material_value.replace("um", "").strip()
                            ws1.range(f"B{ws1_write_pointer}").value = f"Core FR4 {base_material_value}"

                    
                    else:
                        #assuming its always layer here
                        is_outer_layer = True if row_number == start_cell or row_number == end_cell else False
                        layer_value = ws2.range(f"B{row_number}").value
                        layer_value = layer_value.split('um')[0]
                        
                        ws1.range(f"B{ws1_write_pointer}:D{ws1_write_pointer}").merge()
                        ws1.range(f"A{ws1_write_pointer}:E{ws1_write_pointer}").color = (255,204,153)
                        ws1.range(f"E{ws1_write_pointer}").number_format = 'General'
                        if is_outer_layer:
                            text_1 = f"Kupfer Aussenlage {layer}"
                            text_2 = f"copper outerlayer {layer}"
                            combined_text = text_1 + "\n" + text_2

                            cell = ws1.range(f"A{ws1_write_pointer}")
                            cell.value = combined_text

                            cell.api.GetCharacters(1, len(text_1) - len(str(layer))).Font.Size = 10
                            cell.api.GetCharacters(len(text_1) - len(str(layer)) + 1, len(text_1)).Font.Size = 10
                            cell.api.GetCharacters(len(text_1) + 1, len(text_1) + len(text_2) - len(str(layer))).Font.Size = 8
                            cell.api.GetCharacters(len(text_1) + len(text_2) - len(str(layer)) + 1, len(text_1) + len(text_2)).Font.Size = 8
                            cell.api.Font.Name = "Arial"
                            
                            ws1.range(f"B{ws1_write_pointer}:D{ws1_write_pointer}").value = f"{layer_value} µm Kupferfolie + galv Kupfer \n {layer_value} µm copper foil + plated copper"
                            ws1.range(f"E{ws1_write_pointer}").formula = urwete_mapping_formula_outer_layer
                            ws1.range(f"F{ws1_write_pointer}").formula = Lagenauf_formula_layer 
                        else:
                            text_1 = f"Kupfer Innenlage {layer}"
                            text_2 = f"copper innerlayer {layer}"
                            combined_text = text_1 + "\n" + text_2

                            cell = ws1.range(f"A{ws1_write_pointer}")
                            cell.value = combined_text

                            cell.api.GetCharacters(1, len(text_1) - len(str(layer))).Font.Size = 10
                            cell.api.GetCharacters(len(text_1) - len(str(layer)) + 1, len(text_1)).Font.Size = 10
                            cell.api.GetCharacters(len(text_1) + 1, len(text_1) + len(text_2) - len(str(layer))).Font.Size = 8
                            cell.api.GetCharacters(len(text_1) + len(text_2) - len(str(layer)) + 1, len(text_1) + len(text_2)).Font.Size = 8
                            cell.api.Font.Name = "Arial"

                            ws1.range(f"B{ws1_write_pointer}:D{ws1_write_pointer}").value = f"{layer_value} µm Kupferfolie \n {layer_value} µm copper foil"
                            ws1.range(f"E{ws1_write_pointer}").formula = urwete_mapping_formula_inner_layer
                            ws1.range(f"F{ws1_write_pointer}").formula = Lagenauf_formula_layer 

                        layer += 1

                    ws1.api.Rows(ws1_write_pointer+1).Insert()
                    ws1.range(f"A{ws1_write_pointer}:A{ws1_write_pointer+1}").merge()
                    ws1.range(f"B{ws1_write_pointer}:B{ws1_write_pointer+1}").merge()
                    ws1.range(f"E{ws1_write_pointer}:E{ws1_write_pointer+1}").merge()
                    ws1.range(f"F{ws1_write_pointer}:F{ws1_write_pointer+1}").merge()
                    # print(f"Updated row number in ws1: {ws1_write_pointer}")
                    ws1_write_pointer += 2
                # print(f"Last updated row number in ws1: {ws1_write_pointer - 1}")
                for row_number, cell in enumerate(ws1.range('A:A'), start=1):
                    if row_number > 100:
                        break

                    if cell.color == green_color:
                        if start_row is None:
                            start_row = row_number
                        else:
                            end_row = row_number

                search_string = "Ist:"

                for row_number, cell in enumerate(ws1.range('B:B'), start=1):
                    if row_number > 100:
                        break

                    if type(cell.value) == str and search_string in cell.value:
                        # print(f"Search string found in cell {cell.address}: {cell.value}")
                        next_cell = ws1.range(f"C{row_number}")
                        sum_formula = f"=SUM(F{start_row}:F{end_row})/1000"
                        next_cell.formula = sum_formula
                        # print(f"Search string found in cell {cell.address}: {cell.value}")

                # if sum_formula:
                #     print(f"SUM formula: {sum_formula}")     
                if smb is None and smt is None:
                    if end_row is not None:
                        ws1.range(f"A{end_row}").api.EntireRow.Delete()
                    if start_row is not None:
                        ws1.range(f"A{start_row}").api.EntireRow.Delete()
                    end_row = None
                    start_row = None
                elif smb is None:
                    if end_row is not None:
                        ws1.range(f"A{end_row}").api.EntireRow.Delete()
                    end_row = None
                elif smt is None:
                    if start_row is not None:
                        ws1.range(f"A{start_row}").api.EntireRow.Delete()
                    start_row = None
            except Exception as e:
                print(f"An error occurred while processing the 'Lagenaufbau' sheet: {e}")

        def cell_mapping_Schliffbilder(file1, file2):
            try:
                [ws1, ws2] = open_sheet_with_name('Schliffbilder', 'Schliffbilder')
                # print(cell_mapping_Schliffbilder)
                
                # Update the cells using the mappings
                for cell_ref, (source_cell, prefix) in cell_mapping_for_Schliffbilder.items():
                    source_value = ws2.range(source_cell).value
                    
                    if "Ø" in ws1.range(cell_ref).value:
                        index = ws1.range(cell_ref).value.find("Ø")
                        ws1.range(cell_ref).value = ws1.range(cell_ref).value[:index + 1] + " " +source_value #+ ws1.range(cell_ref).value[index + 1:]
                    
                    elif "Lage" in ws1.range(cell_ref).value:
                        index = ws1.range(cell_ref).value.find("Lage")
                        ws1.range(cell_ref).value = ws1.range(cell_ref).value[:index + 4] + str(int(source_value)) + ws1.range(cell_ref).value[index + 7:]

                    if "Hole diameter" in ws1.range(cell_ref).value and "layer" in ws1.range(cell_ref).value:
                        hole_index = ws1.range(cell_ref).value.find("Hole diameter")
                        layer_index = ws1.range(cell_ref).value.find("layer")

                        if layer_index > hole_index:
                            ws1.range(cell_ref).value = (
                            ws1.range(cell_ref).value[:hole_index ] +
                            ws1.range(cell_ref).value[hole_index :layer_index + 5] +
                            " " + str(int(source_value)) +
                            ws1.range(cell_ref).value[layer_index + 7:]
                            )
                        

                    elif "Hole diameter" in ws1.range(cell_ref).value:
                        index = ws1.range(cell_ref).value.find("Hole diameter")
                        ws1.range(cell_ref).value = ws1.range(cell_ref).value[:index + 12] + " " +source_value + ws1.range(cell_ref).value[index + 18:]

                    elif "layer" in ws1.range(cell_ref).value:
                        index = ws1.range(cell_ref).value.find("layer")
                        ws1.range(cell_ref).value = ws1.range(cell_ref).value[:index + 5] + str(int(source_value)) + ws1.range(cell_ref).value[index + 8:]

                    # else:
                    #     ws1.range(cell_ref).value = source_value
            except Exception as e:
                print(f"An error occurred while processing the 'Schliffbilder' sheet: {e}")            

        def cell_mapping_SchliffbilderViaHolefilling(file1, file2):
            try:
                [ws1, ws2] = open_sheet_with_name('Schliffbilder Via Hole filling', 'Schliffbilder Via Hole filling')

                for cell_ref, (source_cells, prefix) in cell_mapping_for_SchliffbilderViaHolefilling.items():
                        source_values = []
                        for source_cell in source_cells.split(', '):
                            source_values.append(ws2.range(source_cell).value)

                        target_cell_value = ws1.range(cell_ref).value
                        if "Hole diameter" in target_cell_value and "layer" in target_cell_value:
                            x, y = source_values
                            substituted_prefix = prefix.format(x=x, y=y)
                            ws1.range(cell_ref).value = substituted_prefix
                            # print(substituted_prefix)

                        elif "Hole diameter" in ws1.range(cell_ref).value:
                            index = ws1.range(cell_ref).value.find("Hole diameter")
                            ws1.range(cell_ref).value = ws1.range(cell_ref).value[:index + 12] + " " + source_values[0] + ws1.range(cell_ref).value[index + 18:]

                        elif "layer" in ws1.range(cell_ref).value:
                            index = ws1.range(cell_ref).value.find("layer")
                            ws1.range(cell_ref).value = ws1.range(cell_ref).value[:index + 5] + source_values[0] + ws1.range(cell_ref).value[index + 8:]

                        else:
                            ws1.range(cell_ref).value = source_values
            except Exception as e:
                print(f"An error occurred while processing the 'Schliffbilder via hole filling' sheet: {e}")       
                

        def cell_mapping_MasterData(file1, file2): 
            try:
                [ws1, ws2] = open_sheet_with_name(file1, 'MasterData')
                for ws1 in wb1.sheets:
                    for cell_ref, (source_cell, prefix, sheet_name) in cell_mapping_for_MasterData.items():
                        value = ws2.range(source_cell).value
                        if ws1.name == sheet_name:
                            if isinstance (value, str) and (value.lower() == 'white' or value.lower() == 'green' or value.lower() == 'yellow'):
                                ws1_range = ws1.range(cell_ref)
                                data_validation = ws1_range.api.Validation
                                if data_validation:
                                    # print(f"Data validation is present in cell: {cell_ref}")
                                    # print(f"data_validation Type: {data_validation.Type}")
                                    try:
                                        if data_validation.Type == 3:  
                                            formula_text = data_validation.Formula1
                                            formula_range = ws1.range(formula_text)  
                                            if value.lower() == 'white':
                                                cell_value = formula_range[0].value
                                            elif value.lower() == 'yellow':
                                                cell_value = formula_range[1].value
                                            elif value.lower() == 'green':
                                                cell_value = formula_range[2].value
                                            ws1.range(cell_ref).value = cell_value 
                                    except Exception as e:
                                        pass
                            # if value is not None:
                            elif value or value is None:
                                ws1_range = ws1.range(cell_ref)
                                data_validation = ws1_range.api.Validation
                                if data_validation:
                                    try:
                                        if data_validation.Type == 3:  
                                            formula_text = data_validation.Formula1
                                            formula_range = ws1.range(formula_text)  
                                            if value is not None:
                                                cell_value = formula_range[0].value
                                            else:
                                                cell_value = formula_range[1].value 
                                            ws1.range(cell_ref).value = cell_value
                                    except:
                                        ws1.range(cell_ref).value = value
                            else:
                                ws1.range(cell_ref).value = value
            except Exception as e:
                print(f"An error occured while processing the 'master data' sheet :{e}")

        def cell_mapping_ManualData(file1, file2):
            try:
                [ws1, ws2] = open_sheet_with_name(file1, 'ManualData')
                for ws1 in wb1.sheets:
                    # print(f"Processing sheet: {ws1.name}")
                    # Update the cells using the mappings
                    for cell_ref, (source_cell, prefix, sheet_name) in cell_mapping_for_ManualData.items():
                        value = ws2.range(source_cell).value
                        if ws1.name == sheet_name:
                            # print(f"Updating cell {cell_ref} with value {value}")
                            ws1.range(cell_ref).value = value
            except Exception as e:
                print(f"An error occured while processing the 'manual data' sheet :{e}")

        # # Call the functions with the file paths 
        cell_mapping_Laufzettel(sys.argv[1],sys.argv[2])                    
        cell_mapping_Urwerte(sys.argv[1],sys.argv[2])                    
        cell_mapping_Prufplan(sys.argv[1],sys.argv[2])                    
        cell_mapping_Lagenaufbau(sys.argv[1],sys.argv[2]) 
        cell_mapping_Schliffbilder(sys.argv[1],sys.argv[2])  
        cell_mapping_SchliffbilderViaHolefilling(sys.argv[1],sys.argv[2])                    
        cell_mapping_MasterData(sys.argv[1],sys.argv[2]) 
        cell_mapping_ManualData(sys.argv[1],sys.argv[2])                  


        # save after updation 
        save_and_close_excel(app, wb1, wb2, new_filename)
        print("DONE", flush=True)
        sys.stdout.flush()
        
    # except Exception as e:
    #     print(f"An error occurred: {str(e)}")
    #     sys.stdout.flush()
    # finally:
    #     app.quit()    

if __name__ == "__main__":
    folder_path = sys.argv[4]
    new_filename = sys.argv[5]
    main()

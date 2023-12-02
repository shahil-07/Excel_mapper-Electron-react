import subprocess

template = 'Siemens_Amberg_Rev-16_F01.xlsm'
sd = 'empb_data.xlsx '
cm = 'cell_mapping_siemens.xlsx'

subprocess.run(["./public/dist/Excel_mapper.exe", template, sd, cm])
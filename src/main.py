import pandas as pd
import os
import glob

folder_path = 'src\\data\\raw'
excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

if not excel_files:
    print("Nenhuma planilha encontrada")
else:
    dfs = []
    for excel_file in excel_files:
        try:
            df_temp = pd.read_excel(excel_file)
            file_name = os.path.basename(excel_file)

            df_temp['filename'] = excel_file
            
            # LOCATION
            if 'brasil' in file_name.lower():
                df_temp['location'] = 'br'
            elif 'france' in file_name.lower():
                df_temp['location'] = 'fr'
            elif 'italian' in file_name.lower():
                df_temp['location'] = 'it'

            #CAMPAIGN
            df_temp['campaign'] = df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')

            dfs.append(df_temp)

        except Exception as e:
            print(f"Erro ao ler arquivo {excel_file} : {e}")

    if dfs:
        result = pd.concat(dfs, ignore_index=True)
        output_file = os.path.join('src', 'data', 'ready', 'clean.xlsx')
        writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
        result.to_excel(writer, index=False) #sheet_name='name'
        writer.close()
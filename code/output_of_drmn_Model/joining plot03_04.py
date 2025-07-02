import os
import pandas as pd


base_path = r"D:\project\code new\Problem"


folders = [os.path.join(base_path, name) for name in os.listdir(base_path)
           if os.path.isdir(os.path.join(base_path, name))]

def reverse_and_merge_excels(folder_path):
    file1 = os.path.join(folder_path, "output_plot03.xlsx")
    file2 = os.path.join(folder_path, "output_plot04.xlsx")
    output_file = os.path.join(folder_path, "output.xlsx")

    if not os.path.exists(file1) or not os.path.exists(file2):
        print(f"Missing files in {folder_path}")
        return

    
    excel1 = pd.read_excel(file1, sheet_name=None)
    excel2 = pd.read_excel(file2, sheet_name=None)

    writer = pd.ExcelWriter(output_file, engine='openpyxl')

    for sheet_name in excel1.keys():
        if sheet_name not in excel2:
            continue

        df1 = excel1[sheet_name]
        df2 = excel2[sheet_name]

        if 'Observed' not in df1.columns or 'Observed' not in df2.columns:
            print(f"Missing 'Observed' column in {sheet_name} of {folder_path}")
            continue

        # Observed
        observed = df1['Observed'].reset_index(drop=True)

        # reversing
        df1_other = df1.drop(columns=['Observed']).iloc[::-1].reset_index(drop=True)
        df2_other = df2.drop(columns=['Observed']).iloc[::-1].reset_index(drop=True)

        
        final_data = pd.concat([observed, df1_other, df2_other], axis=1)

      
        top_row = [''] + ['New'] * df1_other.shape[1] + ['PSO'] * df2_other.shape[1]
        col_names = ['Observed'] + list(df1_other.columns) + list(df2_other.columns)

       
        final_data.columns = col_names

        
        top_df = pd.DataFrame([top_row], columns=col_names)
        final_data = pd.concat([top_df, final_data], ignore_index=True)

        
        final_data.to_excel(writer, sheet_name=sheet_name, index=False)

    writer.close()
    print(f"Processed: {folder_path}")


for folder in folders:
    reverse_and_merge_excels(folder)

print("Done")

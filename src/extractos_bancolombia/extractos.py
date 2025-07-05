import pandas as pd
import glob
import os


# solo NO haz skiprows de 2 en el primer archivo
excel_files = glob.glob(os.path.join("docs", "*.xlsx"))
df_list = [pd.read_excel(excel_files[0])] + [pd.read_excel(f, skiprows=2) for f in excel_files[1:]]
df = pd.concat(df_list, ignore_index=True)

print("DataFrame concatenado:")
print(df)

#guardar el dataframe concatenado en un nuevo archivo Excel
output_file = "docs/concatenated_output.xlsx"
df.to_excel(output_file, index=False)
print(f"DataFrame guardado en {output_file}")

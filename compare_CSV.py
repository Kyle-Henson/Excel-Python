import pandas as pd

# Completely drop all duplicate rows in the new csv from the master csv
old_file = 'G:\\Master.csv'
new_file = 'G:\\New.csv'
out_file = 'G:\\Results.csv'

df1=pd.read_csv(old_file)
df2=pd.read_csv(new_file)

# https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.drop_duplicates.html
df_final=pd.concat([df1,df1,df2]).drop_duplicates(
    keep=False, subset=['Column1', 'Column2', 'Column3'])
print(df_final.shape)

df_final.to_csv(out_file, index=False)
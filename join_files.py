import pandas as pd

# Загрузите три файла
file1 = 'tables0.xlsx'
file2 = 'tables1.xlsx'
file3 = 'tables2.xlsx'

# Прочитайте каждый файл в DataFrame
df1 = pd.read_excel(file1)
df2 = pd.read_excel(file2)
df3 = pd.read_excel(file3)

df1 = df1.loc[:, ~df1.columns.str.contains('^Unnamed')]
df2 = df2.loc[:, ~df2.columns.str.contains('^Unnamed')]
df3 = df3.loc[:, ~df3.columns.str.contains('^Unnamed')]

# Объедините DataFrame по строкам (axis=0)
merged_df = pd.concat([df1, df2, df3], axis=0)

# Сброс индекса после объединения (если нужно)
merged_df.reset_index(drop=True, inplace=True)

# Сохраните объединенный DataFrame в новый Excel-файл
merged_df.to_excel('merged_file.xlsx', index=False)

print("Файлы успешно объединены и сохранены в 'merged_file.xlsx'")
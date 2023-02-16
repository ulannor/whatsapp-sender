


# data = pd.read_excel(filepath, index_col=0, dtype='object')
df = pd.read_excel(filepath, dtype='object')

if 'wa_status' not in df.columns:
    df.insert(3, "wa_count", 0)
    df.insert(4, "wa_status", '')
    df.insert(5, "wa_message", '')

n = 50

wa_list = [col for col in df if col.startswith('whatsapp')]
non_wa_list = [col for col in df if col.startswith('phone')]

print(wa_list)

print(non_wa_list)

for i in range(n):
    i_list = []
    if pd.isna(df.iloc[i]['whatsapp 1']):
        for el in non_wa_list:
            res = df.iloc[i][el]
            if pd.isna(res):
                continue
            i_list.append(res)
    else:
        for el in wa_list:
            res = df.iloc[i][el]
            if pd.isna(res):
                continue
            i_list.append(res)

    # print(i_list)





# data.to_excel(f'.\\sourcedata\\test.xlsx')



wb_obj = op.load_workbook(filepath, read_only=True)

sheet_obj = wb_obj.active


for row in range(1, 5):
    print(sheet_obj[row])
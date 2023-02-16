import pandas as pd

leadsfile = 'test'
messagetxt = 'text'
n = 20

filepath = f'.\\edited\\{leadsfile}.xlsx'
txtpath = f'.\\messages\\{messagetxt}.txt'

df = pd.read_excel(filepath, dtype='object')

print(df.iloc[0])

waNumberColList = [col for col in df if col.startswith('whatsapp')]
nonwaNumberColList = [col for col in df if col.startswith('phone')]

try:
    for _ in range(0, n):
        walist = []

        for col in waNumberColList:
            phone = df.iloc[_][col]
            if pd.isna(phone):
                break
            print(phone)
            print('waphone')
            print(type(phone))
            walist.append(phone)
            print(df.iloc[_]['address name'])

        for col in nonwaNumberColList:
            phone = df.iloc[_][col]
            if pd.isna(phone):
                break
            elif phone.startswith('+996312'):
                break
            elif phone in walist:
                break
            print(phone)
            print('simphone')
            print(type(phone))
            print(df.iloc[_]['address name'])

except IndexError as e:
    print(e)

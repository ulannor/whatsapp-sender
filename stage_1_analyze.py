import pandas as pd
from collections import Counter
import openpyxl as op
from openpyxl.styles import PatternFill, Font
from openpyxl.formatting import Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter
import sender_functions as sf

filename = 'нотариальные услуги'

filepath = f'.\\sourcedata\\{filename}.xlsx'

df = pd.read_excel(filepath, dtype='object')

if 'wanumbers_available' not in df.columns:
    df.insert(10, "wanumbers_available", 0)
    df.insert(11, "message_count", 0)

originalColList = list(df.columns)
print(originalColList)
waNumberColList = [col for col in df if col.startswith('whatsapp')]
nonwaNumberColList = [col for col in df if col.startswith('phone')]

for col in waNumberColList:
    df.loc[df[col].notnull(), 'wanumbers_available'] += 1

logFilepath = f'.\\wa_mailout_log\\wa_mailout_log.xlsx'

dfLog = pd.read_excel(logFilepath, dtype='object')

urlList = dfLog['2GIS URL'].tolist()

urlAdjustedList = [urlList[i] for i in range(len(urlList)) if (i == 0) or urlList[i] != urlList[i - 1]]
urlCounterDict = Counter(urlAdjustedList)

df["message_count"] = df["2GIS URL"].map(urlCounterDict)

dfLog['temp_index'] = dfLog['2GIS URL']

dfLog.set_index('temp_index', inplace=True)

df['temp_index'] = df['2GIS URL']
df.set_index('temp_index', inplace=True)

dfLog['last_message_text'] = dfLog.dropna(subset=['message_text']).groupby('2GIS URL', group_keys=False)[
    'message_text'].last()
dfLog['last_message_date'] = dfLog.dropna(subset=['date_and_time']).groupby('2GIS URL', group_keys=False)[
    'date_and_time'].last()
dfLog['message_wanumbers'] = dfLog.dropna(subset=['message_wanumber']).groupby('2GIS URL', group_keys=False)[
    'message_wanumber'].apply(lambda x: list(set(x)))
dfLog['message_nonwanumbers'] = dfLog.dropna(subset=['message_nonwanumber']).groupby('2GIS URL', group_keys=False)[
    'message_nonwanumber'].apply(lambda x: list(set(x)))
dfLog['2GISid_processed'] = dfLog.dropna(subset=['organization id']).groupby('2GIS URL', group_keys=False)[
    'organization id'].apply(lambda x: len(list(set(x))))
dfLog['2GISurl_processed'] = dfLog.dropna(subset=['2GIS URL']).groupby('2GIS URL', group_keys=False)['2GIS URL'].apply(
    lambda x: len(list(set(x))))
dfLog['wanumbers_processed'] = dfLog.dropna(subset=['message_wanumber']).groupby('2GIS URL', group_keys=False)[
    'message_wanumber'].apply(lambda x: len(list(set(x))))
dfLog['nonwanumbers_processed'] = dfLog.dropna(subset=['message_nonwanumber']).groupby('2GIS URL', group_keys=False)[
    'message_nonwanumber'].apply(lambda x: len(list(set(x))))

dfLog.to_excel(f'.\\test\\{filename}_edited.xlsx', index=False)

newColList = list(dfLog.columns)[15:23]

dfLogNoDups = dfLog.drop_duplicates(subset='2GIS URL', keep='last')

for name in newColList:
    if name == '2GISid_processed':
        df = df.merge(dfLogNoDups[[name, 'organization id']], on='organization id', how='left')
    else:
        df = df.merge(dfLogNoDups[[name, '2GIS URL']], on='2GIS URL', how='left')

for name in newColList:
    col = df.pop(name)
    df.insert(12 + newColList.index(name), col.name, col)

df.to_excel(f".\\test\\{filename}_edited2.xlsx", index=False)

print(df.columns.tolist())
print(len(df.columns.tolist()))

tempLst = []
for _ in waNumberColList + nonwaNumberColList:
    tempLst.append(df.columns.tolist().index(_) + 1)

print(tempLst)

wb = op.load_workbook(f".\\test\\{filename}_edited2.xlsx")
ws = wb.active

red_text = Font(color="9C0006")
red_fill = PatternFill(bgColor="FFC7CE")
dxf = DifferentialStyle(font=red_text, fill=red_fill)
rule = Rule(type="duplicateValues", text="highlight", dxf=dxf)
bgn = 1
end = 100000
for _ in tempLst:
    column_letter = get_column_letter(_)
    ws.conditional_formatting.add(f'{column_letter}{bgn}:{column_letter}{end}', rule)
ws.conditional_formatting.add(f'A{bgn}:A{end}', rule)
ws.auto_filter.ref = ws.dimensions

wb.save(f".\\test\\{filename}_edited3.xlsx")

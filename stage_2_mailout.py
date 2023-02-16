import pandas as pd

leadsfile = 'нотариальные услуги'
messagetxt = 'text'
n = 20

filepath = f'.\\edited\\{leadsfile}_edited.xlsx'
txtpath =  f'.\\messages\\{messagetxt}.txt'

df = pd.read_excel(filepath, dtype='object')


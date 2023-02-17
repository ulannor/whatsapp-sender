import pandas as pd
import sender_functions as sf


leadsfile = 'test'
messagetxt = 'text'
logfile = 'testlog'

n = 20

filepath = f'.\\edited\\{leadsfile}.xlsx'
txtpath = f'.\\messages\\{messagetxt}.txt'
logpath = f'.\\wa_mailout_log\\{logfile}.xlsx'

df = pd.read_excel(filepath, dtype='object')


waNumberColList = [col for col in df if col.startswith('whatsapp')]
nonwaNumberColList = [col for col in df if col.startswith('phone')]

msg = sf.read_txt_file(txtpath)


try:
    for _ in range(0, n):
        walist = []

        for col in waNumberColList:
            phone = df.iloc[_][col]
            if pd.isna(phone):
                break
            walist.append(phone)
            dftemp = df.iloc[_, 0:11]
            dftemp['message_wanumber'] = phone
            dftemp['message_nonwanumber'] = ''
            dftemp['date_and_time'] = pd.Timestamp.now().strftime("%d/%m/%Y, %H:%M:%S")
            dftemp['message_text'] = msg
            dftemp = dftemp.to_frame()
            phone = sf.format_phone(phone)
            print(phone)
            sf.send_msg(phone, msg)
            sf.logwriter(dftemp, logpath)

        for col in nonwaNumberColList:
            phone = df.iloc[_][col]
            if pd.isna(phone):
                break
            elif phone.startswith('+996312'):
                break
            elif phone in walist:
                break
            dftemp = df.iloc[_, 0:11]
            dftemp['message_wanumber'] = ''
            dftemp['message_nonwanumber'] = phone
            dftemp['date_and_time'] = pd.Timestamp.now().strftime("%d/%m/%Y, %H:%M:%S")
            dftemp['message_text'] = msg
            dftemp = dftemp.to_frame()
            phone = sf.format_phone(phone)
            print(phone)
            sf.send_msg(phone, msg)
            sf.logwriter(dftemp, logpath)

except IndexError as e:
    print(e)

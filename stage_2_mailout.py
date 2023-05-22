import pandas as pd
import sender_functions as sf

# file names and parameters
leadsfile = 'нотариальные услуги_edited'   # name of the leads file
messagetxt = 'text'                        # name of the message text file
n = 120                                    # number of rows to process in the leads file

# file paths
LEADSFILEPATH = f'.\\edited\\{leadsfile}.xlsx'  # path to the leads file
MSGPATH = f'.\\messages\\{messagetxt}.txt'  # path to the message text file
LOGPATH = f'.\\wa_mailout_log\\wa_mailout_log.xlsx'  # path to the log file

# read the leads file into a pandas dataframe
df = pd.read_excel(LEADSFILEPATH, dtype='object')

# get the list of columns containing WhatsApp numbers and non-WhatsApp numbers
waNumberColList = [col for col in df if col.startswith('whatsapp')]
nonwaNumberColList = [col for col in df if col.startswith('phone')]

# read the message text from the text file
msg = sf.read_txt_file(MSGPATH)

# iterate over each row in the leads file, sending messages to WhatsApp and non-WhatsApp numbers
for _ in range(0, n):
    walist = []

    # iterate over columns containing WhatsApp numbers
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
        sf.logwriter(dftemp, LOGPATH)

    # iterate over columns containing non-WhatsApp numbers
    for col in nonwaNumberColList:
        phone = df.iloc[_][col]
        if pd.isna(phone):
            break
        elif phone.startswith('+996312'):  # ignore phone numbers starting with +996312
            break
        elif phone in walist:  # ignore phone numbers that have already been sent a WhatsApp message
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
        sf.logwriter(dftemp, LOGPATH)

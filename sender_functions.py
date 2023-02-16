
def format_phone(phone):
    if len(phone) == 10 and phone[0] == '0':
        phone = phone.replace('0', '996', 1)
        return phone
    elif len(phone) == 9:
        phone = phone + '996'
        return phone
    for i in phone:
        if i in '-() +;':
            phone = phone.replace(i, '')
            return phone
        else:
            return phone
    phone = '+' + phone.replace('.0', '')
    return phone

def read_txt_file(txtpath):
    with open(txtpath, 'r', encoding='utf-8') as f:
        text_msg = ''
        for row in f:
            text_msg += row
    return text_msg


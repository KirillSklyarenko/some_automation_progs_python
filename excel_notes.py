#! python3

import openpyxl
import smtplib
import datetime

file = openpyxl.load_workbook(r'C:\____\file.xlsx', data_only=True)
sheet = file['notes_to_do']


def send_mail(texts):
    for i in texts:
        subject = 'My Note Alert'
        message = f'Subject: {subject}\n\n{i}'
        smtpobj = smtplib.SMTP('smtp.gmail.com', 587)
        smtpobj.ehlo()
        smtpobj.starttls()
        smtpobj.login('___@gmail.com', '___')
        smtpobj.sendmail('___@gmail.com', '___@gmail.com', message)
        smtpobj.quit()


header = [cell.value for cell in sheet[1]]
col_b = [i for i in sheet['B'] if type(i.value) is datetime.datetime]
col_c = [i for i in sheet['C'] if type(i.value) is datetime.datetime]

# compare current date, if it falls between start date of period in column B and end date of period in column C
# if it does, make a dict of row for the corresponding message and append to the list
mess_list = []
for x, y in zip(col_b, col_c):
    if datetime.datetime.strftime(x.value, "%Y, %m, %d") <= datetime.datetime.strftime(datetime.datetime.today(), "%Y, %m, %d") <= datetime.datetime.strftime(y.value, "%Y, %m, %d"):
        x_row_values = [i.value for i in sheet[x.row]]
        mess_list.append(dict(zip(header, x_row_values)))

# just for practice. Extract the value of "text" key in dict and append to list; the list of texts is passed to function
send_mail([i['text'] for i in mess_list])

# make a list of messages for events. Messages include notes for a week before, a day before and today
sheet_b = file['dates']
mess_list1 = []
for i in sheet_b['B']:
    if type(i.value) is datetime.datetime:
        if datetime.datetime.strftime(i.value, "%d, %m") == datetime.datetime.strftime(datetime.datetime.today(), "%d, %m"):
            a, b, c, d = sheet_b[i.row]
            message_today = f'{a.value} has {c.value} today!, anniversary: {d.value}'
            mess_list1.append(message_today)
        x = i.value
        today = datetime.datetime.today()
        if x.month == today.month and x.day - today.day == 1:
            a, b, c, d = sheet_b[i.row]
            message_tomorrow = f'{a.value} has {c.value} tomorrow, on {datetime.datetime.strftime(b.value, "%d.%m")}, anniversary: {d.value}'
            mess_list1.append(message_tomorrow)
        if x.month == today.month and x.day - today.day == 7:
            a, b, c, d = sheet_b[i.row]
            message_nextweek = f'{a.value} has {c.value} in a week, on {datetime.datetime.strftime(b.value, "%d.%m")}, anniversary: {d.value}'
            mess_list1.append(message_nextweek)

send_mail(mess_list1)

__author__ = 'hanwang'

#### HOW TO USE ####
# Put this python file in the same folder with roster.xlsx
####################

import xlrd
import smtplib

count = 0

# Reading email list
# No need to change this part while using this script
# mail_list=[[UIN,Last name, First name, Email]...]

mail_list = []

hw1 = xlrd.open_workbook('roster.xlsx').sheet_by_name('HW1')
for i in xrange(1,hw1.nrows):
    data = hw1.row(i)
    mail_list.append([str(int(data[0].value)), data[1].value, data[2].value, data[4].value])
    count += 1

# Select a sheet that has grades using sheet_by_name()

data = xlrd.open_workbook('roster.xlsx').sheet_by_name('HW1')
head = data.row(0)

d = 0
# Make sure the last column of the sheet is "Total"
if head[len(head)-1].value=='Total':
    d = len(head)-1
else:
    for i in xrange(len(head)):
        if head[i].value=='Total':
            d = i
            break

grades = data.col_values(d)[1:count+1]

# The theme of the email
title = 'HW1'

for i in xrange(len(grades)):
    uin = mail_list[i][0]
    last_name = mail_list[i][1]
    first_name = mail_list[i][2]
    mail_addr = mail_list[i][3]
    score = str(grades[i])
    fromaddr = '606softwareengineering@gmail.com'
    toaddrs  = '606softwareengineering@gmail.com'
    msg = "\r\n".join([
      "From: 606softwareengineering@gmail.com",
      "To: "+mail_addr,
      "Subject: CSCE606 "+title+" Grades",
      "",
      "Hi,"+first_name+",\n Your "+title+" gets "+score+" out of 100.0. If you have any questions, please reply to TAs.\n\nCSCE606 Software Engineering"
      ])


    # Credentials (if needed)
    username = 'username'
    password = 'password'

    # The actual mail send
    server = smtplib.SMTP('smtp.gmail.com:587')
    server.starttls()
    server.login(username,password)
    server.sendmail(fromaddr, toaddrs, msg)

    server.quit()

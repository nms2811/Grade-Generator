#!/home/mnoh1/software/bin/python3

# STUDENT OATH:
# -------------
#
# "I declare that the attached project is wholly my own work in accordance
# with Seneca Academic Policy. No part of this project has been copied
# manually or electronically from any other source (including web sites) or
# distributed to other students."
#
# Minseop Noh   ________________________   mnoh1  _________________


# import modules for CGI handling
import cgi,cgitb
import smtplib
import re
import os
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import copy
import xlsxwriter     # (to use a set of functions to write to and generate
              # excel spreadsheets in *.xlsx format)
from validate_email import validate_email


cgitb.enable( )  # enabled for CGI script troubleshooting
         # script langauge/runtime errors are displayed and sent back to
         # the browser

# create instance of FieldStorage to process CGI form values
form = cgi.FieldStorage( )

course   = form.getvalue('course')
quiz = form.getvalue('quiz')
quizVal = form.getvalue('value1')
lab = form.getvalue('lab')
labVal = form.getvalue('value2')
assignment = form.getvalue('assignment')
assignmentVal = form.getvalue('value3')
test = form.getvalue('test')
testVal = form.getvalue('value4')
exam = form.getvalue('exam')
examVal = form.getvalue('value5')
color = form.getvalue('bgcolor')
email = form.getvalue('email')
currentDate = datetime.datetime.now().strftime("%Y.%m.%d")
excelFileName = course.lower() + '.' + currentDate + '.xlsx'
def validation() :
    regEx = re.compile("[A-Z]{3}[0-9]{3}")
    count = 0
    total_mark = 0
    if labVal == "0.5" :
        total_mark += (int(quiz) * int(quizVal)) + (int(lab) * float(labVal)) + (int(assignment) * int(assignmentVal))
    else:
        total_mark += (int(quiz) * int(quizVal)) + (int(lab) * int(labVal)) + (int(assignment) * int(assignmentVal))
    #end if
    total_mark += (int(test) * int(testVal)) + (int(exam) * int(examVal))
    if validate_email(email) :
        count += 1
    else :
        print("email is not validated\n")
    #end if
    if total_mark == 100 :
        count += 1
    else :
        print("total mark is not 100%\n")
    #end if 
    if regEx.match(course) :
        count += 1
    else :
        print("Format of course name is wrong\n")
    #end if
    if count == 3:
        return True
    else :
        return False
    #end if
#end def

def excelPart(personalInfo) :
    count = 4
    workbook = xlsxwriter.Workbook(excelFileName)
    worksheet = workbook.add_worksheet( )
    cell_format = workbook.add_format()
    cell_format.set_bg_color(color)
    worksheet.set_column('A:A', 10)
    worksheet.set_column('B:B', 25)
    worksheet.set_column('C:C', 15)
    worksheet.set_column('D:D', 10)
    worksheet.write(0, 0, "Student #", cell_format)
    worksheet.write(0, 1, "Last", cell_format)
    worksheet.write(0, 2, "First", cell_format)
    worksheet.write(0, 3, "Login", cell_format)
    if int(quiz) != 0 :
        for i in range(int(quiz)) :
            worksheet.write(0,4 + i, "Q" + str(i + 1), cell_format)
            count += 1
        #end for
    #end if
    if int(lab) != 0 :
        for i in range(int(lab)) :
            worksheet.write(0, count, "L" + str(i + 1), cell_format)
            count += 1
        #end for
    #end if
    if int(assignment) != 0 :
        for i in range(int(assignment)) :
            worksheet.write(0, count, "A" + str(i + 1), cell_format)
            count += 1
        #end for
    #end if
    if int(test) != 0 :
        for i in range(int(test)) :
            worksheet.write(0, count, "T" + str(i + 1), cell_format)
            count += 1
        #end for
    #end if
    if int(exam) != 0 :
            worksheet.write(0, count, "EX", cell_format) 
            count += 1
        #end for
    #end if
    worksheet.write(0, count, "Score", cell_format)
    count += 1
    worksheet.write(0, count, "Grade", cell_format)
    #end for

    for i in range(len(personalInfo)) :
        string = ''
        c = 0
        check = 0
        checkString = ""
        worksheet.write(i + 1, 0,  personalInfo[i][2])
        worksheet.write(i + 1, 1,  personalInfo[i][0])
        worksheet.write(i + 1, 2,  personalInfo[i][1])
        worksheet.write(i + 1, 3,  personalInfo[i][3])
        for j in range (0, 5) :
            if check == 0:
                if j == 0 :
                    c += int(quiz)
                    checkString = "quiz"
                    if check == 0 and c != check :
                        string += '=SUM(SUM(' + chr(69) + str(i + 2) + ':' + chr(68 + int(c) ) + str(i + 2) + ')*' + quizVal
                        check += int(quiz)
                    #end if
                elif j == 1:
                    c += int(lab)
                    checkString = "lab"
                    if check == 0 and c != check :
                        string += '=SUM(SUM(' + chr(69) + str(i + 2) + ':' + chr(68 + int(c) ) + str(i + 2) + ')*' + labVal
                        check += int(lab)
                    #end if
                elif j == 2:
                    c += int(assignment)
                    checkString = "assignment"
                    if check == 0 and c != check :
                        string += '=SUM(SUM(' + chr(69) + str(i + 2) + ':' + chr(68 + int(c)) + str(i + 2) + ')*' + assignmentVal
                        check += int(assignment)
                    #end if
                elif j == 3:
                    c += int(test)
                    checkString = "test"
                    if check == 0 and c != check :
                        string += '=SUM(SUM(' + chr(69) + str(i + 2) + ':' + chr(68 + int(c) ) + str(i + 2) + ')*' + testVal
                        check += int(test)
                    #end if
                else:
                    c += int(exam)
                    checkString = "exam"
                    if check == 0 and c != check :
                        string += '=SUM(SUM(' + chr(69) + str(i + 2) + ')*' + examVal
                        check += int(exam)
                    #end if
                #end if/else
            else :
                if j == 0 :
                    c += int(quiz)
                    if check != c and checkString != "quiz":
                        string += ',SUM(' + chr(69 + check) + str(i + 2) + ':' + chr(68 + int(c)) + str(i + 2) + ') *' + quizVal
                        check += int(quiz)
                    #end if
                elif j == 1:
                    c += int(lab)
                    if check != c and checkString != "lab":
                        string += ', SUM(' + chr(69 + check) + str(i + 2) + ':' + chr(68 + int(c)) + str(i + 2) + ')*' + labVal
                        check += int(lab)
                    #end if
                elif j == 2:
                    c += int(assignment)
                    if check != c and checkString != "assignment":
                        string += ', SUM(' + chr(69 + check) + str(i + 2) + ':' + chr(68 + int(c)) + str(i + 2) + ')*' + assignmentVal
                        check += int(assignment)
                    #end if
                elif j == 3:
                    c += int(test)
                    if check != c and checkString != "test":
                        string += ', SUM(' + chr(69 + check) + str(i + 2) + ':' + chr(68 + int(c)) + str(i + 2) + ')*' + testVal
                        check += int(test)
                    #end if
                else:
                    c += int(exam)
                    if check != c and checkString != "exam":
                        string += ', ' + chr(69 + check) + str(i + 2) + '*' + examVal 
                        check += int(exam)
                    #end if
                #end if/else
            #end if/else
        #end for
        string += ')/100'
        worksheet.write_formula(i + 1, count - 1, string)
        worksheet.write_formula(i + 1, count, '=IF('+ chr(64 + count)+ str(i + 2) +'<55,"F",IF('+ chr(64 + count )+ str(i + 2) +'<60,"D",IF('+ chr(64 + count)+ str(i + 2) +'<70,"C",IF(' + chr(64 + count )+ str(i + 2) +'<80, "B",IF(' + chr(64 + count)+ str(i + 2) + '<95,"A","A+"))))) ')
    #end for
    workbook.close( )
#end def

print("Content-type: text/html\n\n")

print("<html>\n")
print("<head>\n")
print("<title>Assignment 2</title>\n")
print("</head>\n")
print("<body>\n")
if (validation()) :
    personalInfo = []
    fh = open("data.dat", 'w')
    fh.write("Smith;Bill;012345678;bsmith\n")
    fh.write("Swift;Tom;111222333;tswift\n")
    fh.write("Meadows;Audrey;001002003;ameadows\n")
    fh.write("Bright;Sally;166831798;sbright")
    fh.write("Taylor;John;333444999;jtaylor")
    fh.close()
    fh1 = open("data.dat", 'r')
    data = fh1.read( )
    newlist = data.split("\n")
    for i in range(len(newlist)):
        temp = newlist[i].split(";")
        personalInfo.append(temp)
    #end for
    excelPart(personalInfo)
    print("Successfully worked!!!!!!")
    os.chmod(excelFileName, 0o777)       #modify
    fromaddr = "mnoh1@myseneca.ca"      # modify
    toaddr = email      # modify

    msg = MIMEMultipart( )

    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = "prg469_191aUserID Excel Generator"     # modify

    body = "Excel file is attached"      # modify

    msg.attach(MIMEText(body, 'plain'))

    filename = excelFileName        # modify
    attachment = open(filename, "rb")  # modify

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

    msg.attach(part)

    server = smtplib.SMTP('outlook.office365.com', 587)
    server.starttls( )
    server.login(fromaddr, ",Als891124")  # modify
    text = msg.as_string( )
    server.sendmail(fromaddr, toaddr, text)
    server.quit( )
else :
    print("<br>rewrite the form\n")
#end if/else
print("</body>\n")
print("</html>\n")

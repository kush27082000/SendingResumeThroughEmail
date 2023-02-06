#!/usr/bin/env python
import smtplib
import os
from email.mime.text import MIMEText
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from xlutils.copy import copy
from xlwt import Workbook, easyxf
from xlrd import open_workbook, cellname
from time import gmtime, strftime
import datetime
import argparse
from email import Encoders

username  = None
password  = None
real_name = None
debug_mode = False
parser = argparse.ArgumentParser(description='Send Email Applications')
parser.add_argument('-g', '--gen', help='generate template excels', action='store_true')
parser.add_argument('-c', '--commit', help='commit mode, otherwise will not send the email', action='store_true')
parser.add_argument('-t', '--test', help='test mode, will send email to self instead', action='store_true')

args = parser.parse_args()

def gen_temp(gen_type):
    # type = 1, generate app_info excel
    # type = 2, generate log_info excel
    # type = 3, clean all so generate both
    def gen_app_temp():
        app_book = Workbook(encoding='utf-8')
        sheet1 = app_book.add_sheet('Application Info')
        sheet1.write(0, 0, 'Company Name')
        sheet1.write(0, 1, 'Job Title')
        sheet1.write(0, 2, 'Contact Name')
        sheet1.write(0, 3, 'Contact Address')
        sheet1.write(0, 4, 'Recipient Email')
        sheet1.write(0, 5, 'Transcript')
        sheet1.write(0, 6, 'GRE')
        sheet1.write(0, 7, 'Template No')
        app_book.save('./application_info.xls')

    def gen_log_temp():
        log_book = Workbook(encoding='utf-8')
        sheet2 = log_book.add_sheet('Application Info')
        sheet2.write(0, 0, 'Time')
        sheet2.write(0, 1, 'Company Name')
        sheet2.write(0, 2, 'Job Title')
        sheet2.write(0, 3, 'Contact Name')
        sheet2.write(0, 4, 'Contact Address')
        sheet2.write(0, 5, 'Recipient Email')
        sheet2.write(0, 6, 'Transcript')
        sheet2.write(0, 7, 'GRE')
        sheet2.write(0, 8, 'Template No')
        log_book.save('./Personal_Data/log.xls')

    if gen_type == 1:
        gen_app_temp()
    elif gen_type == 2:
        gen_log_temp()
    elif gen_type == 3:
        gen_app_temp()
        gen_log_temp()

def extract_application():
    # Should return a list of dicts
    # 1. company_name
    # 2. job_title
    # 3. contact_name
    # 4. contact_address
    # 5. recip_email
    # 6. attach transcript
    # 7. attach GRE
    # 8. template no.
    read_book = open_workbook('./application_info.xls')
    r_sheet = read_book.sheet_by_index(0)
    info_list = []
    for row_index in range(1, r_sheet.nrows):
        if len(r_sheet.cell(row_index, 7).value.strip()) == 0:
            temp_no = 1         # Default template No
        else:
            temp_no = r_sheet.cell(row_index, 7).value
        info_list.append( dict(
            company_name    = r_sheet.cell(row_index, 0).value,
            job_title       = r_sheet.cell(row_index, 1).value,
            contact_name    = r_sheet.cell(row_index, 2).value,
            contact_address = r_sheet.cell(row_index, 3).value,
            recip_email     = r_sheet.cell(row_index, 4).value,
            att_trans       = r_sheet.cell(row_index, 5).value,
            att_gre         = r_sheet.cell(row_index, 6).value,
            template_no     = '%d' % temp_no,
            ))
    return info_list

def read_gmail_account():
    # Read Account Info
    f = open('./Personal_Data/gmail_account.txt')
    global username
    username = f.readline().split('=')[1].strip()
    global password
    password = f.readline().split('=')[1].strip()
    global real_name
    real_name = f.readline().split('=')[1].strip()
    global send_from
    send_from = f.readline().split('=')[1].strip()
    f.close()

def render_CL(info):
    fp = open('./Personal_Data/CL_%s.html' % info['template_no'])
    str_data = fp.read()
    fp.close()
    date = datetime.date.today()
    info['date'] = '%s %d, %s' %(date.strftime('%b'), int(date.strftime('%d')), date.strftime('%Y'))
    info['time'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    for key in info:
        str_data = str_data.replace('{%%%s}' % key, str(info[key]))
    msg = MIMEMultipart()
    msg.attach(MIMEText(str_data, 'html'))
    # Attachments
    attach_list = ['Resume.pdf',]
    if info['att_trans'] == 'Y':
        attach_list.append('Transcript.pdf')
    if info['att_gre'] == 'Y':
        attach_list.append('GRE.pdf')
    for attach in attach_list:
        path = './Personal_Data/%s' % attach
        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(path,'rb').read())
        Encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="%s"'
                   % os.path.basename(path))
        msg.attach(part)
    return msg

def gen_log(info):
    # Generate log info
    read_book = open_workbook('./Personal_Data/log.xls')
    r_sheet = read_book.sheet_by_index(0)
    write_book = copy(read_book)
    w_sheet = write_book.get_sheet(0)

    # Copy read_book to write_book, which copys the existing logs
    for row_index in range(r_sheet.nrows):
        for col_index in range(r_sheet.ncols):
            w_sheet.write(row_index, col_index, r_sheet.cell(row_index, col_index).value)

    w_sheet.write(r_sheet.nrows, 0, info['time'])
    w_sheet.write(r_sheet.nrows, 1, info['company_name'])
    w_sheet.write(r_sheet.nrows, 2, info['job_title'])
    w_sheet.write(r_sheet.nrows, 3, info['contact_name'])
    w_sheet.write(r_sheet.nrows, 4, info['contact_address'])
    w_sheet.write(r_sheet.nrows, 5, info['recip_email'])
    w_sheet.write(r_sheet.nrows, 6, info['att_trans'])
    w_sheet.write(r_sheet.nrows, 7, info['att_gre'])
    write_book.save('./Personal_Data/log.xls')

    # Clean up the app_info excel

def sendEmail( recip_email, subject, msg):
    # Every email address should render the CV template and load info
    msg['To'] = recip_email
    msg['Subject'] = subject
    msg['From'] = send_from

    # Different modes:

    # Debug mode, send to 127.0.0.1, normall send
    if debug_mode:
        server = smtplib.SMTP('127.0.0.1')
        server.sendmail(msg['From'], [msg['To'], username], msg.as_string())
        server.quit()
        return
    # Test mode, send to google, but only self
    server = smtplib.SMTP('smtp.gmail.com:587')
    server.starttls()
    server.ehlo()
    server.login(username,password)
    # Always bcc to self
    if args.test:
        server.sendmail(msg['From'], [username,], msg.as_string())
    else:
        server.sendmail(msg['From'], [msg['To'], username,], msg.as_string())
    server.quit()

def main():
    if args.gen:
        print 'Generating Excel templates'
        print 'Please fill in the template and run this app again'
        # Going to clean all
        gen_temp(3)
        return
    read_gmail_account()
    info_list = extract_application()
    for info in info_list:
        msg = render_CL(info)
        subject = 'Application for %s from %s' % (info['job_title'], real_name)
        sendEmail( info['recip_email'], subject, msg)
        gen_log(info)
        if not args.test:
            gen_temp(1)

if __name__ == '__main__':
    print 'Debug mode is %s' % debug_mode
    main()

# Things TO DO
#Done 1. Add template choice
#Done 2. Add Option parse
# 3. Add error control
#Done 4. Add attach
#Done 5. Release the excel rows when the email is sent
#Done 6. Add BCC, now always bcc to self

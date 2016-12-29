from decimal import Decimal
import traceback

import smtplib
from email.mime.text import MIMEText
from email.header import Header

from get_grades import get_sorted_and_ranked_grades


def send_grades(sender_addr, mail_password, subject, grades_record):
    mail_encoding = 'utf-8'
    mail_host = 'mail.fudan.edu.cn'
    mail_smtp_port = 25
    mail_suffix = '@fudan.edu.cn'
    mail_username = sender_addr

    for receiver in grades_record:
        # you can modify the content format for your convenience
        rec_id = str(receiver['id']).split('.')[0]
        grade = str(receiver['grade'])
        rank = str(receiver['rank'])
        percentage = (str(Decimal(receiver['percentage']) * 100) + '%')

        content = '学号: %s\n成绩: %s\n排名: %s\n百分比: %s' % (
            rec_id,
            grade,
            rank,
            percentage
        )

        receiver_addr = rec_id + mail_suffix
        message = MIMEText(content, 'plain', mail_encoding)
        message['From'] = Header(sender_addr, mail_encoding)
        message['To'] = Header(receiver_addr, mail_encoding)
        message['Subject'] = Header(subject, mail_encoding)

        try:
            smtpObj = smtplib.SMTP()
            smtpObj.connect(mail_host, mail_smtp_port)
            smtpObj.login(mail_username, mail_password)
            smtpObj.sendmail(sender_addr, receiver_addr, message.as_string())
            print('success: %s' % (receiver_addr))
        except smtplib.SMTPException as e:
            print('fail: %s' % (receiver_addr))

if __name__ == '__main__':
    sender_addr = input('your fudan email address: ')
    mail_password = input('your fudan email password: ')
    subject = input('your mail title (e.g: 你的某学科总成绩):  ')
    path_to_grades_excel = input('path to grades excel (e.g: grades.xlsx): ')
    id_column = input('column name for student id (e.g: 学号): ')
    grade_column = input('column name for grade (e.g: 成绩): ')
    confirm = input('PLEASE CONFIRM BEFORE SENDING (y/n): ')
    while confirm != 'y' and confirm != 'n':
        confirm = input('FORMAT ERROR. PLEASE CONFIRM BEFORE SENDING (y/n): ')
    if confirm == 'y':
        try:
            print('fetching grades, sorting and ranking them...')
            grades_record = get_sorted_and_ranked_grades(
                path_to_grades_excel, id_column, grade_column)
            print(grades_record)
            print(type(grades_record))
            print('done.\n')
            print('sending grades...')
            send_grades(sender_addr, mail_password, subject, grades_record)
            print('done.')
            print('goodbye.')
        except FileNotFoundError as e1:
            print(
                'something wrong with your path to grade excel. cannot find it.')
        except KeyError as e2:
            print(
                'something wrong with your column name. cannot recognize it.')
        except BaseException as e3:
            print(e3)
            traceback.print_exc()
    else:
        print('goodbye.')

import re
import traceback
import decimal
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
from get_grades import get_sorted_and_ranked_grades


def send_grades(sender_addr, mail_password, sender_name, subject, postscript, grade_records):
    mail_encoding = 'utf-8'
    mail_host = 'mail.fudan.edu.cn'
    mail_smtp_port = 25
    mail_suffix = '@fudan.edu.cn'
    mail_username = sender_addr

    for record in grade_records:
        # you can modify the content format for your convenience
        content = '学号: %s\n成绩: %s\n排名: %s\n排名百分比: %s\n%s' % (
            record['id'],
            record['grade'],
            record['rank'],
            record['percentage'],
            '\n' + postscript + '\n' if (
                (postscript is not None) and (len(postscript) > 0)) else ''
        )

        receiver_addr = record['id'] + mail_suffix
        message = MIMEText(content, 'plain', mail_encoding)
        message['From'] = formataddr((sender_name, sender_addr))
        message['To'] = receiver_addr
        message['Subject'] = Header(subject, mail_encoding)

        try:
            smtpObj = smtplib.SMTP()
            smtpObj.connect(mail_host, mail_smtp_port)
            smtpObj.login(mail_username, mail_password)
            smtpObj.sendmail(sender_addr, receiver_addr, message.as_string())
            print('success: %s' % (receiver_addr))
        except smtplib.SMTPException as e:
            print('FAIL: %s, cause: %s' % (receiver_addr, e))

if __name__ == '__main__':
    sender_addr = input('your fudan email address: ')
    while not re.match(r'^\d+@fudan\.edu\.cn$', sender_addr):
        sender_addr = input(
            'format error. it should be 学号@fudan.edu.cn. please re-enter: ')

    mail_password = input('your fudan email password: ')
    sender_name = input('your name(shown as the sender): ')
    subject = input('your mail title (e.g: 你的某学科总成绩): ')
    postscript = input('your mail postscript(附言) (e.g: 如有任何问题, 请发邮件联系我): ')
    path_to_grades_excel = input('path to grades excel (e.g: grades.xlsx): ')
    id_column = input('column name for student id (e.g: 学号): ')
    grade_column = input('column name for grade (e.g: 成绩): ')
    confirm = input('PLEASE CONFIRM BEFORE SENDING (y/n): ')

    while confirm != 'y' and confirm != 'n':
        confirm = input('FORMAT ERROR. PLEASE CONFIRM BEFORE SENDING (y/n): ')

    if confirm == 'y':
        try:
            print('fetching grades, sorting and ranking them...')
            grade_records = get_sorted_and_ranked_grades(
                path_to_grades_excel, id_column, grade_column)
            print('done.\n')
            print('sending grades...')
            send_grades(
                sender_addr, mail_password, sender_name, subject, postscript, grade_records)
            print('done.')
            print('goodbye.')
        except FileNotFoundError as e1:
            print(
                'something wrong with your path to grade excel. cannot find it.')
        except KeyError as e2:
            print(
                'something wrong with your column name. cannot recognize it.')
        except decimal.InvalidOperation as e3:
            print(
                'something wrong with the grade type or your column name for grade. grade should be number.')
        except BaseException as e3:
            print(e3)
            traceback.print_exc()
    else:
        print('goodbye.')

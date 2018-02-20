# encoding:utf-8
import smtplib
import time
from email.mime.text import MIMEText
from email.header import Header


class SendMail():
    def sendemail(self, sender, receivers,
                  mail_host, mail_user, mail_pass,
                  subject, htmltext):
        message = MIMEText(htmltext, 'html', 'utf-8')
        message['From'] = sender
        message['To'] = receivers
        message['Subject'] = Header(subject, 'utf-8')

        try:
            smtpObj = smtplib.SMTP()
            smtpObj.connect(mail_host, 25)
            smtpObj.login(mail_user, mail_pass)
            smtpObj.sendmail(sender, receivers, message.as_string())

            msg = "Success"
            print(msg)
        except Exception as e:
            msg = "Fail" + str(e)
            print(msg)
        finally:
            smtpObj.quit()

        date = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        with open('history.log', 'a') as log:
            log.write(date + ' ' + msg + '\n')

        return msg

#-*-coding:utf-8-*-

"""
@author: Wu Deyang
@date: 20200705
"""

from time import sleep
from argparse import ArgumentParser
from getpass import getpass
from smtplib import SMTP
from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
from pandas import read_excel

__version__ = "1.0.0.20200705_alpha"

arg_parser = ArgumentParser(description="GPA Calculator")
arg_parser.add_argument("-V", "--version", action="version", version="%(prog)s"+__version__)
arg_parser.add_argument("-g", "--gpa", dest="gpa_file", help="GPA and emails excel file")
arg_parser.add_argument("-m", "--mail", dest="mail_file", default="mail_content.html", help="Mail template file")
arg_parser.add_argument("-u", "--username", dest="username", help="Tsinghua email username")
arg_parser.add_argument("-p", "--password", action="store_true", dest="password", help="Tsinghua email password")
args = arg_parser.parse_args()

class MailSender:
    def __init__(self, gpa_file, mail_file, username, password):
        self.polymer_data = read_excel(gpa_file, sheet_name='高分子')
        self.chem_data = read_excel(gpa_file, sheet_name='化工')
        with open(mail_file, encoding='utf-8') as f:
            self.mail_template = f.read()
        self.username = username
        self.password = password
    
    def mail_send(self):
        from_addr = "{}@mails.tsinghua.edu.cn".format(self.username)
        password = self.password
        smtp_server = "mails.tsinghua.edu.cn"

        server = SMTP(smtp_server, 25)
        server.set_debuglevel(2)
        server.login(from_addr, password)

        for stu_data in [self.polymer_data, self.chem_data]:
            for index in stu_data.index:
                print()
                print("="*20)
                name = stu_data.at[index, '姓名']
                num = stu_data.at[index, '学号']
                to_addr = stu_data.at[index, '邮箱']
                class_num = stu_data.at[index, '班级人数']
                grade_num = stu_data.at[index, '专业人数']
                
                gpa1_term = stu_data.at[index, '最近必限']
                class1_term = stu_data.at[index, '最近必限班级排名']
                grade1_term = stu_data.at[index, '最近必限专业排名']
                gpa2_term = stu_data.at[index, '最近必限任']
                class2_term = stu_data.at[index, '最近必限任班级排名']
                grade2_term = stu_data.at[index, '最近必限任专业排名']

                gpa1_total = stu_data.at[index, '总体必限']
                class1_total = stu_data.at[index, '总体必限班级排名']
                grade1_total = stu_data.at[index, '总体必限专业排名']
                gpa2_total = stu_data.at[index, '总体必限任']
                class2_total = stu_data.at[index, '总体必限任班级排名']
                grade2_total = stu_data.at[index, '总体必限任专业排名']
                
                mail_content = self.mail_template % (
                    name, num, 
                    gpa1_term, class1_term, class_num, grade1_term, grade_num,
                    gpa2_term, class2_term, class_num, grade2_term, grade_num,
                    gpa1_total, class1_total, class_num, grade1_total, grade_num,
                    gpa2_total, class2_total, class_num, grade2_total, grade_num
                )
                
                msg = MIMEText(mail_content, 'html', 'utf-8')
                msg['From'] = self.format_addr(from_addr)
                msg['To'] = self.format_addr('%s <%s>' % (name, to_addr))
                msg['Subject'] = Header("化8年级学分绩与排名情况参考", "utf-8").encode()
                server.sendmail(from_addr, [to_addr], msg.as_string())
                print("="*20)
                print()
                sleep(3)
        server.quit()
    
    def format_addr(self, address):
        name, addr = parseaddr(address)
        return formataddr((Header(name, 'utf-8').encode(), addr))

if __name__ == "__main__":
    if args.password:
        password = getpass("password:")
    s = MailSender(args.gpa_file, args.mail_file, args.username, password)
    s.mail_send()
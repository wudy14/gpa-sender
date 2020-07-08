#-*-coding:utf-8-*-

"""
@author: Wu Deyang
@date: 20200705
"""

from os.path import exists
from time import sleep
from argparse import ArgumentParser
from getpass import getpass
from smtplib import SMTP
from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
from pandas import read_excel, merge, DataFrame, ExcelWriter

__version__ = "1.0.1.20200708_alpha"

arg_parser = ArgumentParser(description="GPA Calculator and Email Sender")
arg_parser.add_argument("-V", "--version", action="version", version="gpa_sender "+__version__)
arg_parser.add_argument("--mode", choices=["gpa", "email"], dest="mode", help="Use gpa calculator or email sender")
# GPA Calculator
arg_parser.add_argument("-f", "--file", dest="grade_file", help="Grade excel file")
arg_parser.add_argument("-t", "--template", dest="gpa_email_template", default="gpa_email_template.xlsx", help="GPA and emails excel file template")
# Email Sender
arg_parser.add_argument("-g", "--gpa", dest="gpa_email_file", help="GPA and emails excel file")
arg_parser.add_argument("-m", "--mail", dest="mail_template", default="mail_template.html", help="Mail template file")
arg_parser.add_argument("-u", "--username", dest="username", help="Tsinghua email username")
arg_parser.add_argument("-p", "--password", action="store_true", dest="password", help="Tsinghua email password")
args = arg_parser.parse_args()

class GPACalculator:
    def __init__(self, grade_file, gpa_email_template):
        if not exists(grade_file):
            raise ValueError("Invalid grade file path " + grade_file)
        if not exists(gpa_email_template):
            raise ValueError("Invalid grade file path " + gpa_email_template)
        self.grade = read_excel(grade_file, sheet_name='report', encoding='gb2312')
        self.gpa_email_template = gpa_email_template
        self.gpa_email_chem = read_excel(gpa_email_template, sheet_name="化工")
        self.gpa_email_poly = read_excel(gpa_email_template, sheet_name="高分子")
        self.gpa_email_chem.loc[:, "学号"] = self.gpa_email_chem.loc[:, "学号"].astype(str)
        self.gpa_email_poly.loc[:, "学号"] = self.gpa_email_poly.loc[:, "学号"].astype(str)
        self.grade.loc[:, "学号"] = self.grade.loc[:, "学号"].astype(str)
        self.terms = self.grade["学年学期"].unique()
        self.terms.sort()
        self.course_types = {"必限":["必修", "限选"], "必限任":["必修", "限选", "任选"]}
        self.gpa = self.grade[["学号", "姓名", "教学班级"]].groupby("学号").first()
    
    def calculate(self):
        for term in self.terms:
            for type_name, course_type in self.course_types.items():
                gpa_index = term + "-" + type_name
                gpa_frame = self.gpa_calculate([term], course_type, gpa_index)
                self.gpa = merge(self.gpa, gpa_frame, on="学号", how="outer")
        
        for type_name, course_type in self.course_types.items():
            gpa_index = "总体" + "-" + type_name
            gpa_frame = self.gpa_calculate(self.terms, course_type, gpa_index)
            self.gpa = merge(self.gpa, gpa_frame, on="学号", how="outer")

    def gpa_calculate(self, term, course_type, gpa_index):
        useful_cols = ["学号", "学分", "绩点成绩"]
        grade_frame = self.grade.loc[(self.grade["学年学期"].isin(term)) & (self.grade["课程属性"].isin(course_type)), useful_cols].copy()
        grade_frame.loc[:, "总绩点"] = grade_frame.loc[:, "绩点成绩"] * grade_frame.loc[:, "学分"]
        grade_frame.dropna("index", how="any", inplace=True)
        gpa_frame = grade_frame.groupby("学号").sum().copy()
        gpa_frame.loc[:, gpa_index] = gpa_frame.loc[:, "总绩点"] / gpa_frame.loc[:, "学分"]
        gpa_frame.loc[:, gpa_index] = gpa_frame.loc[:, gpa_index].round(decimals=2)
        gpa_frame.drop(["学分", "绩点成绩", "总绩点"], axis=1, inplace=True)
        gpa_frame.sort_values(gpa_index, ascending=False, inplace=True)
        gpa_frame.reset_index(inplace=True)
        return gpa_frame

    def sort(self):
        for frame in [self.gpa_email_chem, self.gpa_email_poly]:
            for idx, row in frame.iterrows():
                for type_name, course_type in self.course_types.items():
                    try:
                        frame.loc[idx, "最近"+type_name] = self.gpa.loc[self.gpa["学号"]==row["学号"], self.terms[-1]+"-"+type_name].values[0]
                        frame.loc[idx, "总体"+type_name] = self.gpa.loc[self.gpa["学号"]==row["学号"], "总体-"+type_name].values[0]
                    except(IndexError):
                        raise IndexError("Student {} in \"gpa_email_template.xlsx\" does not exist in \"gpa.xlsx\"".format(row["学号"]))
        
        # 高分子
        grade_num = self.gpa_email_poly.loc[:, "学号"].unique().size
        class_num = grade_num
        self.gpa_email_poly.loc[:, "专业人数"] = grade_num
        self.gpa_email_poly.loc[:, "班级人数"] = class_num
        for term in ["最近", "总体"]:
            for type_name, course_type in self.course_types.items():
                gpa_col = term + type_name
                for scope in ["班级", "专业"]:
                    sort_col = gpa_col + scope + "排名"
                    self.gpa_email_poly.loc[:, sort_col] = self.gpa_sort(self.gpa_email_poly.loc[:, gpa_col].copy())

        # 化工
        grade_num = self.gpa_email_chem.loc[:, "学号"].unique().size
        self.gpa_email_chem.loc[:, "专业人数"] = grade_num
        for term in ["最近", "总体"]:
            for type_name, course_type in self.course_types.items():
                gpa_col = term + type_name
                sort_col = gpa_col + "专业排名"
                self.gpa_email_chem.loc[:, sort_col] = self.gpa_sort(self.gpa_email_chem.loc[:, gpa_col])

        classes = self.gpa_email_chem.loc[:, "教学班级"].unique()
        for c in classes:
            class_num = self.gpa_email_chem.loc[self.gpa_email_chem["教学班级"]==c, "教学班级"].count()
            self.gpa_email_chem.loc[self.gpa_email_chem["教学班级"]==c, "班级人数"] = class_num
            for term in ["最近", "总体"]:
                for type_name, course_type in self.course_types.items():
                    gpa_col = term + type_name
                    sort_col = gpa_col + "班级排名"
                    gpa_series = self.gpa_email_chem.loc[self.gpa_email_chem["教学班级"]==c, gpa_col].copy()
                    self.gpa_email_chem.loc[self.gpa_email_chem["教学班级"]==c, sort_col] = self.gpa_sort(gpa_series)
    
    def gpa_sort(self, grade_series):
        sort_frame = DataFrame(columns=["gpa"])
        sort_frame.loc[:, "gpa"] = grade_series
        sort_frame.sort_values("gpa", ascending=False, inplace=True)
        sort_frame.loc[:, "sorted"] = list(range(1, grade_series.size+1))
        for i, idx in enumerate(sort_frame.index):
            if i == 0:
                last_idx = idx
            else:
                if sort_frame.loc[idx, "gpa"] == sort_frame.loc[last_idx, "gpa"]:
                    sort_frame.loc[idx, "sorted"] = sort_frame.loc[last_idx, "sorted"]
                last_idx = idx
            
        sort_frame.sort_index(inplace=True)
        return sort_frame.loc[:, "sorted"].copy()

    def output_generate(self):
        print("Save gpa to gpa.xlsx")
        self.gpa.to_excel("gpa.xlsx", index=False)
        print("Done!")
    
    def gpa_email_update(self):
        print("Save gpa and ranking to gpa_email.xlsx")
        writer = ExcelWriter("gpa_email.xlsx")
        self.gpa_email_chem.to_excel(writer, index=False, sheet_name="化工")
        self.gpa_email_poly.to_excel(writer, index=False, sheet_name="高分子")
        writer.save()
        print("Done!")

class MailSender:
    def __init__(self, gpa_email_file, mail_template, username, password):
        if not exists(gpa_email_file):
            raise ValueError("Invalid grade file path " + gpa_email_file)
        if not exists(mail_template):
            raise ValueError("Invalid grade file path " + mail_template)
        self.polymer_data = read_excel(gpa_email_file, sheet_name='高分子')
        self.chem_data = read_excel(gpa_email_file, sheet_name='化工')
        with open(mail_template, encoding='utf-8') as f:
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
    if not args.mode:
        raise ValueError("No mode chosen! Enter \"gpa_sender.exe -h\" for help.")
    if args.mode == "gpa":
        if not args.grade_file:
            raise ValueError("No grade file chosen!")
        c = GPACalculator(args.grade_file, args.gpa_email_template)
        c.calculate()
        c.output_generate()
        c.sort()
        c.gpa_email_update()
    elif args.mode == "email":
        if not args.gpa_email_file:
            raise ValueError("No gpa file chosen!")
        if not args.username:
            raise ValueError("No username given!")
        if args.password:
            password = getpass("password:")
        else:
            raise ValueError("No password given!")
        s = MailSender(args.gpa_email_file, args.mail_template, args.username, password)
        s.mail_send()

#-*-coding:utf-8-*-

"""
@author: Wu Deyang
@date: 20200705
"""

from os.path import exists
from argparse import ArgumentParser
from pandas import read_excel, merge, DataFrame, ExcelWriter

__version__ = "1.0.0.20200705_alpha"

arg_parser = ArgumentParser(description="GPA Calculator")
arg_parser.add_argument("-V", "--version", action="version", version="%(prog)s"+__version__)
arg_parser.add_argument("-f", "--file", dest="grade_file", help="Grade excel file")
arg_parser.add_argument("-o", "--output", dest="output_file", default="gpa.xlsx", help="Ouput excel file")
arg_parser.add_argument("-g", "--gpa", dest="gpa_file", default="gpa_email_template.xlsx", help="GPA and emails excel file")
args = arg_parser.parse_args()

class GPACalculator:
    def __init__(self, grade_file, output_file, gpa_file):
        if not exists(grade_file):
            raise ValueError("Invalid grade file path " + grade_file)
        if not exists(gpa_file):
            raise ValueError("Invalid grade file path " + gpa_file)
        self.grade = read_excel(grade_file, sheet_name='report', encoding='gb2312')
        self.gpa_file = gpa_file
        self.gpa_email_chem = read_excel(gpa_file, sheet_name="化工")
        self.gpa_email_poly = read_excel(gpa_file, sheet_name="高分子")
        self.gpa_email_chem.loc[:, "学号"] = self.gpa_email_chem.loc[:, "学号"].astype(str)
        self.gpa_email_poly.loc[:, "学号"] = self.gpa_email_poly.loc[:, "学号"].astype(str)
        self.grade.loc[:, "学号"] = self.grade.loc[:, "学号"].astype(str)
        self.output_file = output_file
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
                    frame.loc[idx, "最近"+type_name] = self.gpa.loc[self.gpa["学号"]==row["学号"], self.terms[-1]+"-"+type_name].values[0]
                    frame.loc[idx, "总体"+type_name] = self.gpa.loc[self.gpa["学号"]==row["学号"], "总体-"+type_name].values[0]
        
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
            if i > 0:
                if sort_frame.loc[idx, "gpa"] == sort_frame.loc[last_idx, "gpa"]:
                    sort_frame.loc[idx, "sorted"] = sort_frame.loc[last_idx, "sorted"]
            last_idx = idx
            
        sort_frame.sort_index(inplace=True)
        return sort_frame.loc[:, "sorted"].copy()

    def output_generate(self):
        self.gpa.to_excel(self.output_file, index=False)
    
    def gpa_email_update(self):
        writer = ExcelWriter("gpa_email.xlsx")
        #print(self.gpa_email_chem)
        self.gpa_email_chem.to_excel(writer, index=False, sheet_name="化工")
        self.gpa_email_poly.to_excel(writer, index=False, sheet_name="高分子")
        writer.save()
        
if __name__ == "__main__":
    c = GPACalculator(args.grade_file, args.output_file, args.gpa_file)
    c.calculate()
    c.output_generate()
    c.sort()
    c.gpa_email_update()
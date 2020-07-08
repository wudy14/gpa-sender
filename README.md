# GPA Sender

## 安装

测试环境Python版本：3.7.1

```shell
pip install -r requirements.txt
```

## 文件列表

```
.
+-- gpa_sender.py  //主程序
+-- 2019年秋全部成绩查询.xls  //各科成绩查询样例
+-- gpa_email_test.xlsx  //邮件发送测试程序
+-- gpa_email_template.xlsx  //学生邮箱
+-- mail_template.html  //发送邮件模板
```

## 学分绩计算

(1) 修改 `gpa_email_template.xlsx` 的学生名单及邮箱，最终邮件发送名单以此为准

(2) 运行程序

在主程序所在文件夹打开命令提示符，输入：

```shell
python gpa_sender.py --mode gpa -f GRADE_FILE  //各科成绩查询文件
python gpa_sender.py --mode gpa --file GRADE_FILE  //各科成绩查询文件
```

程序计算输出各学期学分绩及总体学分绩到 `gpa.xlsx`，输出所有同学学分绩及排名到 `gpa_email.xlsx`
`gpa_email.xlsx` 为最终要发送的文件，发送前最后再检查一遍。

## 邮件发送

(1) 修改邮件模板 `mail_template.html`，内容可自定义

(2) 运行程序

在主程序所在文件夹打开命令提示符，输入：

```shell
python gpa_sender.py --mode email -g GPA_EMAIL_FILE -u USERNAME -p  //成绩排名与邮箱文件，清华邮箱用户名
python gpa_sender.py --mode email --gpa GPA_EMAIL_FILE --username USERNAME --password  //成绩排名与邮箱文件，清华邮箱用户
```

程序将直接读取成绩排名与邮箱文件发送，正式发送前务必检查 ```gpa_email.xlsx```，可先使用 ```gpa_email_test.xlsx``` 做测试，注意修改测试文件邮箱。

## 其他

(1) `python gpa_sender.py -V` 或 `python gpa_sender.py --version` 输出版本号;

(2) `python gpa_sender.py -h` 或 `python gpa_sender.py --help` 输出帮助信息.

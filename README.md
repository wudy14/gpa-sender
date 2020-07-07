# GPA Sender

## 文件列表

```
.
+-- gpa_sender.exe          //主程序
+-- 2019年秋全部成绩查询.xls //各科成绩查询样例
+-- gpa_email_test.xlsx     //邮件发送测试程序
+-- gpa_email_template.xlsx //学生邮箱
+-- mail_content.html       //发送邮件模板
```

## 学分绩计算

(1) 修改 `gpa_email_template.xlsx` 的学生名单及邮箱，最终邮件发送名单以此为准
(2) 运行程序
在主程序所在文件夹打开命令提示符，输入：
```shell
gpa_sender.exe --mode gpa -f (各科成绩查询文件)
gpa_sender.exe --mode gpa --file (各科成绩查询文件)
```
程序计算输出各学期学分绩及总体学分绩到 `gpa.xlsx`，输出所有同学学分绩及排名到 `gpa_email.xlsx`
`gpa_email.xlsx` 为最终要发送的文件，发送前最后再检查一遍

## 邮件发送

(1) 修改邮件模板 `mail_content.html`，内容可自定义
(2) 运行程序
在主程序所在文件夹打开命令提示符，输入：
```shell
gpa_sender.exe --mode email -g (成绩及排名文件) -u (清华邮箱用户名) -p
gpa_sender.exe --mode email --gpa (成绩及排名文件) --username (清华邮箱用户名) --password
```
程序将直接读取成绩排名文件发送，正式发送前务必检查 ```gpa_email.xlsx```，可先使用 ```gpa_email_test.xlsx``` 做测试，注意修改测试文件邮箱。

## 其他

(1) `gpa_sender.exe -V` 或 `gpa_sender.exe --version` 输出版本号
(2) `gpa_sender.exe -h` 或 `gpa_sender.exe --help` 输出帮助信息

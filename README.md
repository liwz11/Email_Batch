# Email_Batch

Dependencies: chardet-3.0.4, python-dateutil-2.7.3


Uasage:

```
python email_poplib.py <email_addr> <password> [options]

options:

 -t start_time        Start time of the email you want to check, with a format 
                      of YYYY-mm-dd HH:MM:SS. (OPTIONAL)
 -k keyword           Keyword in subject of the email you want to check. (OPTIONAL)

Example:

python email_poplib.py test@mails.tsinghua.edu.cn test123456 '2018-09-25 00:00:00' '中期材料'

```
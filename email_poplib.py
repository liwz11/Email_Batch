import sys, os, time, chardet
from datetime import datetime

import poplib, email
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr

def format_filename(str):
    str = str.replace('/', '+').replace(':', '+').replace('*', '+')
    str = str.replace('?', '+').replace('<', '+').replace('>', '+')
    str = str.replace('\\', '+').replace('"', '+').replace('|', '+')
    return str

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        if charset == 'gb2312':
            charset = 'gb18030'
        value = value.decode(charset)
    
    if isinstance(value, bytes):
        value = value.decode()
    
    return value

def get_email_headers(msg):
    headers = {}
    for header in ['From', 'To', 'Cc', 'Subject', 'Date']:
        value = msg.get(header, '')
        if value:
            if header == 'Date':
                headers['Date'] = value
            if header == 'Subject':
                headers['Subject'] = decode_str(value)
            if header == 'From':
                name, addr = parseaddr(value)
                name = decode_str(name)
                from_addr = u'%s <%s>' % (name, addr)
                headers['From'] = from_addr
            if header == 'To':
                all_to = value.split(',')
                to = []
                for x in all_to:
                    name, addr = parseaddr(x)
                    name = decode_str(name)
                    to_addr = u'%s <%s>' % (name, addr)
                    to.append(to_addr)
                headers['To'] = ','.join(to)
            if header == 'Cc':
                all_cc = value.split(',')
                cc = []
                for x in all_cc:
                    name, addr = parseaddr(x)
                    name = decode_str(name)
                    cc_addr = u'%s <%s>' % (name, addr)
                    cc.append(to_addr)
                headers['Cc'] = ','.join(cc)
    return headers
 
def get_email_attachments(msg, savepath):
    if os.path.exists(savepath):
        return 'FILE_EXISTS'
    
    attachments = []
    for part in msg.walk():
        filename = part.get_filename()
        if filename:
            filename = decode_str(filename)
            filename = format_filename(filename)
            
            attachments.append(filename)
            
            if not os.path.exists(savepath):
                os.mkdir(savepath)
            filepath = os.path.join(savepath, filename)
            if not os.path.exists(filepath):
                data = part.get_payload(decode = True)
                f = open(filepath, 'wb')
                f.write(data)
                f.close()
    
    return attachments

if __name__ == '__main__':
    poplib._MAXLINE = 40960

    # 账户信息
    user = sys.argv[1]
    pwd = sys.argv[2]
    pop3_server = user.split('@')[1]
    
    t1_str = '2000-01-01 00:00:00'
    if len(sys.argv) > 3:
        t1_str = sys.argv[3]
    t1_date = datetime.strptime(t1_str, '%Y-%m-%d %H:%M:%S')
    t1 = int(time.mktime(t1_date.timetuple()))
    
    keyword = ''
    if len(sys.argv) > 4:
        keyword = sys.argv[4]

    # 连接到POP3服务器
    server = poplib.POP3_SSL(pop3_server)
    server.set_debuglevel(0)
    print(server.getwelcome())
    print('\n')

    # 身份认证
    server.user(user)
    server.pass_(pwd)

    mails_count, mails_size = server.stat()
    print('mails_count =', mails_count)
    print('mails_size =', mails_size, 'bytes')
    print('\n')
    
    # 创建附件目录
    attach_path = os.path.join(os.getcwd(), 'attach')
    if not os.path.exists(attach_path):
        os.mkdir(attach_path)
    
    for i in range(1, mails_count + 1):
        res, byte_lines, size = server.retr(i)
        
        # 转码，拼接邮件内容
        str_lines = []
        for x in byte_lines:
            #print(chardet.detect(x))
            if chardet.detect(x)['encoding'] == 'GB2312':
                str_lines.append(x.decode('GBK'))
            else:
                str_lines.append(x.decode())
        mail_content = '\n'.join(str_lines)
        
        # 解析邮件
        msg = Parser().parsestr(mail_content)
        headers = get_email_headers(msg)
        
        # 过滤
        if 'Date' not in headers:
            continue
        if 'Subject' not in headers:
            continue
        
        timestr = headers['Date'].split(' (')[0]
        timedate = datetime.strptime(timestr, '%a, %d %b %Y %H:%M:%S %z')
        timestamp = int(time.mktime(timedate.timetuple()))
        
        if timestamp < t1:
            continue
        if keyword != '' and keyword not in headers['Subject']:
            continue
        
        print('')
        print(i, '/', mails_count)
        
        # 打印邮件头
        print('Subject:', headers['Subject'])
        print('From:', headers['From'])
        if 'To' in headers:
            print('To:', headers['To'])
        if 'Cc' in headers:
            print('Cc:', headers['Cc'])
        if 'Date' in headers:
            print('Date:', headers['Date'])
        
        #下载附件
        subject = headers['Subject']
        subject = format_filename(subject)
        
        timestamp = str(timestamp)
        savepath = os.path.join(attach_path, subject + '_' + timestamp)
        attachments = get_email_attachments(msg, savepath)
        print('Attachments:', attachments)
    
    print('')
    server.quit()
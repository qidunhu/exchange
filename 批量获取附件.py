#运行该脚本需要安装python exchangelib库，pip install exchangelib
from exchangelib import DELEGATE, Account, Credentials, Message, Mailbox, HTMLBody

def Email_Attachment(username,password,Folder_name, File_path):
    #连接邮箱所需信息
    creds = Credentials(
        username=username,
        password=password
    )
    account = Account(
        primary_smtp_address=username,
        credentials=creds,
        autodiscover=True,
        access_type=DELEGATE
    )

    #迭代收件箱下的文件夹中的附件，并保存到本地
    for item in account.inbox.children:
        if item.name==Folder_name:#只保存Folder_name参数文件夹下的附件
            index=0
            totalcount=0
            page=0
            while True:          
                for model in item.all()[page:page+50]:
                    index=index+1
                    for attachment in model.attachments:
                        if isinstance(attachment,type(attachment)):
                            with open(File_path + attachment.name, 'wb') as f:
                                f.write(attachment.content)
                if totalcount==index:
                    break
                page=page+50
                totalcount=index

            
Email_Attachment('ceshi2@genecast.com.cn','XXXXXX','test','/Users/qi.dunhu/Desktop/files/')

#参数说明:
#Email函数需要传递如下四个参数
#1、uername为邮箱地址,如qi.dunhu@genecast.com.cn
#2、password为邮箱密码
#3、Folder_name为要保存的邮箱文件夹名称,如要获取收件箱下的test文件夹里所有附件,此处需要填写test
#4、File_path为要将附件保存到本地的路径

#exchangelib库说明连接https://pypi.org/project/exchangelib/   https://github.com/ecederstrand/exchangelib

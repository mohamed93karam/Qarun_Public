
import os
from exchangelib import *
import datetime

import webbrowser
import time


path = fr'D:\Daily Report'

credentials = Credentials('mkarm', 'passwordHere')
config = Configuration(server='webmail.qarun.net', credentials=credentials)

account = Account('mkarm@qarun.net', credentials=credentials, config=config)

def main():
    # print(account.root.tree())

    end_date = datetime.datetime.now()

    start_date = end_date - datetime.timedelta(days=20)
    delta = datetime.timedelta(days=1)
    while start_date <= end_date:
        print(start_date.strftime("%Y-%m-%d"))
        save_today_mail(start_date)
        start_date += delta
    os.startfile('C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE')



    

def save_today_mail(today):



    # nextMonth = today + datetime.timedelta(days=31)

    pathToSave = fr'\{today.year}\{today.strftime("%m")}-{today.strftime("%b")}\{today.strftime("%d")}'
    print(path + pathToSave)



    for item in account.inbox.filter(subject__contains=f'EXCEL PRODUCTION MORNING REPORT {today.strftime("%d-%b-%y")}').order_by('datetime_received'):
        local_path = os.path.join(path + pathToSave, item.subject + ".eml")
        if not os.path.exists(path + pathToSave):
            os.makedirs(path + pathToSave)

            if os.path.exists(local_path):
                    continue
        with open(local_path, 'wb') as file:
            file.write(item.mime_content)

        print(item.subject, item.sender, item.datetime_received, item.body)
        for attachment in item.attachments:
            print (attachment.name)
            if isinstance(attachment, FileAttachment):
                local_path = os.path.join(path + pathToSave, attachment.name)
                if not os.path.exists(path + pathToSave):
                    os.makedirs(path + pathToSave)

                if os.path.exists(local_path):
                    continue
                with open(local_path, 'wb') as f:
                    f.write(attachment.content)
                print('Saved attachment to', local_path)
    #     item.soft_delete()
    #     for item in avocetReport:
    #         item.soft_delete()
    #     return False
    # return True
    

main()

import win32com.client as win32
import os
import shutil

def loginOutlook():
    # Prompt the user for login credentials
    email = "rpabotadminuat@sudlife.in"
    password = "Sud@4321"
    try:
        # Log in to Outlook with the provided credentials
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
        outlook.Logon(email, password, False, True)
        return outlook
    except Exception as e:
        print("Failed to log in to Outlook.")
        print(e)
        return None

def readPassword(save_folder):
    outlook = loginOutlook()
    print(outlook)
    file1.write(str(outlook)+"\n")
    if not outlook:
        return

    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items
    count = 0
    

    # Traverse messages in reverse order
    for i in range(messages.Count, 0, -1):
        message = messages.Item(i)
        print(message)
        if message.UnRead:
            # Process attachments if available
            attachments = message.Attachments
            if attachments.Count > 0:
                for attachment in attachments:
                    file_name = attachment.FileName
                    # Check if the attachment is a PDF or TIFF file
                    if file_name.lower().endswith('.csv') :
                        count += 1
                        received_time = message.ReceivedTime
                        # date_str = received_time.strftime('%Y%m%d') 
                        # time_str = received_time.strftime('%H%M%S')
                        date_str1 = received_time.strftime('%Y%m%d') 
                        time_str1= received_time.strftime('%H%M%S')
                        New_filename=(os.path.splitext(file_name))[0]+"_"+date_str1+time_str1+".csv"
                        New_filename=New_filename.replace(" ","_")
                        # New_filename=str(file_name.split(".")[0])+"_"+date_str+time_str+".csv"
                        # date_time_folder = os.path.join(save_folder, date_str, time_str)
                        if not os.path.exists(save_folder):
                            os.makedirs(save_folder)
                        # if not os.path.exists(path):
                        #     os.makedirs(path)
                        save_path = os.path.join(save_folder, New_filename)
                        attachment.SaveAsFile(save_path)
                        q=f"Saved .csv attachment from '{message.subject}' DateTime:{date_str1}_{time_str1}: {file_name} to {save_folder}"
                        file1.write(q+"\n")
                        print(q)
                        if not os.path.exists(path):
                            os.makedirs(path)
                        shutil.copy(save_path,path2)
                        shutil.move(save_path,path)
                        print(f"{New_filename}----->moved to desired path")
                        file1.write(f"{New_filename}----->moved to desired path\n")

        message.UnRead = False
    print(f"Found {count} unread email(s) and processed their attachments.\n")
    file1.write(f"Found {count} unread email(s) and processed their attachments.\n")

import os.path
from datetime import datetime,timedelta
import time

try:
    now = datetime.now() 
    logdate=str(now.strftime("%Y%m%d_%H%M%S"))
    log_path = r"E:\Mail\mail_Logs"
    if not os.path.exists(log_path):
            os.mkdir(log_path)
    completeName = log_path+"\\mailLogs_"+logdate+'.txt'
    file1 = open(completeName, "a")
    file1.write(str(now)+"\n")
    save_folder = r"E:\Mail\attachments"
    path=r"\\10.1.22.86\t1day_flat_temp"
    path2=r"\\10.1.22.86\EMAIL_Audit"
    readPassword(save_folder)
    file1.close()
except Exception as e:
    print(e)
    file1.write(str(e)+"\n")

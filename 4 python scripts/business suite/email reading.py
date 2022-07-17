
# outlook must be on

# library for system manipulation
import os
import pandas
import shutil
from time import sleep

# library for outlook manipulation
import win32com.client as client

# using absolute path because of win32com library ? unclear
reports_dir = "m_scripts" + os.sep + "reports"
try:
    shutil.rmtree(reports_dir)
except:
    print("error deleting reports directory structure")
    sleep(3)
sleep(3)
os.mkdir(reports_dir)

# start outlook instance
outlook = client.Dispatch("Outlook.Application")

# reading
namespace = outlook.GetNamespace("MAPI")

# get default folders
# drafts = namespace.GetDefaultFolder(16)
# inbox = namespace.GetDefaultFolder(6)

# get main account folder
account_folder = namespace.Folders['Mladen.Ilic@ncr.com']
# get custom folders
custom_folder = account_folder.Folders['AAAAA_Auto_Opportunities']

# folders methods
# folder_name = custom_folder.Name
# message_count = custom_folder.Items.Count
# parent_folder = custom_folder.Parent.Name

# get message
# concrete_email_i = custom_folder.Items[0]
# contrete_email_ii = custom_folder.Items(1)

# message methods
# sender_name = concrete_email_i.SenderName
# sender_adress = concrete_email_i.SenderEmailAddress
# subject = concrete_email_i.Subject
# body = concrete_email_i.Body

# loop
for message in custom_folder.Items:
    print(message.Subject)
    # print(message.Body)

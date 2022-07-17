
# outlook must be on
# set cc recipient list
# set subject and body text
# set <br>
# set configuration file
# set reports folder for per customer attachements
# set source folder for general attachments
# disable counter for testing
subject_text = " - SSQ1 2022 update resume"
body_text = """In regard to the latest hold for SSQ1 2022, we are glad to inform you that our Software Engineering provided solution and that we are good to proceed.<br>
<br>
Actions to be taken:<br>
<br>
1.	We would need to upgrade DCS agent to the latest version 3.45.1 (dcsdocs-CSMWindowsAgent-190422-1110-130). <br>
a.	This is non-intrusive action which doesn’t require reboot and will be done in background without any interference with the unit functionality.<br>
<br>
2.	Deployment of SSQ1 2022<br>
3.	The hotfix (Common-XFS-Hotfix-W1594_PSSCM-15686).<br>
<br>
Release notes are attached.<br>
<br>
We would like to use one unit as pilot to start with.<br>
<br>
Unit __________ proposal.<br>
<br>
In case you want to change unit please let us know. In addition, please find entire schedule in attachment<br>
<br>
<br>
"""

import os
# library for outlook manipulation
import win32com.client as client
# import customer configuration
from customers import masterdict
# remove generic customer
del masterdict["generic"]
# start outlook instance
outlook = client.Dispatch("Outlook.Application")

# using absolute path because of win32com library ? unclear
reports_dir = "m_scripts" + os.sep + "reports"
source_dir = "m_scripts" + os.sep + "source"

# # counter for testing !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# count = 0 # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

# loop through customer configuration
for customer_name in masterdict:

    # # counter for testing !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    # if count > 0: # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    #     break # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    # count = count + 1 # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

    # customer dict
    customer_dict = masterdict[customer_name]
    # create mail item
    message = outlook.CreateItem(0)
    # interactive display in outlook
    message.Display()
    # message components

    # list of recipients
    to = []
    cc = ["SN230200@ncr.com", "MN230117@ncr.com", "Mladen.Ilic@ncr.com", "Dusan.Djordjevic@ncr.com", "US.ATMIncMgt@ncr.com"]
    dear = []

    # get list of recipients names & emails
    bank_data = customer_dict["bank data"]

    if bank_data["lifecycle"] == "yes":

        bank_contacts = bank_data["bank contacts"]
        for contact in bank_contacts:
            to.append(bank_contacts[contact])
            if not contact.startswith("generic"):
                contact_name = contact.split()[0]
                dear.append(contact_name)

        account_support = bank_data["account support"]
        for contact in account_support:
            cc.append(account_support[contact])

    elif bank_data["lifecycle"] == "no":

        project_trasition_manager = bank_data["project transition manager"]
        for contact in project_trasition_manager:
            to.append(project_trasition_manager[contact])
            if not contact.startswith("generic"):
                contact_name = contact.split()[0]
                dear.append(contact_name)

    # convert to input strings
    to = "; ".join(to)
    cc = "; ".join(cc)
    
    dear_count = len(dear)

    if dear_count == 1:
        dear = ", ".join(dear)
        dear = "Hello " + dear + ","

    elif dear_count > 1:
        pop = dear.pop()
        dear = ", ".join(dear)
        dear = "Hello " + dear + " and " + pop + ","

    subject = customer_name + subject_text

    message.To = to
    message.CC = cc
    # message.BCC = "mi250175@ncr.com"
    message.Subject = subject
    # plain text body
    # message.Body = "Wish you a happy birthday!"
    # html formated body
    html_body = """
        <body>
            <main>
                <p style="margin:0; padding:0; border:0">{}</p><br>
                <p style="margin:0; padding:0; border:0">{}</p><br>
                <p style="margin:0; padding:0; border:0">Regards,</p><br>
                <h5 style="margin:0; padding:0; border:0">
                    Mladen Ilic<br>SWD Specialist<br>NAMER Software Distribution &<br>Global Endpoint Security - Managed Services
                </h5>
                <h6 style="margin:0; padding:0; border:0">
                    NCR Corporation<br>msn: <a href="mi250175@ncr.com">mi250175@ncr.com</a><br>
                    Phone:+<br><a href="mladen.ilic@ncr.com">mladen.ilic@ncr.co</a> | <a href="www.ncr.com">www.ncr.com</a>
                </h6>
                <p style="margin:0; padding:0; border:0">
                    <img src="C:/Users/mi250175/OneDrive - NCR Corporation/Desktop/scripts/sig_files/image001.png"<br>
                </p>
                <p style="padding:0; margin:0; border:0; color:gray;">
                    <b style="margin-right:50px">NCR Social Media:</b>
                    <a href="https://www.linkedin.com/company/ncr-corporation">
                        <img src="C:/Users/mi250175/OneDrive - NCR Corporation/Desktop/scripts/sig_files/image002.png" 
                            style="padding-right:50px"></a>
                    <a href="https://www.facebook.com/ncrcorp">
                        <img src="C:/Users/mi250175/OneDrive - NCR Corporation/Desktop/scripts/sig_files/image003.png" 
                            style="padding-right:50px"></a>
                    <a href="https://twitter.com/NCRCorporation">
                        <img src="C:/Users/mi250175/OneDrive - NCR Corporation/Desktop/scripts/sig_files/image004.png" 
                            style="padding-right:50px"></a>
                    <a href="https://plus.google.com/102373146691782027099/posts">
                        <img src="C:/Users/mi250175/OneDrive - NCR Corporation/Desktop/scripts/sig_files/image005.png" 
                            style="padding-right:50px"></a>
                    <a href="https://www.youtube.com/user/ncrcorporation">
                        <img src="C:/Users/mi250175/OneDrive - NCR Corporation/Desktop/scripts/sig_files/image006.png" 
                            style="padding-right:50px"></a>
                </p>
                <p style="margin:0; padding:0; border:0">
                    <i>“Client satisfaction is our top Priority”</i>
                </p>
            </main>
        </body>
    """
    # using template without parameters
    # message.HTMLBody = html_body
    # using template with parameters, add parameters in {}
    message.HTMLBody = html_body.format(dear, body_text)

    # get customer specific files for attachment
    for subdir, dirs, files in os.walk(reports_dir):
        for filename in files:
            filepath = subdir + os.sep + filename

            # remove file extension .xlsx -5 characters
            size = len(filename)
            shorted_filename = filename[:size - 5]

            #print("customer name ", customer_name)
            #print("new filename ", shorted_filename)
            #print()

            # check if match with customer
            if shorted_filename.startswith(customer_name):

                # get absolute path
                abs_file_path = os.path.abspath(filepath)

                # add attachement
                message.Attachments.Add(abs_file_path)

    # get generic files for attachment
    for subdir, dirs, files in os.walk(source_dir):
        for filename in files:
            filepath = subdir + os.sep + filename

            # get absolute path
            abs_file_path = os.path.abspath(filepath)

            # add attachement
            message.Attachments.Add(abs_file_path)

    # save as draft
    message.Save()

    # send message
    # message.Send()

    # close message window
    message.Close(0)

    # delete message
    # message.Delete()

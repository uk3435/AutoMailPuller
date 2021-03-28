import win32com.client
import os
import datetime as dt
import glob
import os
import time


def mail_oku():

  

    outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6).Folders["{This sub folder under INCOMING folder}"]  #6 indicates INCOMING fFolder
    #inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items


    date_time = dt.datetime.now()
    lastHourDateTime = dt.datetime.now() - dt.timedelta(minutes=10) #10 minutes is set as a sample. 
    messages_ = messages.Restrict("[ReceivedTime] >= '" +lastHourDateTime.strftime('%m/%d/%Y %H:%M %p')+"'")

    #messages__=messages_.Restrict("@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0037001E LIKE '%test%'") #To: 0E04001E
                                                                                                                #Cc: 0E03001E
                                                                                                                #Bcc: 0E02001E
                                                                                                                #Subject: 0037001E
                                                                                                                #Sender's email address: 0C1F001E
                                                                                                                #Sender's name: 0C1A001E
                                                                                                                # Above codes are special gift for my followers :) I searched them lot :)
                                                                                                                # In the example I used 0x0037001E which is Mail's subject.
              
    messages__=messages_.Restrict("@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0C1F001E LIKE '%{write word(s) in subject that are searching}%'")
    #message=messages.GetLast()

    for message in messages__:
        
        message_=message.body
        print (message_) # It is up to you what you will do for these mails. I put print method. You can save them a file or smth else.
  
  

mail_oku() #mail_oku method is Turkish name which is read_mail :)
         
    






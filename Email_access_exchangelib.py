import os
os.environ["PATH"] += os.pathsep + "Kerberos Installation Directory"
from exchangelib import Credentials , Account,Configuration,DELEGATE,EWSDateTime, EWSTimeZone,Q
from os.path import basename
from os import listdir
from os.path import isfile, join
from datetime import timedelta

def connect():
    
    global s,save_path,pyfolder,local_path
    s=0
    server = 'outlook.office365.com'
    email = os.environ['username']
    username = os.environ['username']
    password = os.environ['password']

    try:
        creds = Credentials(username=username, password=password)
        config = Configuration(server=server, credentials=creds)
        account = Account(primary_smtp_address=email, autodiscover=False, config = config, access_type = DELEGATE)
        
        

        try:            
                       
            tz = EWSTimeZone.localzone()
            now = EWSDateTime.now().replace(tzinfo=tz)
            thirty_minutes_ago = now - timedelta(minutes=20)
            folder = account.inbox


            # Use the following syntax to filter the emails as per the mentioned subject
            last_email = folder.filter(subject__contains='123',datetime_received__gt=thirty_minutes_ago).order_by('-datetime_received')


            # If you have option to choose between 2 subject use the following syntax:
            last_email = folder.filter(Q(subject__icontains='ABC') | Q(subject__icontains='XYZ'),datetime_received__gt=thirty_minutes_ago).order_by('-datetime_received') 
            count_folder = last_email.count()
        
            if count_folder >= 1:
                
               local_path = "Directory where you wish to download the attachments"
               for item in last_email[:1]:
                    for attachment in item.attachments:     
                        local_path = os.path.join(local_path, attachment.name)
                                
                        with open(local_path, 'wb') as f:
                            f.write(attachment.content)

            else:
                print("No email found with the mentioned fiters")
                            
        
        except Exception as e:
          print(f"Failed to Fetch Emails from EWS server due to error {e}")  
           


    except Exception as e:
        print(f"Unable to connect to EWS server due to error {e}")        
        exit()
        
connect()
    
        
          
 

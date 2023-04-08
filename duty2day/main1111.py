import csv
import smtplib, ssl
from schedule import every, repeat, run_pending
import time
#$ pip install schedule
from datetime import datetime
print("This scrip is sending the message to duty specialist.")



result = []
duty_2_day = ""
location = ""

date = datetime.now().strftime("%d.%m.%Y")
print(date)
message = {'PAB': """\
    Subject: You are duty at PAB

    Hello Sir! You are duty today at PAB!""", 
    'MCR': """\
    Subject: You are duty at MCR

    Hello Sir! You are duty today at MCR!""",
    'MA': """\
    Subject: You are duty at MA

    Hello Sir! You are duty today at MA!"""}


with open('duty_list.csv', newline='') as file:
    duty_list = csv.reader(file, delimiter=',')
    for duty in duty_list:
        result.append(duty)


#@repeat(every(2).minutes)   
@repeat(every().day.at("07:30")) 
def who_is_duty(): 
    
    
    for duty in result:
        if date == duty[0]:
            location = duty[1]            
            index = duty.index('D')            
            duty_2_day = result[0][index]
  
      
    print(duty_2_day)
    print(location)
    smtp_server = "smtp.office365.com"
    port = 587  # For starttls
    sender_email = "IT.KBL.Duty2day@outlook.com"
    password = "Kazmin23"
    # input("Type your password and press enter: ")

    # Create a secure SSL context
    context = ssl.create_default_context()

    # Try to log in to server and send email
    try:
        server = smtplib.SMTP(smtp_server,port)
        server.ehlo() # Can be omitted
        server.starttls(context=context) # Secure the connection
        server.ehlo() # Can be omitted
        server.login(sender_email, password)
        # TODO: Send email here
        
        for receiver in duty_2_day.split(' '):
            server.sendmail(sender_email, receiver, message.get(location))        
    except Exception as e:
        # Print any error messages to stdout
        print(e)
    finally:
        server.quit()
        
        
def main():    
    
    
    who_is_duty()
    
    #schedule.every().day.at("07:30").do(send_msg)
    
    while True:
        run_pending()
        print("The message was send.")
        time.sleep(1)
    
if __name__ == '__main__':
    main()
        




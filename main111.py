import smtplib, ssl
import schedule
import datetime

#$ pip install schedule
    
emails = ("serik.jaibergenov@kazminerals.com",
"faizkhan.assankhanov@kazminerals.com",
"maxim.abylkassov@kazminerals.com",
"nurzhan.saktaganov@kazminerals.com",
"a.shadmanov@kazminerals.com",
"zhenis.borbassov@kazminerals.com",
"anton.vlassenko@kazminerals.com",
"IT.KBL.Duty2day@outlook.com",
"grigoriy.zhilyayev@kazminerals.com",
"maxim.rychkin@kazminerals.com",
"ivan.konanykhin@kazminerals.com",
"viktor.bobko@kazminerals.com"
)


locations = ("""\
Subject: You are duty at PAB

Hello Sir! You are duty today at PAB!""", 
"""\
Subject: You are duty at MCR

Hello Sir! You are duty today at MCR!""",
"""\
Subject: You are duty at MA

Hello Sir! You are duty today at MA!""")

print("This scrip is sending the message to duty specialist.")


def send_msg():
    
    day = datetime.datetime.today().day

    receiver_email = ""
    receiver_email2 = ""
    message = ""

    if day == 1 or day == 6 or day == 11 or day == 16 or day == 27:
        receiver_email = emails[0]
        receiver_email2 = emails[1]
        
        if day == 1 or day == 16:
            message = locations[0]
        elif day == 6 or day == 27:
            message = locations[2]
        else:
            message = locations[1]
        
    elif day == 2 or day == 7 or day == 12 or day == 17 or day == 22:
        receiver_email = emails[2]
        receiver_email2 = emails[3]
        
        if day == 2 or day == 17:
            message = locations[1]
        elif day == 7 or day == 22:
            message = locations[0]
        else:
            message = locations[2]
        
    elif day == 3 or day == 8 or day == 13 or day == 18 or day == 23 or day == 28:
        receiver_email = emails[4]
        receiver_email2 = emails[5]
        
        if day == 3 or day == 18:
            message = locations[2]
        elif day == 8 or day == 23:
            message = locations[1]
        else:
            message = locations[0]
        
    elif day == 9 or day == 14 or day == 19 or day == 24 or day == 29:
        receiver_email = emails[6]
        receiver_email2 = emails[7]
        
        if day == 9 or day == 24:
            message = locations[2]
        elif day == 14 or day == 29:
            message = locations[1]
        else:
            message = locations[0]
        
    elif day == 4 or day == 15 or day == 20 or day == 25 or day == 30:
        receiver_email = emails[8]
        receiver_email2 = emails[9]
        
        if day == 4 or day == 25:
            message = locations[0]
        elif day == 15 or day == 30:
            message = locations[2]
        else:
            message = locations[1]
        
    elif day == 5 or day == 10 or day == 21 or day == 26 or day == 31:
        receiver_email = emails[10]
        receiver_email2 = emails[11] 
        
        if day == 5 or day == 26:
            message = locations[1]
        elif day == 10 or day == 31:
            message = locations[0]
        else:
            message = locations[2]

  

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
        server.sendmail(sender_email, receiver_email, message)
        server.sendmail(sender_email, receiver_email2, message)
        print("The message was send.")
    except Exception as e:
        # Print any error messages to stdout
        print(e)
    finally:
        server.quit()
    

def main():
    #schedule.every(2).minutes.do(send_msg)
    schedule.every().day.at("07:30").do(send_msg)
    
    while True:
        schedule.run_pending()
    
if __name__ == '__main__':
    main()
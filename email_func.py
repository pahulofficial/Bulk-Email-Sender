import smtplib
def email_send_func(_to,_subj,_msg,_from,_pass):
    # print(_to,_subj,_msg,_from,_pass)
    session = smtplib.SMTP("smtp.gmail.com",587) #Created a Session
    session.starttls() #Start the TCP 
    session.login(_from,_pass)
    msg= "Subject: {}\n\n{}".format(_subj,_msg)
    session.sendmail(_from,_to,msg)

    m = session.ehlo()

    if m[0]==250:
        return "Sent"
    else:
        return "Failed"
    
    session.close()
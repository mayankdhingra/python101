import smtplib

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.login("email id", "password") #email id and password

message='I Love You 3000'

def repeat(a,n):
	i=1
	org_msg=a
	while i<n:
		a=a + '\n' + org_msg
		i+=1
	return(a)

# def repeat(a,n):
    #print((((a)+'\n')*n)[:-1])
    # b = (a+'\n')*n

message = repeat(message,3000) #message to be printed, number of times it is to be printed 

server.sendmail("sender email id","recipient_email id",message) #sender email id as above

import os
from dotenv import load_dotenv
import win32com.client
from termcolor import cprint

load_dotenv()
input_user = ""
sender_email = os.getenv("email")
phone_number = os.getenv("phone")
subject = "Internship Application"
body = \
f"""Dear hiring manager,

I am Gilles De Praeter, a student of HoGent in Belgium. I am studying for a Bachelor's degree in Applied Information Technology: System and Network administration. Next schoolyear, 2022-2023, I have to look for an internship for the second semester, more specifically from the end of February until the end of may. I am looking for an internship as a network and/or system administrator in British Columbia, Canada. I was wondering if your company has any open positions for internships during that period?

I think I could bring contributions to the team and make a worthy candidate for an internship because of my knowledge I gained during my two years in college. In the attachments, I have uploaded my resume and a file called "Curriculum", which contains my curriculum of the full Bachelor's degree. I have completed every course of the first two pages and I will complete every course of the third page before my internship. During my two years, I have studied some programming languages like Java, JavaScript, Python, PowerShell, Bash, Cisco IOS and SQL. I have also studied the basics of web development including HTML5 and CSS, computer science concepts like virtualisation, a big part of CCNAv7 of Cisco, lots of bash scripting using Vagrant and VirtualBox and other useful information.

I am looking for an internship in British Columbia, Canada, because I would like to live there in the future. This would be the perfect opportunity for me to experience what life is really like over there. It would also grant me opportunities in the future and allow me to set worldwide career goals. It would open doors for me of which I could only dream about.

During my two years at Hogent, I have maintained an average of 77%. This number reflects that I have a lot of dedication towards my studies and shows that I am truly passionate about what I study. I would love to continue exercising my passion at your company and work on a project with my full dedication.

I look forward to sharing more details of my experience and motivations with you. Thank you for your consideration.

Sincerely,

Gilles De Praeter
{sender_email}
{phone_number}
"""

outlook = win32com.client.Dispatch('outlook.application')

while(input_user != "q"):
    receiver_email = input("To who do you want to send an email: ")
    mail = outlook.CreateItem(0)
    mail.To = receiver_email
    mail.Subject = subject
    mail.Body = body
    mail.Attachments.Add(os.getenv("PathResume"))
    mail.Attachments.Add(os.getenv("PathCurriculum"))
    mail.Send()
    cprint("Email has been sent!", "green")
    input_user = input("Would you like to send another email? Press 'q' to quit: ")
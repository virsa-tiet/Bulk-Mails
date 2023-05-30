# CREATE A EXCEL SHEET WITH NAMES A DN MAIL HEADING

import pandas as pd
import smtplib,time


your_email = "mail"
your_password = "key"


server = smtplib.SMTP_SSL('smtp.gmail.com',465)
server.ehlo()
server.login(your_email, your_password)

# Read the file
email_list = pd.read_excel("newlist.xlsx") #list final wali

# Get all the Names, Email Addreses, Subjects and Messages
all_names = email_list['NAME']
all_emails = email_list['EMAIL']
all_messages ="We hope this email finds you well. On behalf of VIRSA, the Annual cultural fest of Thapar Institute of Engineering and Technology, we feel proud to present to you 'VIRSA MELA and SHAAM-E-VIRASAT (the cultural eve) 2023'. We believe that your brand would be an excellent match for this event, and we would be honoured to have you amongst us as our collaborating partner.\n\nVIRSA aims at promoting the rich culture of Punjab by conducting awe-inspiring pre-fest and main events with nationwide publicity and pious religious celebrations. After an incredible and successful run of online events, Virsa is back again with the most awaited events. It is an extravaganza combination of the two-day event starting with the 'Shaam-e-virasat' and followed by the highly anticipated 'Virsa Mela'.\n\n'Shaam-e-virasat' on 19th April, is a showcase of the Bhangra competition, Giddha, a traditional fashion walk and a singing performance by a renowned artist of Punjab. This event attracts a lot of participation from across the country and is a perfect platform for your brand to showcase itself to a wide and diverse audience, with Satinder Sartaj as a potential celebrity for the night.\n\n'Virsa Mela' on 23rd April, is a proper depiction of the Punjabi traditional fair with mesmerizing performances and vibrant stalls which are visited by the entire Thapar campus. With over 10,000 in-campus audiences and around 700 spectators online, VIRSA provides an excellent platform to showcase your products and services to a massive and diverse audience.\n\nAs a partner, you'll receive maximum exposure to the audience thanks to the event's reach and coverage in various esteemed newspapers and media. We have a reach of over 50k on our Instagram and Facebook official Virsa handles. Besides this, a mail is sent to all Thaparians and faculty before the event, where again the sponsor is highlighted.\n\nIn addition to the above, we have a team of 200 individuals, each with 500-600 followers on social media, who will share posters of the event and our sponsors, providing even greater reach for your brand. In our last event 'VIRSA TALKS', we collaborated with personalities like Gurpreet Guggi, Rana Ranbir, Bir Singh, Jarnail Singh, Anmol avatar, and many more.\n\nBy partnering with us, your brand will be associated with a culturally significant event that has a significant impact on the student community. We would be honoured to have your brand join us as a partner for VIRSA MELA and CULTURAL EVE 2023.\n\nDon't miss this excellent opportunity to showcase your brand to a diverse and enthusiastic audience. We look forward to hearing back from you soon.\n\nThank you for your time and consideration.\n https://drive.google.com/file/d/1HCyRiLTjhEU40h1IlucDiTst7DwPB3X9/view?usp=share_link \n\nBest regards,\n\nTarik Bhateja\nOverall Events Co-Ordinator\n+91 84276 99163\n\nParampreet Kaur \nOverall Events Co-Ordinator\n+91 7814657866\n\nGurpartap Singh\nOverall Students Co-Ordinator\n+91 7901810523"
message = all_messages
subject = 'Sponsorship and Collaboration for VIRSA fest,Thapar Institute of Engineering and Technology'
# Loop through the emails
c=0
for idx in range(len(all_emails)):
    c+=1

    # Get each records name, email, subject and message
    name = all_names[idx]
    email = all_emails[idx]

    msg=f"Dear {name},\n"


    # Create the email to send
    full_email = ("From: {0} <{1}>\n"
                  "To: {2} <{3}>\n"
                  "Subject: {4}\n\n"
                  "{5}{6}"
                  .format("SSA Virsa", your_email, name, email, subject, msg, message))#name change

    # In the email field, you can add multiple other emails if you want
    # all of them to receive the same text
    try:
        server.sendmail(your_email, [email], full_email)
        print(f'Email to {name},{email} successfully sent!')
        time.sleep(1)
    except Exception as e:
        print('Email to {} could not be sent--------------------------------------------------------------because {}'.format(name, str(e)))

# Close the smtp server
server.close()

print(c)



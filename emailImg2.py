import os
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
import qrcode
from PIL import Image
import pandas as pd
import cv2

html= """\
<html>
    <head></head>
    <body>
 <p>Greetings of the day!</p>
<p>The day you have been waiting for has arrived. </p>
<p>
   <b>
   Venue:
   </b>
   Fete Area
</p>
<p>
   <b>
   The slot allotted to you: 
   </b>1
</p>
<p>
   <b>
   Timings:
   </b> 6:00 P.M. - 9:00 P.M.  
</p>
<p>To avoid confusion entry for boys will be from the H-hostel intersection and for girls from COS entrance. No confusion regarding the same should happen during the main day.</p>
<p>The following are the guidelines that should be strictly followed on the main day: 
<ul>
   <li>
      Adhere to the slot timings allotted in the mail. Entry would not be allowed otherwise.
   </li>
   <li> Carrying the TFF pass issued to you is mandatory.</li>
   <li>
      The QR code attached to this mail should be downloaded beforehand.
   </li>
</ul>
</p>
<p>To avoid any obstacles upon entry and to help us maintain the event's smooth flow keep this mail handy when you arrive on the main day.</p>
<p>Regards,
   </br>
   Team **** 
</p>
    </body>
</html>
"""


# no reply 8
your_email="#######"
your_password="########"

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)
  
email_list = pd.read_excel('C:/Users/hp/Downloads/Slot2TFF.xlsx',engine='openpyxl')

names = email_list['NAME']
emails = email_list['EMAIL']
rollnos=email_list['ROLLNO']

counter=0

Logo_link = 'logo.png'
 
logo = Image.open(Logo_link)
 
# taking base width
basewidth = 100
wpercent = (basewidth/float(logo.size[0]))
hsize = int((float(logo.size[1])*float(wpercent)))
logo = logo.resize((basewidth, hsize), Image.ANTIALIAS)
QRcode = qrcode.QRCode(
    version=1,
    error_correction=qrcode.constants.ERROR_CORRECT_H,
    box_size=14,
    border=2,
)
 
s = smtplib.SMTP_SSL('smtp.gmail.com', 465)
s.ehlo()
s.login(your_email, your_password)

for i in range(len(emails)):
    name = names[i]
    email = emails[i]
    rollno = rollnos[i]
    data="Name" + " : " + name + " \n " + "Roll No. :"  + str(rollno) + " \n " + "Slot : 2" 
    QRcode.add_data(data)
 
    # generating QR code
    QRcode.make()
    
    # saving QR code
    # adding color to QR code
    QRimg = QRcode.make_image(
        fill_color='black', back_color="white").convert('RGB')
    pos = ((QRimg.size[0] - logo.size[0]) // 2,
       (QRimg.size[1] - logo.size[1]) // 2)
    QRimg.paste(logo, pos)
    QRimg.save('gfg_QR'+str(counter)+'.png')
    QRcode.clear()
    # img=cv2.imread('gfg_QR'+str(counter)+'.png')
    # for every record get the name and the email addresses
    
    
    img_data=open('gfg_QR'+str(counter)+'.png','rb').read()
    msg = MIMEMultipart()
    msg['Subject'] = 'Instructions and Invitation for TFF Main Day'
    msg['From'] = 'Thapar Food Festival'
    msg['To'] = email
    #paste message text in this
    
    part2 = MIMEText(html, 'html')
    msg.attach(part2)
    image = MIMEImage(img_data,name="QrCode")
    msg.attach(image)

  
    # # the message to be emailed
    # message = "Hello " + name + "Your QR code is attached" 
   
  
    # sending the email
    s.sendmail("anagpal2_be20@thapar.edu", email, msg.as_string())
    print("Email sent to " + email)
    counter=counter+1

print(counter) 
# close the smtp server
server.close()
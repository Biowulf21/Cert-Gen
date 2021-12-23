from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from openpyxl import Workbook, load_workbook
import smtplib
from email.message import EmailMessage
import imghdr


# send the Email
def sendMessage(subject, body, receipient_email, receipient_name, image):
    sendcount = 0
    sender = 'gdscxavier.exrel@gmail.com'
    password = 'gdscexrel@21'
   # all = len(receipients)
    # print(str(all) + " recipients found.\n")
    sent = 0
    failedTo = []

    print(
        f'Receiver name is: {receipient_name}, Receiver Email is: {receipient_email}')

    sendTo = receipient_email
    message = EmailMessage()
    message['Subject'] = subject
    message['From'] = sender
    message['To'] = sendTo
    message.set_content(body)
    sendcount += 1
    with open(f'Certificates/{receipient_name}.png', 'rb') as f:
        img_attachment = f.read()
        file_type = imghdr.what(f.name)
        file_name = f.name
    # adding image attachment to the email_message object
    message.add_attachment(img_attachment, maintype='image',
                           subtype=file_type, filename=file_name)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(sender, password)
            print('has logged in hehe')
            smtp.send_message(message)
            print(f'sent to {receipient_name}')

            sent = sent+1
    except Exception as e:
        print(f'failed with error {e}')
        failedTo.append(sendTo)


def advisory(receipient_email, receipient_name, image):
    subject = "Tech-Knowledge-G: Cyber Security Certificate"
    body = "Good day, GDSC Member! \n\n" \
           "Thank you for joining our first ever GDSC Xavier Tech-Knowlegde-G Seminar. We would like to show our appreciation by presenting you your certificate of participation for the event as we really appreciate your time and effort in joining. \n\n" \
        "We hope to see you in our upcoming events next year.\n\n" \
           "Thank you again and Merry Christmas!\n\n" \
           "Warmest Regards, \n" \
           "James Jilhaney, GDSC Xavier Chief Technology Officer"

    sendMessage(subject=subject, body=body, receipient_email=receipient_email,
                receipient_name=receipient_name, image=image)


wb = load_workbook('names.xlsx')
ws = wb.active
list_of_names = []
list_of_emails = []
directory = {}
# gets the length of the rows of the file to limit the iteration below
row_count = ws.max_row
print(f"max row count is: {row_count}")

# iterates over all the rows and returns a tuple
for row in ws.iter_rows(min_row=2, max_row=row_count):
    # open the certificate template
    image = Image.open('participants.png')
    draw = ImageDraw.Draw(image)
    # choose font and input the Path as well as size of it
    font = ImageFont.FreeTypeFont(
        '/storage/Python/Certificate/opensans.ttf', size=60)
    # draw the name of the participant to the template based on the x and y axis of image
    draw.text(xy=(725, 680), text='{}'.format(
        row[1].value), fill=(255, 255, 255), font=font)
    image.save('Certificates/{}.png'.format(row[1].value))

    # picks a new image file that matches the name of the
    imageSaved = Image.open(f'Certificates/{row[1].value}.png')

    print(
        f"Email is: {row[0].value}, Name is: {row[1].value}\n")

    advisory(receipient_email=row[0].value,
             receipient_name=row[1].value, image=imageSaved)

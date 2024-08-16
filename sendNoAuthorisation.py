import win32com.client as win32
from utilities import read_excel
import os


DEFAULT_MESSAGE = ("{zwrot}\nW załączniku przesyłam wniosek o udostępnienie "
                   "informacji publicznej.\n\nZ poważaniem\n{imie} {nazwisko}"
                   "\n{tytul} Stowarzyszenia Młoda Lewica RP")

def send_all_emails():
    user_data, towns_data = read_excel()
    # Create an instance of the Outlook application
    outlook = win32.Dispatch('outlook.application')

    # Get pdfs
    files = os.listdir('debug/output/')

    for town_data in towns_data:
        # Create a new mail item
        mail = outlook.CreateItem(0)

        # Set email attributes
        mail.To = town_data[6]
        mail.Subject = 'Wniosek o udostępnienie informacji publicznej'
        message = DEFAULT_MESSAGE
        if town_data[4].upper() == "M":
            message = message.replace("{zwrot}", "Szanowny Panie,")
        elif town_data[4].upper() == "K":
            message = message.replace("{zwrot}", "Szanowna Pani,")
        message = message.replace("{imie}", user_data[0])
        message = message.replace("{nazwisko}", user_data[1])
        message = message.replace("{tytul}", user_data[3])
        mail.Body = message

        # Choose the sending account
        mail.SentOnBehalfOfName = user_data[4]

        # Attach a file
        attachment_path = os.path.abspath(__file__).replace("send.py", f'debug/output/{town_data[0]}.pdf')
        mail.Attachments.Add(attachment_path)

        # Send the email
        mail.Send()

        print("Email sent successfully.")

if __name__ == "__main__":
    send_all_emails()

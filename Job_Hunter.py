import PySimpleGUI as sg

#---------------------- 
#imports for sending emails
import smtplib
import ssl
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.multipart import MIMEMultipart  
from email.mime.base import MIMEBase 
from email import encoders
import time
  
# Add some color
# to the window
sg.theme('SandyBeach')     

#Define the layout and items to be included in the windows
#each item has its own key to be accessed and retrieve the data
layout = [
    [
    [sg.Text("Please enter your Name, Email, Email's Password", font=('', 15, ''))],
    [sg.Text('Name', size =(15, 1)), sg.InputText(key='-name-')],
    [sg.Text('Email', size =(15, 1)), sg.InputText(key='-email-')],
    [sg.Text('Password', size =(15, 1)), sg.InputText(key='-password-', password_char='*')],
            [sg.Text("")],
    ],
    
    [
        
    [sg.Text("Please enter Receiver's Email, Vacancy Name\nSeperate with , for multiple entries", font=('', 15, ''))],
    [sg.Text('Email', size =(15, 1)), sg.InputText(key='-remail-')],
    [sg.Text('Vacancy', size =(15, 1)), sg.InputText(key='-vacancy-')],
        
    ],
    [
        
    [sg.Text("Please select the following", font=('', 15, ''))],
    
        
    ],
    [

        sg.Text("Resume PDF"),
        sg.In(size=(24, 1), enable_events=True, key="-cv-"),
        sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),)),

    ],
    [

        sg.Text("Cover Letter",size=(10,1)),
        sg.In(size=(24, 1), enable_events=True, key="-cl-"),
        sg.FileBrowse(file_types=(("Text Files", "*.txt"),)),

    ],
    [sg.Text("Progress",size=(10,2)),sg.ProgressBar(50, orientation='h', size=(20, 20), border_width=1, key='progbar',bar_color=['#009900', '#FFFFFF'])]
    ,
    [sg.Submit(key='submit'), sg.Cancel(key="exit")]
]

#Initiate window
window = sg.Window('Job Hunter', layout)


#Run the application with a while loop
while True:
    event, values = window.read()
    
    #close the application if window is closed or cancelled is clicked
    if event == "exit" or event == sg.WIN_CLOSED:
        break
    
    #Extract data from the fields when submit is clicked
    if event == "submit":
        sender_name = values['-name-']
        sender_email = values['-email-']
        password = values['-password-']
        #split emails with ,
        remails = values['-remail-'].split(", ")
        #split jobs with ,
        job_vacancies = values['-vacancy-'].split(", ")
        filename = values['-cv-']
        c_l = values['-cl-']
        with open(c_l, encoding='utf8') as c:
            cover = c.read()
        
        #not a necessary  step but to manipulate the email list
        df = remails
        print("these are the emails", df)
        
        start = 0
        count = 1
        for x in range(len(df)):
            name = 'none' 
            position = 'none' 
            email = df[x]
            f_position = job_vacancies[x]
            if start == 0:
                start += 1
                print(' ')
                print('Program initiated...')
                time.sleep(2)
                print('Will send an email to companies in the following order -->',email)
                send_email = "Y"
                if send_email == 'N':
                    print('Program interrupted!')
                    break

            print("Sending email to "+ email + '...')
            # Configurating user's info
            msg = MIMEMultipart()

            if name.lower() == 'none':
                msg['To'] = formataddr((email, email))
            else:
                msg['To'] = formataddr((name, email))

            msg['From'] = formataddr((sender_name, sender_email))
            msg['Subject'] = '6 years experienced person seeking a job'

            if position.lower() == 'none':
                msg.attach(MIMEText(cover.format(x = f_position)))
            else:
                msg.attach(MIMEText(cover.format(x = position)))

            try:
                # Open PDF file in binary mode
                with open(filename, "rb") as attachment:
                        part = MIMEBase("application", "octet-stream")
                        part.set_payload(attachment.read())
                  # Encode file in ASCII characters to send by email
                encoders.encode_base64(part)
                  # Add header as key/value pair to attachment part
                part.add_header(
                    "Content-Disposition",
                    f"attachment; filename= {sender_name}_CV.pdf",
                  )
                msg.attach(part)
            except Exception as e:
                print(f'Oh no! We didn\'t found the attachment!\n{e}')
                break
            try:
                # Creating a SMTP session | use 587 with TLS, 465 SSL and 25
                server = smtplib.SMTP('smtp.gmail.com', 587)
                # Encrypts the email
                context = ssl.create_default_context()
                server.starttls(context=context)
                # We log in into our Google account
                server.login(sender_email, password)
                # Sending email from sender, to receiver with the email body
                server.sendmail(sender_email, email, msg.as_string())
                print('Email sent!')
                #a count system to stop showing "to_continue_or_not" at the last email in the list.
            except Exception as e:
                print(f'Oh no! Something bad happened!\n{e}')
                break
            finally:
                print(f'Closing the server for email number {count} ...\n')
                count += 1
                val=0
                k=len(df)
                for i in range(k):
                    val=val+100/(k-i)    
                    window['progbar'].update_bar(val)
                sg.Popup('Progress Done,\nGood Luck Hunting a Job', title="DONE",keep_on_top=True)

                server.quit()
        print('---Sending Emails Done---'.upper())
            
        


window.close()

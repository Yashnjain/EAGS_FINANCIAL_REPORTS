import smtplib
import email.mime.multipart
import email.mime.text
import email.mime.base
import email.encoders as encoders
import bu_config
import os


from bu_alerts import __version__

__author__ = "chetanbiourja"
__copyright__ = "chetanbiourja"
__license__ = "mit"

def send_mail(
    receiver_email: str, 
    mail_subject: str, 
    mail_body: str, 
    attachment_location: str = None, 
    sender_email: str = None, 
    sender_password: str=None, 
    multiple_attachment_list: list = None
    ) -> bool:
    """The Function responsible to do all the mail sending logic.

    Args:
        sender_email (str): Email Id of the sender.
        sender_password (str): Password of the sender.
        receiver_email (str): Email Id of the receiver.
        mail_subject (str): Subject line of the email.
        mail_body (str): Message body of the Email.
        attachment_location (str, optional): Absolute path of the attachment. Defaults to None.

    Returns:
        bool: [description]
    """
    done = False
    try:
        # config = bu_config.get_config(process_name="BU_ALERTS", table_name= "BU_CONFIG_PARAMS")
        if not sender_email or sender_password:
            sender_email = "biourjapowerdata@biourja.com"
            # sender_password = r"bY3mLSQ-\Q!9QmXJ"
            sender_password = "Texas08642"
        receivers = receiver_email.split(",")
        msg = email.mime.multipart.MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = mail_subject
        body = mail_body
        msg.attach(email.mime.text.MIMEText(body, 'html'))

        if attachment_location:
            with open(attachment_location, 'r') as attachment:
                # instance of MIMEBase and named as p
                p = email.mime.base.MIMEBase('application', 'octet-stream')
                # To change the payload into encoded form
                p.set_payload((attachment).read())
                encoders.encode_base64(p)  # encode into base64
                p.add_header('Content-Disposition',
                             "attachment; filename= %s" % attachment_location)
                msg.attach(p)  # attach the instance 'p' to instance 'msg'
        if multiple_attachment_list:
            for f in multiple_attachment_list:
                path, file_name = os.path.split(f)
                binary_file = open(f, 'rb')
                
                try:
                    payload = email.mime.base.MIMEBase('application', 'octate-stream', Name=file_name)
                except:
                    payload = email.mime.base.MIMEBase('application', 'octet-stream', Name=file_name)
                payload.set_payload((binary_file).read())
                #enconding the binary into base64
                encoders.encode_base64(payload)
                payload.add_header('Content-Decomposition', 'attachment', filename=file_name)
                msg.attach(payload)

        # s = smtplib.SMTP('smtp.gmail.com', 587) # creates SMTP session
        s = smtplib.SMTP('smtp.office365.com',
                         587) # creates SMTP session
        s.starttls()  # start TLS for security
        s.login(sender_email, sender_password)  # Authentication
        text = msg.as_string()  # Converts the Multipart msg into a string

        s.sendmail(sender_email, receivers, text)  # sending the mail
        s.quit()  # terminating the session
        done = True
        print("Email sent successfully.")
    except Exception as e:
        print(
            f"Could not send the email, error occured, More Details : {e}")
    finally:
        return done

if __name__ == "__main__":
    send_mail(receiver_email='chetan.surwade@biourja.com',mail_subject='test',mail_body='')
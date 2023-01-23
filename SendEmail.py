from email.mime.image import MIMEImage
import os
import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

class SendMail:
    emailSubject = None
    senderMailId = None
    sendMailPassword = None
    smtpServerDomainName = None
        
    def __init__(self, senderMailId, sendMailPassword, emailSubject):
        self.port = 465
        self.sendMailPassword = sendMailPassword
        self.smtpServerDomainName = "smtp.gmail.com"
        self.senderMailId = senderMailId
        self.emailSubject = emailSubject

    @classmethod
    def fromConfig(cls, gmailSenderId, gmailSenderPassword, gmailEmailSubject):
        return cls(gmailSenderId, gmailSenderPassword, gmailEmailSubject)


    def send(self, donarName, receiverEmailId, donarReceiptFilePath):
        sslContext = ssl.create_default_context()
        smptServer = smtplib.SMTP_SSL(self.smtpServerDomainName, self.port, context=sslContext)
        smptServer.login(self.senderMailId, self.sendMailPassword)

        print("Sending Email To = "+receiverEmailId)
        ##
        emailBody = self.getEmailBody(donarName)
        mail = MIMEMultipart()
        mail['To'] = receiverEmailId
        mail['From'] = self.senderMailId
        mail['Subject'] = self.emailSubject
        
        html_content = MIMEText(emailBody, 'html')
        mail.attach(html_content)

        ##
        attachmentFileName = donarReceiptFilePath
        ##attachment = open(donarReceiptFilePath, "rb")
        ##attachmentMimeBase = MIMEBase('application', 'octet-stream')
        ##attachmentMimeBase.set_payload((attachment).read())
        ##encoders.encode_base64(attachmentMimeBase)
        ##attachmentMimeBase.add_header('Content-Disposition', "attachment; filename= %s" % attachmentFileName)
        ##mail.attach(attachmentMimeBase)
     
        fp = open(donarReceiptFilePath, 'rb')
        attachmentMimeBase = MIMEBase('application', 'octet-stream')
        attachmentMimeBase.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(attachmentMimeBase)
        attachmentMimeBase.add_header('Content-Disposition', 'attachment',
                                      filename=os.path.basename(donarReceiptFilePath))
        mail.attach(attachmentMimeBase)

        ## 
        smptServer.sendmail(self.senderMailId, receiverEmailId, mail.as_string())
    

    def getEmailBody(self, donarName) :
        htmlContent = """<p>Dear <strong>%s</strong>,</p>
        <p>Thank you for your donation. YOur generosity is appreciated. Please refer to the attached document for donation details.</p>
        <p>Sincerly<br />"""% (donarName)
        return htmlContent
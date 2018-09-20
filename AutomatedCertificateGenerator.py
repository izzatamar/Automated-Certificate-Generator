# coding: utf-8

# In[ ]:

from reportlab.lib import colors
import pandas as pd
from pandas import DataFrame
from reportlab.lib.pagesizes import A4, inch, landscape
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4, portrait
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm, inch
import csv
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER
from reportlab.pdfbase.pdfmetrics import stringWidth
import datetime

import io
import urllib.request
import time
import ftputil


downloaded=[]




def sendEmail():
    email_user = 'email user'
    email_password = 'email password'
    email_send = 'target email'

    subject = 'subject'

    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = email_send
    msg['Subject'] = subject

    body = 'Hi there, sending this email from Python!'
    msg.attach(MIMEText(body, 'plain'))

    filename = 'LEE KA SOON(AQAD).pdf'
    attachment = open(filename, 'rb')

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= " + filename)

    msg.attach(part)
    text = msg.as_string()
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(email_user, email_password)

    # server.sendmail(email_user, email_send, text)
    attachment.close()
    server.quit()


while True:
    skipped=0
    

    with ftputil.FTPHost('ftp hostname','ftp username', 'ftp password') as ftp:






        #ftp = ftputil.FTPHost('ftp.nfeconsulting.com','amar@nfeconsulting.com', '!@#123!@#123')
        currentDir = ftp.getcwd()
        print('Current Directory: ', ftp.getcwd())

        listDir = []
        print('Initiating Main function')




        filematch = '*csv'

        data = []
        filelist=[]
        filePath=[]
        fileName=[]
        fileRoot=[]

        listDir = []

        current_time = time.time()
        for f in os.listdir():
            creation_time = os.path.getctime(f)

            print((current_time - creation_time) // (24 * 3600))
            if (current_time - creation_time) // (24 * 3600) >= 1 and f.endswith(('.pdf')):
                print(os.listdir())
                os.unlink(f)
                print('{} removed'.format(f))

        print('initiating ftp.walk')
        #for root,dirs,files in ftp.walk(currentDir,topdown=True, onerror=None):
            #print('filename in ftp.walk: ',files)
        for root, dirnames, filenames in ftp.walk(currentDir):
            for filename in filenames:
                if filename.endswith(('.csv')):
                    #print('listing .csv files', filename)
                    fileName.append(filename)
                    fileRoot.append(root)
                    filePath.append(ftp.path.join(root, filename))

                if filename in downloaded:
                    print(filename + 'File is already downloaded. Skip.')
                if filename not in downloaded and filename.endswith(('.csv')):
                    #print('GLOBAL ROOT: ', globalRoot)



                    print('Getting ' + filename)
                    print('Current ROOT: ', root)
                    width, height = A4

                    #print('listDir: ', listDir)
                    #print('current root: ', root)
                    print('current dirnames: ', dirnames)
                    #print('current filenames: ',filenames)




                    defaultDownloadDir = os.getcwd()
                    #print('Default download directory: ',defaultDownloadDir)
                    #print('Current Directory: ', currentDir)
                    #print('filename:',filename)
                    #print('ftp.download path: ', defaultDownloadDir +'/'+ filename)

                    print('downloading: ' ,root+filename, 'into: ' , defaultDownloadDir)


                    ftp.download(root+'/'+filename,defaultDownloadDir+'/'+filename)
                    print('file downloaded: ', root+ '/'+filename)




                    with open(defaultDownloadDir +'/'+filename, 'r') as f:

                        data = pd.read_csv(f, delimiter=';')


                        rows = 0


                        labelName = [['CustomerName'], ['CertificateNo'], ['SupplierName'], ['ContractDate'], ['Amount']]

                        elements = []
                        i = 0
                        styles = getSampleStyleSheet()
                        styleN = styles["BodyText"]

                        for index, row in data.iterrows():
                            name = data.loc[index, 'name']
                            if name + '.pdf' not in downloaded:
                                print('Contents in Downloaded',downloaded)
                                compileData = []
                                name = data.loc[index, 'name']
                                certificateNo = data.loc[index, 'certificateNo']
                                print('CertificateNo at index ' + str(index) + 'is ' + certificateNo)

                                identification = data.loc[index, 'identification']
                                valueDate = data.loc[index, 'valuedate']
                                commodity = data.loc[index, 'commodities']
                                location = data.loc[index, 'location']
                                storage = data.loc[index, 'storagefacility']
                                quantity = data.loc[index, 'quantity']
                                purchaseRef = data.loc[index, 'purchaseRef']
                                asagentfor = data.loc[index, 'asagentfor']
                                currency = data.loc[index, 'currency']
                                price = data.loc[index, 'price']

                                compileData.append(commodity)
                                compileData.append(location)
                                compileData.append(storage)
                                compileData.append(quantity)
                                # compileData=list(compileData)
                                print('asdjaskldjasiodjasiodjaisodj', compileData)
                                # compileData2=pd.DataFrame(compileData)
                                compileData2 = []
                                compileData2 = [['Commodity', 'Location', 'Storage Facility', 'Quantity (MT)'],
                                                [commodity, location, storage, quantity]]
                                # compileData2=list(compileData)
                                print('compileData: ', compileData)
                                c = canvas.Canvas(name + '(AQAD).pdf', pagesize=A4)

                                c.setTitle("HOLDING CERTIFICATE")

                                s = getSampleStyleSheet()

                                s = s["Normal"]

                                labelName = [['Commodity'], ['Location'], ['Storage Facility'], ['Quantity (MT)']]

                                print('compileDATA2: ', compileData2)
                                labelName = pd.DataFrame(labelName)
                                labelName = labelName.transpose().values.tolist()

                                style = TableStyle([('ALIGN', (1, 1), (-2, -2), 'RIGHT'),
                                                    ('TEXTCOLOR', (1, 1), (-2, -2), colors.red),
                                                    ('VALIGN', (0, 0), (0, -1), 'TOP'),
                                                    ('TEXTCOLOR', (0, 0), (0, -1), colors.black),
                                                    ('ALIGN', (0, -1), (-1, -1), 'CENTER'),
                                                    ('VALIGN', (0, -1), (-1, -1), 'MIDDLE'),
                                                    ('TEXTCOLOR', (0, -1), (-1, -1), colors.black),
                                                    ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
                                                    ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
                                                    ])
                                table = Table(compileData2)

                                table.setStyle(style)
                                table.wrapOn(c, width, height)
                                table.drawOn(c, 30 * mm, 140 * mm)

                                styles = getSampleStyleSheet()

                                ptext = (
                                            name + ' as the seller, hereby confirmed that the following Commodities (the "Product" have been sold to: ')
                                p = Paragraph(ptext, style=styles["Normal"])
                                p.wrapOn(c, 150 * mm, 100 * mm)  # size of 'textbox' for linebreaks etc.
                                p.drawOn(c, 30 * mm, 240 * mm)  # position of text / where to draw

                                buyerText = 'Buyer: '
                                pBuyer = Paragraph(buyerText, style=styles['Heading2'])
                                pBuyer.wrapOn(c, 75 * mm, 100 * mm)
                                pBuyer.drawOn(c, 81 * mm, 202 * mm)

                                valueDateText = 'Value Date: ' + valueDate
                                pValueDate = Paragraph(valueDateText, style=styles['Heading2'])
                                pValueDate.wrapOn(c, 75 * mm, 100 * mm)
                                pValueDate.drawOn(c, 70 * mm, 195 * mm)

                                ptext2 = 'The ownership of the Product shall be transferred to the Buyer on the above said Value Date.' \
                                         'The Product will be held in the location below, until otherwise advised by the Buyer'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(c, 140 * mm, 100 * mm)
                                p2.drawOn(c, 35 * mm, 160 * mm)

                                titleText = 'HOLDING CERTIFICATE'
                                pTitle = Paragraph(titleText, style=styles['Title'])
                                pTitle.wrapOn(c, 100 * mm, 20 * mm)
                                pTitle.drawOn(c, 50 * mm, 270 * mm)

                                c.setFont('Helvetica-Bold', 50)
                                styles.add(ParagraphStyle(name='certificateFont', fontName='Helvetica', fontSize=14))
                                styleCert = styles['certificateFont']
                                styleCert.alignment = TA_LEFT
                                textWidth = stringWidth(certificateNo, 'Helvetica-Bold', 50)
                                certText = ('Certificate Number: ' + certificateNo)

                                pCert = Paragraph(certText, styleCert)

                                pCert.wrapOn(c, 100 * mm, 20 * mm)
                                pCert.drawOn(c, 70 * mm, 250 * mm)
                                c.rect(20, 20, 550, 790, fill=0)
                                c.line(20, 115, 570, 115)

                                text = "For and on behalf of  "
                                textName = name
                                text2 = 'This Holding Certificate is only valid for this transaction' \
                                        ' and shall not be used for any other purposes except for the transaction' \
                                        ' as stated herein'
                                text3 = 'This document is computer generated and no signature is required.'

                                textParagraph = Paragraph(text, style=styles['Normal'])
                                textParagraph.wrapOn(c, 100 * mm, 100 * mm)
                                textParagraph.drawOn(c, 10 * mm, 35 * mm)

                                textNameParagraph = Paragraph(textName, style=styles['Normal'])
                                textNameParagraph.wrapOn(c, 100 * mm, 100 * mm)
                                textNameParagraph.drawOn(c, 10 * mm, 30 * mm)

                                text2Paragraph = Paragraph(text2, style=styles['BodyText'])
                                text2Paragraph.wrapOn(c, 180 * mm, 100 * mm)
                                text2Paragraph.drawOn(c, 10 * mm, 20 * mm)

                                date = datetime.datetime.now()
                                dateParagraph = Paragraph(date.strftime("%A %Y-%m-%d %H:%M"), style=styles['BodyText'])
                                dateParagraph.wrapOn(c, 180 * mm, 100 * mm)
                                dateParagraph.drawOn(c, 150 * mm, 35 * mm)

                                text3Paragraph = Paragraph(text3, style=styles['Code'])
                                text3Paragraph.wrapOn(c, 180 * mm, 100 * mm)
                                text3Paragraph.drawOn(c, 40 * mm, 15 * mm)

                                c.showPage()  # next page

                                titleText = 'DELIVERY ORDER'
                                pTitle = Paragraph(titleText, style=styles['Title'])
                                pTitle.wrapOn(c, 100 * mm, 20 * mm)
                                pTitle.drawOn(c, 50 * mm, 270 * mm)

                                ptext2 = 'Buyer         :'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(c, 140 * mm, 100 * mm)
                                p2.drawOn(c, 20 * mm, 240 * mm)

                                ptext2 = 'Buyer Reference :'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(c, 140 * mm, 100 * mm)
                                p2.drawOn(c, 115 * mm, 240 * mm)

                                ptext2 = 'Seller Reference :' + purchaseRef
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(c, 140 * mm, 100 * mm)
                                p2.drawOn(c, 115 * mm, 235 * mm)

                                ptext2 = 'Trade Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(c, 140 * mm, 100 * mm)
                                p2.drawOn(c, 115 * mm, 230 * mm)

                                ptext2 = 'Delivery Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(c, 140 * mm, 100 * mm)
                                p2.drawOn(c, 115 * mm, 225 * mm)

                                ptext2 = 'Payment Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(c, 140 * mm, 100 * mm)
                                p2.drawOn(c, 115 * mm, 220 * mm)

                                ptext2 = 'Certificate No \r \r  :' + certificateNo
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(c, 140 * mm, 100 * mm)
                                p2.drawOn(c, 115 * mm, 215 * mm)

                                c.line(20, 600, 570, 600)
                                c.line(20, 580, 570, 580)

                                ptext2 = 'This Delivery Order confirmed that the Commodities stated below have been sold to the Customer mentioned above'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(c, 200 * mm, 100 * mm)
                                p2.drawOn(c, 10 * mm, 205 * mm)

                                table = Table(compileData2)

                                table.setStyle(style)
                                table.wrapOn(c, width, height)
                                table.drawOn(c, 30 * mm, 170 * mm)
                                c.save()

                                #fhandle2 = open(filename + str(i) + '.pdf', 'rb')
                                print('uploading file: ',root+ filename + str(i) + '(AQAD).pdf to ' + defaultDownloadDir+'/'+filename)
                                if ftp.path.isdir(root+'/Upload')==False:
                                    ftp.mkdir(root+'/Upload',)
                                ftp.upload(defaultDownloadDir+'/'+name+'(AQAD).pdf',root+'/Upload/'+name+'(AQAD).pdf')

                                print('Upload Complete')
                                downloaded.append(name + '(AQAD).pdf')
                                downloaded.append(filename)
                                downloaded.append('/Upload/'+name+'(AQAD).pdf')

                                print('CONTENTS IN DOWNLOADED: ',downloaded)

                                d = canvas.Canvas(name + '(PURCHASE).pdf', pagesize=A4)

                                d.setTitle("PURCHASE")
                                text = ("PURCHASER'S REQUEST")
                                ptext = Paragraph(text, style=styles['Normal'])
                                ptext.wrapOn(d, 100 * mm, 20 * mm)
                                ptext.drawOn(d, 15 * mm, 250 * mm)

                                d.line(40, 700, 300, 700)

                                text = 'REPRINT ' + date.strftime("%Y-%m-%d %H:%M %p")
                                ptext = Paragraph(text, style=styles['Normal'])
                                ptext.wrapOn(d, 100 * mm, 20 * mm)
                                ptext.drawOn(d, 150 * mm, 250 * mm)

                                ptext2 = 'Date         :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 240 * mm)

                                ptext2 = 'To : Ableace Raakin Sdn Bhd'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 235 * mm)

                                ptext2 = 'Attention : (INSERT NAME HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 230 * mm)

                                ptext2 = 'Fax : (INSERT FAX NO. HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 225 * mm)

                                ptext2 = 'From : (INSERT NAME HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 220 * mm)

                                styles.add(ParagraphStyle(name='customFont1', fontName='Helvetica-Bold', fontSize=12))
                                styleCert = styles['customFont1']
                                styleCert.alignment = TA_JUSTIFY

                                ptext2 = "Commodities Purchase Agreement between (BUYERNAME) and Ableace Raakin Sdn Bhd dated " + valueDate + '.'
                                p2 = Paragraph(ptext2, style=styles['customFont1'])
                                p2.wrapOn(d, 180 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 205 * mm)

                                d.line(40, 575, 570, 575)

                                styles.add(ParagraphStyle(name='smallerFont', fontName='Helvetica', fontSize=10))
                                styleCert = styles['smallerFont']
                                styleCert.alignment = TA_JUSTIFY

                                ptext2 = "We refer to the above-referenced Agreement (the capitalized terms used in this request have the same meanings as specified in the Agreement) and" \
                                         "we hereby request that you provide to us an offer to sell to us the Commodities specified below in accordance with the terms set forth below: "

                                p2 = Paragraph(ptext2, style=styles['smallerFont'])
                                p2.wrapOn(d, 180 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 185 * mm)

                                ptext2 = "<u>Terms</u>"

                                styles.add(ParagraphStyle(name='customFont2', fontName='Helvetica-Bold', fontSize=11))
                                styleCert = styles['customFont2']
                                styleCert.alignment = TA_JUSTIFY

                                p2 = Paragraph(ptext2, style=styles['customFont2'])
                                p2.wrapOn(d, 180 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 170 * mm)

                                ptext2 = 'Seller : Ableace Raakin Sdn Bhd'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 160 * mm)

                                ptext2 = 'Purchaser : (INSERT NAME HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 155 * mm)

                                ptext2 = 'As Agent For : ' + name
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 150 * mm)

                                ptext2 = "Purchaser's Reference : " + purchaseRef
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 145 * mm)

                                ptext2 = "Certificate Number : " + certificateNo
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 140 * mm)

                                ptext2 = "Commodities : " + commodity
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 135 * mm)

                                ptext2 = "Quantity : " + str(quantity)
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 130 * mm)

                                ptext2 = "Location : " + location
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 125 * mm)

                                ptext2 = "Storage Facility : " + storage
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 120 * mm)

                                ptext2 = "Currency : " + currency
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 115 * mm)

                                ptext2 = "Price : "
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 110 * mm)

                                ptext2 = "Value Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 105 * mm)

                                ptext2 = "Delivery Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 100 * mm)

                                ptext2 = "Payment Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 95 * mm)

                                ptext2 = "For and on behalf of"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 70 * mm)

                                ptext2 = "XXXX Authorized"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 55 * mm)

                                ptext2 = "Execution Team"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 50 * mm)

                                d.showPage()  # next page

                                text = ("SELLER'S OFFER")
                                ptext = Paragraph(text, style=styles['Normal'])
                                ptext.wrapOn(d, 100 * mm, 20 * mm)
                                ptext.drawOn(d, 15 * mm, 265 * mm)

                                d.line(40, 745, 300, 745)

                                text = 'REPRINT ' + date.strftime("%Y-%m-%d %H:%M %p")
                                ptext = Paragraph(text, style=styles['Normal'])
                                ptext.wrapOn(d, 100 * mm, 20 * mm)
                                ptext.drawOn(d, 150 * mm, 265 * mm)

                                ptext2 = 'Date         :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 255 * mm)

                                ptext2 = 'To : Ableace Raakin Sdn Bhd'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 250 * mm)

                                ptext2 = 'Attention : (INSERT NAME HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 245 * mm)

                                ptext2 = 'Fax : (INSERT FAX NO. HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 240 * mm)

                                ptext2 = 'From : (INSERT NAME HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 235 * mm)

                                styleCert = styles['customFont1']
                                styleCert.alignment = TA_JUSTIFY

                                ptext2 = "Commodities Purchase Agreement between (BUYERNAME) and Ableace Raakin Sdn Bhd dated " + valueDate + '.'
                                p2 = Paragraph(ptext2, style=styles['customFont1'])
                                p2.wrapOn(d, 180 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 220 * mm)

                                d.line(40, 620, 570, 620)

                                ptext2 = "We refer to the above-referenced Agreement (the capitalized terms used in this offer have the same meanings as specified in the Agreement) and the" \
                                         " Purchaser's Request dated Tuesday, February 27, 2018 and we hereby offer to sell to you the Commodities as per your Purchaser's Request on the" \
                                         " terms specified therein which terms are as follows:"

                                p2 = Paragraph(ptext2, style=styles['smallerFont'])
                                p2.wrapOn(d, 180 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 200 * mm)

                                ptext2 = "<u>Terms</u>"
                                p2 = Paragraph(ptext2, style=styles['customFont2'])
                                p2.wrapOn(d, 180 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 190 * mm)

                                ptext2 = 'Seller : Ableace Raakin Sdn Bhd'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 180 * mm)

                                ptext2 = 'Purchaser : (INSERT NAME HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 175 * mm)

                                ptext2 = 'As Agent For : ' + name
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 170 * mm)

                                ptext2 = "Purchaser's Reference : " + purchaseRef
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 165 * mm)

                                ptext2 = "Certificate Number : " + certificateNo
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 160 * mm)

                                ptext2 = "Commodities : " + commodity
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 155 * mm)

                                ptext2 = "Quantity : " + str(quantity)
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 150 * mm)

                                ptext2 = "Location : " + location
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 145 * mm)

                                ptext2 = "Storage Facility : " + storage
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 140 * mm)

                                ptext2 = "Currency : " + currency
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 135 * mm)

                                ptext2 = "Price : "
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 130 * mm)

                                ptext2 = "Value Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 125 * mm)

                                ptext2 = "Delivery Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 120 * mm)

                                ptext2 = "Payment Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 115 * mm)

                                ptext2 = "On the Payment Date, you shall credit the Price stated above to our account as follows:"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 105 * mm)

                                ptext2 = "Bank Name : <b>Not applicable </b>"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 95 * mm)

                                ptext2 = "Account No : <b>Not applicable </b>"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 90 * mm)

                                ptext2 = "This Purchase Transactions will be concluded upon your acceptance of this offer by forwarding to us the Purchaser's Acceptance. Thereafter the" \
                                         " title, ownership and liabilities of the Commodities will be transferred to you."
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 75 * mm)

                                ptext2 = "We confirm that we are the legal and beneficial owner of the Commodities at the time of the issuance of this Seller's Offer in accordance with the" \
                                         " attached Identification Documents and that once sold to you, the Commodities will not be resold to any other party in any location."
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 55 * mm)

                                ptext2 = "Our sale of the Commodities specified above to you shall be subject to the terms of the above-referenced Agreement."
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 45 * mm)

                                ptext2 = "For and on behalf of"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 30 * mm)

                                ptext2 = "XXXX Authorized"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 15 * mm)

                                ptext2 = "Execution Team"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 10 * mm)

                                d.showPage()  # next page

                                d.line(20, 710, 570, 710)
                                d.line(20, 730, 570, 730)

                                titleText = 'SALES ATTACHMENT'
                                pTitle = Paragraph(titleText, style=styles['Title'])
                                pTitle.wrapOn(d, 100 * mm, 20 * mm)
                                pTitle.drawOn(d, 50 * mm, 250 * mm)

                                ptext2 = "To : "
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 20 * mm, 245 * mm)

                                ptext2 = "Purchaser's Reference : " + purchaseRef
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 245 * mm)

                                ptext2 = 'Seller Reference :' + purchaseRef
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 240 * mm)

                                ptext2 = 'Trade Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 235 * mm)

                                ptext2 = 'Value Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 230 * mm)

                                ptext2 = 'Delivery Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 225 * mm)

                                ptext2 = 'Attention : '
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 20 * mm, 225 * mm)

                                ptext2 = 'Payment Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 220 * mm)

                                ptext2 = 'As Agent For : ' + asagentfor
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 20 * mm, 220 * mm)

                                ptext2 = 'Certificate No \r \r  :' + certificateNo
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 215 * mm)

                                ptext2 = 'Currency : ' + currency
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 210 * mm)

                                compileData3 = []
                                compileData3 = [['Commodity', 'Location', 'Storage Facility', 'Quantity \n (MT)',
                                                 'Unit Price \n(MT)', 'Price'],
                                                [commodity, location, storage, quantity, "", price]]

                                table = Table(compileData3)

                                table.setStyle(style)
                                table.wrapOn(d, width, height)
                                table.drawOn(d, 5 * mm, 170 * mm)

                                text3 = 'This document is computer generated and no signature is required.'
                                text3Paragraph = Paragraph(text3, style=styles['Code'])
                                text3Paragraph.wrapOn(d, 180 * mm, 100 * mm)
                                text3Paragraph.drawOn(d, 40 * mm, 15 * mm)

                                date = datetime.datetime.now()
                                dateParagraph = Paragraph(date.strftime("%A %Y-%m-%d %H:%M"), style=styles['BodyText'])
                                dateParagraph.wrapOn(d, 180 * mm, 100 * mm)
                                dateParagraph.drawOn(d, 150 * mm, 35 * mm)

                                d.showPage()  # next page

                                text = ("PURCHASER'S ACCEPTANCE")
                                ptext = Paragraph(text, style=styles['Normal'])
                                ptext.wrapOn(d, 100 * mm, 20 * mm)
                                ptext.drawOn(d, 15 * mm, 250 * mm)

                                d.line(40, 700, 300, 700)

                                text = 'REPRINT ' + date.strftime("%Y-%m-%d %H:%M %p")
                                ptext = Paragraph(text, style=styles['Normal'])
                                ptext.wrapOn(d, 100 * mm, 20 * mm)
                                ptext.drawOn(d, 150 * mm, 250 * mm)

                                ptext2 = 'Date         :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 240 * mm)

                                ptext2 = 'To : Ableace Raakin Sdn Bhd'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 235 * mm)

                                ptext2 = 'Attention : (INSERT NAME HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 230 * mm)

                                ptext2 = 'Fax : (INSERT FAX NO. HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 225 * mm)

                                ptext2 = 'From : (INSERT NAME HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 220 * mm)

                                ptext2 = "Commodities Purchase Agreement between (BUYERNAME) and Ableace Raakin Sdn Bhd dated " + valueDate + '.'
                                p2 = Paragraph(ptext2, style=styles['customFont1'])
                                p2.wrapOn(d, 180 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 205 * mm)

                                d.line(40, 575, 570, 575)

                                ptext2 = "We refer to the above-referenced Agreement (the capitalized terms used in this, acceptance have the same meanings as specified in the Agreement)," \
                                         " the Purchaser's Request dated Tuesday, February 27, 2018 and the Seller's Offer dated Tuesday, February 27, 2018 and we hereby accept your offer" \
                                         " to sell to us the Commodities described in the Seller's Offer on the terms specified therein , which terms are set forth below: "

                                p2 = Paragraph(ptext2, style=styles['smallerFont'])
                                p2.wrapOn(d, 180 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 185 * mm)

                                ptext2 = "<u>Terms</u>"

                                p2 = Paragraph(ptext2, style=styles['customFont2'])
                                p2.wrapOn(d, 180 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 170 * mm)

                                ptext2 = 'Seller : Ableace Raakin Sdn Bhd'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 160 * mm)

                                ptext2 = 'Purchaser : (INSERT NAME HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 155 * mm)

                                ptext2 = 'As Agent For : ' + name
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 150 * mm)

                                ptext2 = "Seller's Reference : (INSERT SELLER'S REFERENCE HERE)"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 145 * mm)

                                ptext2 = "Purchaser's Reference : " + purchaseRef
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 140 * mm)

                                ptext2 = "Certificate Number : " + certificateNo
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 135 * mm)

                                ptext2 = "Commodities : " + commodity
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 130 * mm)

                                ptext2 = "Quantity : " + str(quantity)
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 125 * mm)

                                ptext2 = "Location : " + location
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 120 * mm)

                                ptext2 = "Storage Facility : " + storage
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 115 * mm)

                                ptext2 = "Currency : " + currency
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 110 * mm)

                                ptext2 = "Price : " + str(price)
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 105 * mm)

                                ptext2 = "Value Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 100 * mm)

                                ptext2 = "Delivery Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 95 * mm)

                                ptext2 = "Payment Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 90 * mm)

                                ptext2 = "Payment : On the Payment Date, we shall credit the Settlement Account with the Price stated above. "
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 80 * mm)

                                ptext2 = "Delivery : On the Delivery Date, you are instructed to deliver the Commodities to us in accordance with our direct instructions."
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 70 * mm)

                                ptext2 = "For and on behalf of"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 170 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 40 * mm)

                                ptext2 = "We will take reponsibility of the ownership of the Commodities and all its obligations and liabilities."
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 170 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 60 * mm)

                                ptext2 = "Our purchase of the Commodities specified above from you shall be subject to the terms of the above-referenced Agreement ."
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 15 * mm, 50 * mm)

                                text3 = 'This document is computer generated and no signature is required.'
                                text3Paragraph = Paragraph(text3, style=styles['Code'])
                                text3Paragraph.wrapOn(d, 180 * mm, 100 * mm)
                                text3Paragraph.drawOn(d, 40 * mm, 15 * mm)

                                date = datetime.datetime.now()
                                dateParagraph = Paragraph(date.strftime("%A %Y-%m-%d %H:%M"), style=styles['BodyText'])
                                dateParagraph.wrapOn(d, 180 * mm, 100 * mm)
                                dateParagraph.drawOn(d, 150 * mm, 35 * mm)

                                d.showPage()  # next page

                                table = Table(compileData2)

                                table.setStyle(style)
                                table.wrapOn(d, width, height)
                                table.drawOn(d, 30 * mm, 140 * mm)

                                styles = getSampleStyleSheet()

                                ptext = (
                                    'Ableace Raakin Sdn Bhd as the seller, hereby confirmed that the following Commodities (the "Product" have been sold to: ')
                                p = Paragraph(ptext, style=styles["Normal"])
                                p.wrapOn(d, 150 * mm, 100 * mm)  # size of 'textbox' for linebreaks etc.
                                p.drawOn(d, 30 * mm, 230 * mm)  # position of text / where to draw

                                buyerText = 'Buyer: '
                                pBuyer = Paragraph(buyerText, style=styles['Heading2'])
                                pBuyer.wrapOn(d, 75 * mm, 100 * mm)
                                pBuyer.drawOn(d, 81 * mm, 205 * mm)

                                valueDateText = 'Value Date: ' + valueDate
                                pValueDate = Paragraph(valueDateText, style=styles['Heading2'])
                                pValueDate.wrapOn(d, 75 * mm, 100 * mm)
                                pValueDate.drawOn(d, 70 * mm, 200 * mm)

                                valueDateText = 'Identification : ' + str(identification)
                                pValueDate = Paragraph(valueDateText, style=styles['Heading2'])
                                pValueDate.wrapOn(d, 100 * mm, 100 * mm)
                                pValueDate.drawOn(d, 70 * mm, 195 * mm)

                                valueDateText = 'As Agent For : ' + asagentfor
                                pValueDate = Paragraph(valueDateText, style=styles['Heading2'])
                                pValueDate.wrapOn(d, 75 * mm, 100 * mm)
                                pValueDate.drawOn(d, 70 * mm, 190 * mm)

                                ptext2 = 'The ownership of the Product shall be transferred to the Buyer on the above said Value Date.' \
                                         'The Product will be held in the location below, until otherwise advised by the Buyer'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 35 * mm, 160 * mm)

                                titleText = 'HOLDING CERTIFICATE'
                                pTitle = Paragraph(titleText, style=styles['Title'])
                                pTitle.wrapOn(d, 100 * mm, 20 * mm)
                                pTitle.drawOn(d, 50 * mm, 270 * mm)

                                d.setFont('Helvetica-Bold', 50)
                                styles.add(ParagraphStyle(name='certificateFont', fontName='Helvetica', fontSize=14))
                                styleCert = styles['certificateFont']
                                styleCert.alignment = TA_LEFT
                                textWidth = stringWidth(certificateNo, 'Helvetica-Bold', 50)
                                certText = ('Certificate Number: ' + '<b>' + certificateNo + '</b>')

                                pCert = Paragraph(certText, styleCert)

                                pCert.wrapOn(d, 100 * mm, 20 * mm)
                                pCert.drawOn(d, 70 * mm, 250 * mm)

                                d.rect(20, 20, 550, 790, fill=0)
                                d.line(20, 145, 570, 145)

                                ptext2 = 'Ableace Raakin Sdn Bhd'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 10 * mm, 56 * mm)

                                text = "For and on behalf of  "
                                textName = "<b>Ableace Raakin Sdn Bhd</b>"
                                text2 = 'This Holding Certificate is only valid for this transaction' \
                                        ' and shall not be used for any other purposes except for the transaction' \
                                        ' as stated herein'
                                text3 = 'This document is computer generated and no signature is required.'

                                textParagraph = Paragraph(text, style=styles['Normal'])
                                textParagraph.wrapOn(d, 100 * mm, 100 * mm)
                                textParagraph.drawOn(d, 10 * mm, 45 * mm)

                                textNameParagraph = Paragraph(textName, style=styles['Normal'])
                                textNameParagraph.wrapOn(d, 100 * mm, 100 * mm)
                                textNameParagraph.drawOn(d, 10 * mm, 40 * mm)

                                ptext2 = 'Azman Mohd Yussof'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 10 * mm, 35 * mm)

                                ptext2 = 'Operation Department'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 10 * mm, 30 * mm)

                                text2Paragraph = Paragraph(text2, style=styles['BodyText'])
                                text2Paragraph.wrapOn(d, 180 * mm, 100 * mm)
                                text2Paragraph.drawOn(d, 10 * mm, 20 * mm)

                                date = datetime.datetime.now()
                                dateParagraph = Paragraph(date.strftime("%A %Y-%m-%d %H:%M"), style=styles['BodyText'])
                                dateParagraph.wrapOn(d, 180 * mm, 100 * mm)
                                dateParagraph.drawOn(d, 150 * mm, 45 * mm)

                                text3Paragraph = Paragraph(text3, style=styles['Code'])
                                text3Paragraph.wrapOn(d, 180 * mm, 100 * mm)
                                text3Paragraph.drawOn(d, 40 * mm, 15 * mm)

                                d.showPage()  # next page

                                ptext2 = 'A-5-4 , Block A , Megan Avenue 1'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 55 * mm, 275 * mm)

                                ptext2 = '189 Jalan Tun Razak'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 55 * mm, 270 * mm)

                                ptext2 = '50400 Kuala Lumpur. MALAYSIA.'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 55 * mm, 265 * mm)

                                ptext2 = 'Tel: +603-21615166 / +603-21618166 / +603-21619166'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 55 * mm, 260 * mm)

                                ptext2 = 'Fax: +603-21612164 / +603-21613164'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 55 * mm, 255 * mm)

                                titleText = 'DELIVERY ORDER'
                                pTitle = Paragraph(titleText, style=styles['Title'])
                                pTitle.wrapOn(d, 100 * mm, 20 * mm)
                                pTitle.drawOn(d, 50 * mm, 240 * mm)

                                ptext2 = 'Bank Name         :'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 20 * mm, 230 * mm)

                                ptext2 = 'Your Reference :'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 230 * mm)

                                ptext2 = 'Our Reference :' + purchaseRef
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 225 * mm)

                                ptext2 = 'Trade Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 220 * mm)

                                ptext2 = 'Value Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 215 * mm)

                                ptext2 = 'Delivery Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 210 * mm)

                                ptext2 = 'Payment Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 205 * mm)

                                ptext2 = 'Attention : '
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 20 * mm, 210 * mm)

                                ptext2 = 'As Agent For : ' + asagentfor
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 20 * mm, 205 * mm)

                                ptext2 = 'Certificate No \r \r  :' + certificateNo
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 140 * mm, 100 * mm)
                                p2.drawOn(d, 115 * mm, 200 * mm)

                                d.line(20, 555, 570, 555)
                                d.line(20, 535, 570, 535)

                                ptext2 = 'This Delivery Order confirmed that the Commodities stated below have been sold to the Customer mentioned above'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(d, 200 * mm, 100 * mm)
                                p2.drawOn(d, 10 * mm, 190 * mm)

                                table = Table(compileData2)

                                table.setStyle(style)
                                table.wrapOn(d, width, height)
                                table.drawOn(d, 30 * mm, 160 * mm)

                                text3 = 'This document is computer generated and no signature is required.'
                                text3Paragraph = Paragraph(text3, style=styles['Code'])
                                text3Paragraph.wrapOn(d, 180 * mm, 100 * mm)
                                text3Paragraph.drawOn(d, 40 * mm, 15 * mm)

                                date = datetime.datetime.now()
                                dateParagraph = Paragraph(date.strftime("%A %Y-%m-%d %H:%M"), style=styles['BodyText'])
                                dateParagraph.wrapOn(d, 180 * mm, 100 * mm)
                                dateParagraph.drawOn(d, 150 * mm, 35 * mm)

                                d.save()

                                print('uploading file (PURCHASE) : ', root + filename + str(i) + '.pdf to ' + defaultDownloadDir + '/' + filename)
                                if ftp.path.isdir(root + '/Upload') == False:
                                    ftp.mkdir(root + '/Upload', )
                                ftp.upload(defaultDownloadDir + '/' + name + '(PURCHASE).pdf', root + '/Upload/' + name + '(PURCHASE).pdf')

                                print('Upload Complete')
                                downloaded.append(name + '(PURCHASE).pdf')
                                downloaded.append(filename)
                                downloaded.append('/Upload/' + name + '(PURCHASE).pdf')

                                print('CONTENTS IN DOWNLOADED: ', downloaded)


















                                e = canvas.Canvas(name + '(SALES).pdf', pagesize=A4)
                                e.setTitle("SALES")
                                text = ("SELLER'S OFFER")
                                ptext = Paragraph(text, style=styles['Normal'])
                                ptext.wrapOn(e, 100 * mm, 20 * mm)
                                ptext.drawOn(e, 15 * mm, 265 * mm)

                                e.line(40, 745, 300, 745)

                                text = 'REPRINT ' + date.strftime("%Y-%m-%d %H:%M %p")
                                ptext = Paragraph(text, style=styles['Normal'])
                                ptext.wrapOn(e, 100 * mm, 20 * mm)
                                ptext.drawOn(e, 150 * mm, 265 * mm)

                                ptext2 = 'Date         :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 255 * mm)

                                ptext2 = 'To : Antara Commodities (M) Sdn Bhd'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 250 * mm)

                                ptext2 = 'Attention : (INSERT NAME HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 245 * mm)

                                ptext2 = 'Fax : (INSERT FAX NO. HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 240 * mm)

                                ptext2 = 'From : (INSERT NAME HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 235 * mm)

                                # styleCert = styles['customFont1']
                                # styleCert.alignment = TA_JUSTIFY
                                styles.add(ParagraphStyle(name='customFont1', fontName='Helvetica-Bold', fontSize=12))

                                ptext2 = "Commodities Purchase Agreement between (BUYERNAME) and Antara Commodities (M) Sdn Bhd dated " + valueDate + '.'
                                p2 = Paragraph(ptext2, style=styles['customFont1'])
                                p2.wrapOn(e, 180 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 220 * mm)

                                e.line(40, 620, 570, 620)

                                ptext2 = "We refer to the above-referenced Agreement (the capitalized terms used" \
                                         " in this offer have the same meanings as specified in the Agreement)" \
                                         " and we are hereby offer to sell to you the Commodities on the" \
                                         " terms set forth below"

                                styles.add(ParagraphStyle(name='smallerFont', fontName='Helvetica', fontSize=10))
                                p2 = Paragraph(ptext2, style=styles['smallerFont'])
                                p2.wrapOn(e, 180 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 200 * mm)

                                styles.add(ParagraphStyle(name='customFont2', fontName='Helvetica-Bold', fontSize=11))
                                ptext2 = "<u>Terms</u>"
                                p2 = Paragraph(ptext2, style=styles['customFont2'])
                                p2.wrapOn(e, 180 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 190 * mm)

                                ptext2 = 'Purchaser : Antara Commodities (M) Sdn Bhd'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 180 * mm)

                                ptext2 = "Seller's Reference : " + purchaseRef
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 175 * mm)

                                ptext2 = 'Certificate Number : ' + certificateNo
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 170 * mm)

                                ptext2 = "Commodities : As per our Sales Attachment"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 165 * mm)

                                ptext2 = "Quantity : As per our Sales Attachment"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 160 * mm)

                                ptext2 = "Location : As per our Sales Attachment"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 155 * mm)

                                ptext2 = "Storage Facility : As per our Sales Attachment"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 150 * mm)

                                ptext2 = "Currency : " + currency
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 145 * mm)

                                ptext2 = "Price : " + str(price)
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 140 * mm)

                                ptext2 = "Value Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 135 * mm)

                                ptext2 = "Delivery Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 130 * mm)

                                ptext2 = "Payment Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 125 * mm)

                                ptext2 = "This Sale Transaction will be concluded upon your acceptance" \
                                         " of this offer by forwarding to us the Purchaser's Acceptance. " \
                                         "Thereafter the title and ownership of the Commodities will be" \
                                         " transferred to you."
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 170 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 105 * mm)

                                ptext2 = "Our sale of the Commodities specified above to you shall be" \
                                         " subject to the terms of the above-referenced Agreement"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 170 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 90 * mm)

                                ptext2 = "For and on behalf of"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 50 * mm)

                                ptext2 = "XXXX Authorized"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 35 * mm)

                                ptext2 = "Execution Team"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 30 * mm)

                                text3 = 'This document is computer generated and no signature is required.'
                                text3Paragraph = Paragraph(text3, style=styles['Code'])
                                text3Paragraph.wrapOn(e, 180 * mm, 100 * mm)
                                text3Paragraph.drawOn(e, 40 * mm, 15 * mm)

                                date = datetime.datetime.now()
                                dateParagraph = Paragraph(date.strftime("%A %Y-%m-%d %H:%M"), style=styles['BodyText'])
                                dateParagraph.wrapOn(e, 180 * mm, 100 * mm)
                                dateParagraph.drawOn(e, 150 * mm, 35 * mm)

                                e.showPage()  # next page

                                e.line(20, 710, 570, 710)
                                e.line(20, 730, 570, 730)

                                titleText = 'SALES ATTACHMENT'
                                pTitle = Paragraph(titleText, style=styles['Title'])
                                pTitle.wrapOn(e, 100 * mm, 20 * mm)
                                pTitle.drawOn(e, 50 * mm, 250 * mm)

                                ptext2 = "To : "
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 20 * mm, 245 * mm)

                                ptext2 = "Our Reference : " + purchaseRef
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 245 * mm)

                                ptext2 = 'Trade Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 240 * mm)

                                ptext2 = 'Value Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 235 * mm)

                                ptext2 = 'Delivery Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 230 * mm)

                                ptext2 = 'Payment Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 225 * mm)

                                ptext2 = 'Certificate Number :' + certificateNo
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 215 * mm)

                                ptext2 = 'Currency :' + currency
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 210 * mm)

                                ptext2 = 'Attention : '
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 20 * mm, 225 * mm)

                                ptext2 = 'Payment Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 220 * mm)

                                compileData3 = []
                                compileData3 = [['Commodity', 'Location', 'Storage Facility', 'Quantity \n (MT)',
                                                 'Unit Price \n(MT)', 'Price'],
                                                [commodity, location, storage, quantity, "", price]]

                                table = Table(compileData3)

                                table.setStyle(style)
                                table.wrapOn(e, width, height)
                                table.drawOn(e, 5 * mm, 170 * mm)

                                text3 = 'This document is computer generated and no signature is required.'
                                text3Paragraph = Paragraph(text3, style=styles['Code'])
                                text3Paragraph.wrapOn(e, 180 * mm, 100 * mm)
                                text3Paragraph.drawOn(e, 40 * mm, 15 * mm)

                                date = datetime.datetime.now()
                                dateParagraph = Paragraph(date.strftime("%A %Y-%m-%d %H:%M"), style=styles['BodyText'])
                                dateParagraph.wrapOn(e, 180 * mm, 100 * mm)
                                dateParagraph.drawOn(e, 150 * mm, 35 * mm)

                                e.showPage()

                                text = ("PURCHASER'S ACCEPTANCE")
                                ptext = Paragraph(text, style=styles['Normal'])
                                ptext.wrapOn(e, 100 * mm, 20 * mm)
                                ptext.drawOn(e, 15 * mm, 250 * mm)

                                e.line(40, 700, 300, 700)

                                text = 'REPRINT ' + date.strftime("%Y-%m-%d %H:%M %p")
                                ptext = Paragraph(text, style=styles['Normal'])
                                ptext.wrapOn(e, 100 * mm, 20 * mm)
                                ptext.drawOn(e, 150 * mm, 250 * mm)

                                ptext2 = 'Date         :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 240 * mm)

                                ptext2 = 'To : '
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 235 * mm)

                                ptext2 = 'Fax No: (INSERT FAX NO. HERE)'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 230 * mm)

                                ptext2 = 'Attention : XXXX Authorized'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 225 * mm)

                                ptext2 = 'From : Antara Commodities (M)Sdn Bhd'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 220 * mm)

                                ptext2 = "Commodities Purchase Agreement between (BUYERNAME) and Ableace Raakin Sdn Bhd dated " + valueDate + '.'
                                p2 = Paragraph(ptext2, style=styles['customFont1'])
                                p2.wrapOn(e, 180 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 205 * mm)

                                e.line(40, 575, 570, 575)

                                ptext2 = "We refer to the above-referenced Agreement (the capitalized terms used in this, acceptance have the same meanings as specified in the Agreement)," \
                                         " the Purchaser's Request dated Tuesday, February 27, 2018 and the Seller's Offer dated Tuesday, February 27, 2018 and we hereby accept your offer" \
                                         " to sell to us the Commodities described in the Seller's Offer on the terms specified therein , which terms are set forth below: "

                                p2 = Paragraph(ptext2, style=styles['smallerFont'])
                                p2.wrapOn(e, 180 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 185 * mm)

                                ptext2 = "<u>Terms</u>"

                                p2 = Paragraph(ptext2, style=styles['customFont2'])
                                p2.wrapOn(e, 180 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 170 * mm)

                                ptext2 = 'Seller : '
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 160 * mm)

                                ptext2 = 'Purchaser : Antara Commodities (M) Sdn Bhd'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 155 * mm)

                                ptext2 = "Seller's Reference : "
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 150 * mm)

                                ptext2 = "Purchaser's Reference : " + purchaseRef
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 145 * mm)

                                ptext2 = "Certificate Number : " + certificateNo
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 140 * mm)

                                ptext2 = "Commodities : As per your Sales Attachment"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 135 * mm)

                                ptext2 = "Quantity : As per your Sales Attachment"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 130 * mm)

                                ptext2 = "Location : As per your Sales Attachment"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 125 * mm)

                                ptext2 = "Storage Facility : As per your Sales Attachment"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 120 * mm)

                                ptext2 = "Currency : " + currency
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 115 * mm)

                                ptext2 = "Price : " + str(price)
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 110 * mm)

                                ptext2 = "Value Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 105 * mm)

                                ptext2 = "Delivery Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 100 * mm)

                                ptext2 = "Payment Date : " + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 95 * mm)

                                ptext2 = "Payment : On the Payment date, we shall credit your account with the Price stated above "
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 85 * mm)

                                ptext2 = "Delivery : On the Delivery date, you are instructed to deliver the Commodities to us in accordance with our direct" \
                                         " instructions."
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 75 * mm)

                                ptext2 = "For and on behalf of"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 170 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 40 * mm)

                                ptext2 = "<b>Antara Commodities (M) Sdn Bhd</b>"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 170 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 35 * mm)

                                ptext2 = "Shannon Fernandex"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 170 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 30 * mm)

                                ptext2 = "Operation"
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 170 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 25 * mm)

                                ptext2 = "Upon the execution of this acceptance, we will take responsibility of the ownership" \
                                         " of the Commodities and all its obligations and liabilities thereof."
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 170 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 60 * mm)

                                ptext2 = "Our purchase of the Commodities specified above from you shall be subject to the terms of the above-referenced Agreement ."
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 15 * mm, 50 * mm)

                                text3 = 'This document is computer generated and no signature is required.'
                                text3Paragraph = Paragraph(text3, style=styles['Code'])
                                text3Paragraph.wrapOn(e, 180 * mm, 100 * mm)
                                text3Paragraph.drawOn(e, 40 * mm, 15 * mm)

                                date = datetime.datetime.now()
                                dateParagraph = Paragraph(date.strftime("%A %Y-%m-%d %H:%M"), style=styles['BodyText'])
                                dateParagraph.wrapOn(e, 180 * mm, 100 * mm)
                                dateParagraph.drawOn(e, 150 * mm, 35 * mm)

                                #############
                                e.showPage()  # next page

                                #############

                                table = Table(compileData2)

                                table.setStyle(style)
                                table.wrapOn(e, width, height)
                                table.drawOn(e, 30 * mm, 140 * mm)

                                styles = getSampleStyleSheet()

                                ptext = (
                                    'Malaysia Building Sdn Bhd as the seller, hereby confirmed that the following Commodities (the "Product") have been sold to: ')
                                p = Paragraph(ptext, style=styles["Normal"])
                                p.wrapOn(e, 150 * mm, 100 * mm)  # size of 'textbox' for linebreaks etc.
                                p.drawOn(e, 30 * mm, 240 * mm)  # position of text / where to draw

                                buyerText = 'Buyer: Antara Commodities (M) Sdn Bhd'
                                pBuyer = Paragraph(buyerText, style=styles['Normal'])
                                pBuyer.wrapOn(e, 125 * mm, 100 * mm)
                                pBuyer.drawOn(e, 65 * mm, 202 * mm)

                                valueDateText = 'Value Date: ' + valueDate
                                pValueDate = Paragraph(valueDateText, style=styles['Normal'])
                                pValueDate.wrapOn(e, 75 * mm, 100 * mm)
                                pValueDate.drawOn(e, 65 * mm, 195 * mm)

                                ptext2 = 'The ownership of the Product shall be transferred to the Buyer on the above said Value Date.' \
                                         'The Product will be held in the location below, until otherwise advised by the Buyer'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 35 * mm, 160 * mm)

                                titleText = 'HOLDING CERTIFICATE'
                                pTitle = Paragraph(titleText, style=styles['Title'])
                                pTitle.wrapOn(e, 100 * mm, 20 * mm)
                                pTitle.drawOn(e, 50 * mm, 270 * mm)

                                c.setFont('Helvetica-Bold', 50)

                                styles.add(ParagraphStyle(name='certificateFont', fontName='Helvetica', fontSize=14))
                                styleCert = styles['certificateFont']
                                styleCert.alignment = TA_LEFT

                                textWidth = stringWidth(certificateNo, 'Helvetica-Bold', 50)
                                certText = ('Certificate Number: ' + '<b>' + certificateNo + '</b>')
                                pCert = Paragraph(certText, styleCert)
                                pCert.wrapOn(e, 100 * mm, 20 * mm)
                                pCert.drawOn(e, 60 * mm, 250 * mm)

                                e.rect(20, 20, 550, 790, fill=0)
                                e.line(20, 145, 570, 145)

                                text = "For and on behalf of  "
                                textName = "XXXX Authorized"
                                text2 = 'This Holding Certificate is only valid for this transaction' \
                                        ' and shall not be used for any other purposes except for the transaction' \
                                        ' as stated herein'
                                text3 = 'This document is computer generated and no signature is required.'
                                text4 = 'Execution Team'

                                textParagraph = Paragraph(text, style=styles['Normal'])
                                textParagraph.wrapOn(e, 100 * mm, 100 * mm)
                                textParagraph.drawOn(e, 10 * mm, 45 * mm)

                                textNameParagraph = Paragraph(textName, style=styles['Normal'])
                                textNameParagraph.wrapOn(e, 100 * mm, 100 * mm)
                                textNameParagraph.drawOn(e, 10 * mm, 40 * mm)

                                textNameParagraph = Paragraph(text4, style=styles['Normal'])
                                textNameParagraph.wrapOn(e, 100 * mm, 100 * mm)
                                textNameParagraph.drawOn(e, 10 * mm, 35 * mm)

                                text2Paragraph = Paragraph(text2, style=styles['BodyText'])
                                text2Paragraph.wrapOn(e, 180 * mm, 100 * mm)
                                text2Paragraph.drawOn(e, 10 * mm, 20 * mm)

                                date = datetime.datetime.now()
                                dateParagraph = Paragraph(date.strftime("%A %Y-%m-%d %H:%M"), style=styles['BodyText'])
                                dateParagraph.wrapOn(e, 180 * mm, 100 * mm)
                                dateParagraph.drawOn(e, 150 * mm, 45 * mm)

                                text3Paragraph = Paragraph(text3, style=styles['Code'])
                                text3Paragraph.wrapOn(e, 180 * mm, 100 * mm)
                                text3Paragraph.drawOn(e, 40 * mm, 15 * mm)

                                #####
                                e.showPage()  # next page
                                ########
                                text = 'REPRINT ' + date.strftime("%Y-%m-%d %H:%M %p")
                                ptext = Paragraph(text, style=styles['Normal'])
                                ptext.wrapOn(e, 100 * mm, 20 * mm)
                                ptext.drawOn(e, 150 * mm, 250 * mm)

                                titleText = 'DELIVERY ORDER'
                                pTitle = Paragraph(titleText, style=styles['Title'])
                                pTitle.wrapOn(e, 100 * mm, 20 * mm)
                                pTitle.drawOn(e, 50 * mm, 240 * mm)

                                ptext2 = 'Buyer         : Antara Commodities (M) Sdn Bhd'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 20 * mm, 230 * mm)

                                ptext2 = 'Buyer Reference :'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 230 * mm)

                                ptext2 = 'Seller Reference :' + purchaseRef
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 225 * mm)

                                ptext2 = 'Trade Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 220 * mm)

                                ptext2 = 'Value Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 215 * mm)

                                ptext2 = 'Delivery Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 210 * mm)

                                ptext2 = 'Payment Date :' + valueDate
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 205 * mm)

                                ptext2 = 'Attention : Mr Martin Charles Fernandez'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 20 * mm, 210 * mm)

                                ptext2 = 'Certificate No \r \r  :' + certificateNo
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 140 * mm, 100 * mm)
                                p2.drawOn(e, 115 * mm, 200 * mm)

                                e.line(20, 555, 570, 555)
                                e.line(20, 535, 570, 535)

                                ptext2 = 'This Delivery Order confirmed that the Commodities stated below have been sold to the Customer mentioned above'
                                p2 = Paragraph(ptext2, style=styles['Normal'])
                                p2.wrapOn(e, 200 * mm, 100 * mm)
                                p2.drawOn(e, 10 * mm, 190 * mm)

                                table = Table(compileData2)

                                table.setStyle(style)
                                table.wrapOn(e, width, height)
                                table.drawOn(e, 30 * mm, 160 * mm)

                                text3 = 'This document is computer generated and no signature is required.'
                                text3Paragraph = Paragraph(text3, style=styles['Code'])
                                text3Paragraph.wrapOn(e, 180 * mm, 100 * mm)
                                text3Paragraph.drawOn(e, 40 * mm, 15 * mm)

                                date = datetime.datetime.now()
                                dateParagraph = Paragraph(date.strftime("%A %Y-%m-%d %H:%M"), style=styles['BodyText'])
                                dateParagraph.wrapOn(e, 180 * mm, 100 * mm)
                                dateParagraph.drawOn(e, 150 * mm, 35 * mm)

                                e.save()

                                print('uploading file (SALES) : ',
                                      root + filename + str(i) + '.pdf to ' + defaultDownloadDir + '/' + filename)
                                if ftp.path.isdir(root + '/Upload') == False:
                                    ftp.mkdir(root + '/Upload', )
                                ftp.upload(defaultDownloadDir + '/' + name + '(SALES).pdf', root + '/Upload/' + name + '(SALES).pdf')

                                print('Upload Complete')
                                downloaded.append(name + '(SALES).pdf')
                                downloaded.append(filename)
                                downloaded.append('/Upload/' + name + '(SALES).pdf')

                                print('CONTENTS IN DOWNLOADED: ', downloaded)


        else:
            ftp.close()
            skipped+=1


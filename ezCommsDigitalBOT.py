import time
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler
import shutil
import fitz
import natsort
import calendar
from pdf2image import convert_from_path
import os
import moviepy.video.io.ImageSequenceClip
import moviepy.editor as mp
import pypandoc
import argparse
import docx2txt
from mhmovie.code import *
import requests
# create parser
from gtts import gTTS
from mailmerge import MailMerge
from datetime import date
from docx2pdf import convert
#from docx2html import convert as htmlcovert

from datetime import datetime
import os
import lxml.etree
from moviepy.editor import *
# xmlstr is your xml in a string
import time
# 1593586355.9822
from cloudinary.api import delete_resources_by_tag, resources_by_tag
from cloudinary.uploader import upload
from cloudinary.utils import cloudinary_url
from twilio.rest import Client

import xmltodict


parser = argparse.ArgumentParser()

#parser.add_argument("triggerFile")
parser.add_argument("configFile")
args = parser.parse_args()

#triggerFile = args.triggerFile

configFile = args.configFile
# parse the arguments
args = parser.parse_args()
print(configFile)
#exit(1)

#configdir = "D:/e_Docs_Runtime/config/"
#configFile="D:/e_Docs_Runtime/config/properties.xml"
triggerFile=''
xmldir = ''
triggerdir=''
pdfoutdir = ''
htmloutdir = ""
digitaltemplatedir = ""
audiotemplatedir=""
videotemplatedir=""
htmloutdir = ""
audiooutdir = ""
videooutdir = ""

docoutdir = ""
docxoutdir = ""
templatedir = ""
resourcedir = ""
pdfoutputflag = ""
wordoutputflag = ""
htmloutputflag = ""
videooutputflag = ""
dotDoc = ""
dotPdf = ""
dotDocx = ""
dotHtml = ""
dotxml = ""
templateExtn = ""
imageFileExtn = ""
errorPDFPath = ""
templateName=""
xmlFileNameTag = ""
audioExtn =""
videoExtn =""


def append_timestamp(filename):
    tm = time.strftime('%a%d%b%Y%H%M%S')
    filename_with_timestamp = filename + "_" + tm
    return (filename_with_timestamp)

print("Start Time ......" + append_timestamp("start"))
# xmldir pdfoutdir htmloutdir docdir docxdir templatedir configdir resourcedir pdfoutputflag wordoutputflag htmloutputflag videooutputflag
# dotDoc dotPdf dotDocx dotHtml templateExtn imageFileExtn errorPDFPath
def loadconfig():
    with open(configFile) as fdxml:
        properites  = xmltodict.parse(fdxml.read())
        xmldir  = properites['properties']['xmldir']
        triggerdir = properites['properties']['triggerdir']
        pdfoutdir = properites['properties']['pdfoutdir']
        htmloutdir = properites['properties']['htmloutdir']
        audiooutdir = properites['properties']['audiooutdir']
        videooutdir = properites['properties']['videooutdir']
        docoutdir = properites['properties']['docoutdir']
        docxoutdirdigitalchannel = properites['properties']['docxoutdirdigitalchannel']
        docxoutdirAudio = properites['properties']['docxoutdirAudio']
        docxoutdirVideo = properites['properties']['docxoutdirVideo']
        digitaltemplatedir = properites['properties']['digitaltemplatedir']
        audiotemplatedir = properites['properties']['audiotemplatedir']
        videotemplatedir = properites['properties']['videotemplatedir']
        configdir = properites['properties']['configdir']
        resourcedir = properites['properties']['resourcedir']
        pdfoutputflag = properites['properties']['pdfoutputflag']
        wordoutputflag = properites['properties']['wordoutputflag']
        htmloutputflag = properites['properties']['htmloutputflag']
        videooutputflag = properites['properties']['videooutputflag']
        dotDoc = properites['properties']['dotDoc']
        dotPdf = properites['properties']['dotPdf']
        dotDocx = properites['properties']['dotDocx']
        dotHtml = properites['properties']['dotHtml']
        dotxml= properites['properties']['dotxml']
        templateExtn = properites['properties']['templateExtn']
        processCompleteExtn = properites['properties']['processCompleteExtn']
        processErrorExtn = properites['properties']['processErrorExtn']
        imageFileExtn = properites['properties']['imageFileExtn']
        xmlFileNameTag = properites['properties']['xmlFileNameTag']
        audioExtn = properites['properties']['audioExtn']
        videoExtn = properites['properties']['videoExtn']
        errorPDFPath = properites['properties']['errorPDFPath']
        triggerFile=properites['properties']['triggerFile']
        xmlFileNameTag= properites['properties']['xmlFileNameTag']

    fdxml.close()
    #print("inside - congig" , xmlFileNameTag + docxoutdirdigitalchannel)
    return properites

def dataPopulationintoTemplate(document,doc,templateName,docxoutFilename):
    document = executeMappingRule(document, doc, templateName)
    print("execute mapping after", docxoutFilename)
    document.write(docxoutFilename)
    print("write document",docxoutFilename)
    return (document)

#exit(1)
def executeMappingRule(document,doc,templateID):
  if templateID == "RB_CLM_LetterTemplate_6798":
      document.merge(addressLine3=doc['correspondence']['addressLine3'],fullName=doc['correspondence']['fullName'],expiryDate=doc['correspondence']['expiryDate'],addressLine1=doc['correspondence']['addressLine1'],addressLine4=doc['correspondence']['addressLine4'],balanceAmount=doc['correspondence']['balanceAmount'],accountNumber=doc['correspondence']['accountNumber'],balanceDate=doc['correspondence']['balanceDate'],shortName=doc['correspondence']['shortName'],currentDate=doc['correspondence']['currentDate'],addressLine5=doc['correspondence']['addressLine5'],operatorName=doc['correspondence']['operatorName'],addressLine2=doc['correspondence']['addressLine2'])
  elif templateID == "tcs_retirement_letter_template_9867":
      document.merge(PSamount=doc['correspondence']['PSamount'],openBalanceAmount=doc['correspondence']['openBalanceAmount'],netMoneyOut=doc['correspondence']['netMoneyOut'],
                     fullName=doc['correspondence']['fullName'],
                     closingDate=doc['correspondence']['closingDate'],
                     catchUpLimit=doc['correspondence']['catchUpLimit'],
                     netMoneyIn=doc['correspondence']['netMoneyIn'],
                     yearlyRetrun=doc['correspondence']['yearlyRetrun'],
                     quaterlyRetrun=doc['correspondence']['quaterlyRetrun'],
                     InversmentGainorLoss=doc['correspondence']['InversmentGainorLoss'],
                     openBalanceDate=doc['correspondence']['openBalanceDate'],
                     increaseFAmount=doc['correspondence']['increaseFAmount'],
                     closingBalance=doc['correspondence']['closingBalance'])
  elif templateID == "acme_statement_6789":
      document.merge(productName=doc['correspondence']['productName'],
                     CLB=doc['correspondence']['CLB'],
                     IGA=doc['correspondence']['IGA'],
                     accountName=doc['correspondence']['accountName'],
                     lastName=doc['correspondence']['lastName'],
                     OPB=doc['correspondence']['OPB'],
                     statementPeriod=doc['correspondence']['statementPeriod'],
                     QPR=doc['correspondence']['QPR'],
                     agentName=doc['correspondence']['agentName'],
                     APR=doc['correspondence']['APR'],
                     policyPeriod=doc['correspondence']['policyPeriod'],
                     firstName=doc['correspondence']['firstName'])
  elif templateID == "NATW_statement_6789":
      document.merge(balanceDate=doc['correspondence']['balanceDate'], state=doc['correspondence']['state'],
                     addressLine1=doc['correspondence']['addressLine1'], fullName=doc['correspondence']['fullName'],
                     balanceAmount=doc['correspondence']['balanceAmount'],
                     expiryDate=doc['correspondence']['expiryDate'],
                     accountNumber=doc['correspondence']['accountNumber'],
                     operatorName=doc['correspondence']['operatorName'],
                     addressLine4=doc['correspondence']['addressLine4'], shortName=doc['correspondence']['shortName'],
                     addressLine3=doc['correspondence']['addressLine3'],
                     currentDate=doc['correspondence']['currentDate'],
                     addressLine2=doc['correspondence']['addressLine2'],
                     addressLine5=doc['correspondence']['addressLine5'])
  else:
      print("error in template mapping. Please complete the mapping ")
  return document

customerID =""



def generateDigitalPDF(inputXML,properites):
    try:
        with open(inputXML) as fd:
            doc = xmltodict.parse(fd.read())
            print("Rename Complete")
            print("xmlFileNameTag", xmlFileNameTag)
            templateName = (doc['correspondence'][properites['properties']['xmlFileNameTag']])
            #customerName = (doc['correspondence']["firstName"])
            outFilename = templateName
            outFilename = append_timestamp(outFilename)
            print("Template Path:",
                  properites['properties']['digitaltemplatedir'] + templateName + properites['properties']['templateExtn'])
            digitaltemplateFilePath = properites['properties']['digitaltemplatedir'] + templateName + \
                                      properites['properties']['templateExtn']
            document = MailMerge(digitaltemplateFilePath)
            outFilename = append_timestamp(outFilename)
            print(outFilename + "......")
            docxoutFilename = properites['properties']['docxoutdirdigitalchannel'] + outFilename + properites['properties'][
                'dotDocx']
            document = dataPopulationintoTemplate(document, doc, templateName, docxoutFilename)
            document.close()
            convert(docxoutFilename,
                    properites['properties']['pdfoutdir'] + outFilename + properites['properties']['dotPdf'])
            #print("PDF Generation  Done ......" + append_timestamp("PDF"))
            HTML_G = append_timestamp("HTML")
            pdfFilePath = properites['properties']['pdfoutdir'] + outFilename + properites['properties']['dotPdf']
            htmlFilepath=properites['properties']['htmloutdir'] + outFilename + properites['properties']['dotHtml']
            output = pypandoc.convert_file(docxoutFilename, 'html5', outputfile=htmlFilepath)
            print("HTMLG", HTML_G, "HTML Output Location ",htmlFilepath)
        fd.close()
        print("process complete")
        print("File Rename")
        head, tail = os.path.split(inputXML)
        print(head,tail,properites['properties']['processCompleteExtn'] )
        destination = "D:/ezComms_Runtime/processedxml/" + tail + properites['properties']['processCompleteExtn']
        os.rename(inputXML, destination)
        source = inputXML
        head, tail = os.path.split(source)
        print(destination,source)
        print("File Rename Done")
    except:
        #document.close()
        print("The Trigger  Failed:- ")
        #pass

    return (pdfFilePath, outFilename + properites['properties']['dotPdf'])

patterns = "*"
ignore_patterns = ""
ignore_directories = False
case_sensitive = True
my_event_handler = PatternMatchingEventHandler(patterns, ignore_patterns, ignore_directories, case_sensitive)


def on_created(event):
    try:
        print(f" eComms BOT recevied  {event.src_path}!")
        properites = loadconfig()
        inputXML = event.src_path
        filepath , outFileName =  generateDigitalPDF(inputXML,properites)
        print(filepath , outFileName ,  "has been created by the BOT!")
    except :
        pass
def on_deleted(event):
    print(f"eComms BOT Trigger deleted {event.src_path}!")

def on_modified(event):
    # properites = loadconfig()
    # filepath , outFileName =  generateDigitalPDF(inputXML,properites):
    # print(filepath , outFileName ,  "has been Modified!")
    print("Trigger Modified")
def on_moved(event):
    print(f"BOT Moved {event.src_path} to {event.dest_path}")

my_event_handler.on_created = on_created
my_event_handler.on_deleted = on_deleted
my_event_handler.on_modified = on_modified
my_event_handler.on_moved = on_moved
properites = loadconfig()
xmldir  = properites['properties']['xmldir']
go_recursively = True
my_observer = Observer()
my_observer.schedule(my_event_handler, xmldir, recursive=go_recursively)
my_observer.start()

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    my_observer.stop()
    my_observer.join()
except :
    print("The Trigger  Failed:- for the file" , event.src_path)
    pass

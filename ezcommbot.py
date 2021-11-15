import moviepy.video.io.ImageSequenceClip
import moviepy.editor as mp
import natsort
from gtts import gTTS
from moviepy.editor import *
import time
import xmltodict
import sys
import os
import json

#parser = argparse.ArgumentParser()
# #parser.add_argument("triggerFile")
#parser.add_argument("inputXMLFile")
#args = parser.parse_args()
#inputXMLFile = args.inputXMLFile

#configdir = "D:/e_Docs_Runtime/config/"
#configFile="D:/ezComms_Runtime/config/properties.xml"
#LIBRE_OFFICE = r"C:/Program Files/LibreOffice/program/soffice.exe"

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
docxoutdirdigitalchannel = ""
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

def dump_response(response):
    print("Upload response:")
    for key in sorted(response.keys()):
        print("  %s: %s" % (key, response[key]))


def sendMessage(urlinput, to_whatsapp_number, message):
    # client credentials are read from TWILIO_ACCOUNT_SID and AUTH_TOKEN
    client = Client()

    # this is the Twilio sandbox testing number
    from_whatsapp_number = 'whatsapp:+14155238886'
    # replace this number with your own WhatsApp Messaging number
    # to_whatsapp_number = 'whatsapp:+917708605798'

    client.messages.create(body=message,
                           from_=from_whatsapp_number,
                           to=to_whatsapp_number)

    message = client.messages \
        .create(
        media_url=urlinput,
        from_=from_whatsapp_number,
        to=to_whatsapp_number
    )


# xmldir pdfoutdir htmloutdir docdir docxdir templatedir configdir resourcedir pdfoutputflag wordoutputflag htmloutputflag videooutputflag
# dotDoc dotPdf dotDocx dotHtml templateExtn imageFileExtn errorPDFPath
def loadconfig(configFile):
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
    print("inside - congig" , xmlFileNameTag + docxoutdirdigitalchannel)
    return properites

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
    elif templateID == "CITI_POC_template_1001":
        document.merge(lastname=doc['correspondence']['lastname'],
                       surplus=doc['correspondence']['surplus'],trend=doc['correspondence']['trend'],
                       accountType=doc['correspondence']['accountType'],asofdate=doc['correspondence']['asofdate'],
                       consultName=doc['correspondence']['consultName'],firstname=doc['correspondence']['firstname'],
                       cashout=doc['correspondence']['cashout'],cashin=doc['correspondence']['cashin'],membersince=doc['correspondence']['membersince'],
                       contactperson=doc['correspondence']['contactperson'])
    elif templateID == "citi_statement_6789":
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
    elif templateID == "LetterTemplate_6798":
        document.merge(balanceDate=doc['correspondence']['balanceDate'],state=doc['correspondence']['state'],addressLine1=doc['correspondence']['addressLine1'],fullName=doc['correspondence']['fullName'],balanceAmount=doc['correspondence']['balanceAmount'],expiryDate=doc['correspondence']['expiryDate'],accountNumber=doc['correspondence']['accountNumber'],operatorName=doc['correspondence']['operatorName'],addressLine4=doc['correspondence']['addressLine4'],shortName=doc['correspondence']['shortName'],addressLine3=doc['correspondence']['addressLine3'],currentDate=doc['correspondence']['currentDate'],addressLine2=doc['correspondence']['addressLine2'],addressLine5=doc['correspondence']['addressLine5'])
    else:
        print("error in template mapping. Please complete the mapping ")
    return document

def dataPopulationintoTemplate(document,doc,templateName,docxoutFilename):

    document = executeMappingRule(document, doc, templateName)
    print("execute mapping after", docxoutFilename)
    document.write(docxoutFilename)
    print("write document",docxoutFilename)
    return (document)

def dump_response(response):
    print("Upload response:")
    for key in sorted(response.keys()):
        print("  %s: %s" % (key, response[key]))


def word2PDF(input_docx,out_folder,LIBRE_OFFICE):
    p = Popen([LIBRE_OFFICE, '--headless', '--convert-to', 'pdf', '--outdir',
               out_folder, input_docx])
    print([LIBRE_OFFICE, '--convert-to', 'pdf', input_docx])
    p.communicate()

def generateVideoCommunication(inputXML,properites):
    #with open(triggerFile) as input_file:
    startTime = append_timestamp("Start")
    print("Video Generation Start ......" + startTime)
    #print(inputXML)
    status = "failed"
    reason = "video generation failed"
    finlavideocloudpath = "Error"
    try:
        doc = xmltodict.parse(inputXML)
        audiodata =""
        with open (doc['correspondence']['audioText'], "r") as myfile:
            audiodata=myfile.readlines()
        myfile.close()
        _TEXT=audiodata
        language = 'en'
        myobj = gTTS(text=_TEXT, lang=language, slow=False)
        myobj.save(doc['correspondence']['audiooutFileName'])
        print("Audio Generation Done  ......" + append_timestamp("Audio"))
        AUDIO_G = append_timestamp("AUDIO")

        fps = 0.5
        image_folder = doc['correspondence']['imagePath'] 
        image_files = [image_folder + '/' + img for img in sorted(os.listdir(image_folder)) if img.endswith(".jpg")]
        image_files = natsort.natsorted(image_files, reverse=False)
        #print(image_files)
        clip = moviepy.video.io.ImageSequenceClip.ImageSequenceClip(image_files, fps=fps)
        clip.write_videofile(doc['correspondence']['videopocessing'])
        videoclip = VideoFileClip(doc['correspondence']['videopocessing'])
        audioclip = AudioFileClip(doc['correspondence']['audiooutFileName'])
        new_audioclip = CompositeAudioClip([audioclip])
        videoclip.audio = new_audioclip
        videoclip.write_videofile(doc['correspondence']['videooutfile'], audio_codec="aac")
        print("Start Time.........."  + startTime)
        print("Video Generation Done ......" + append_timestamp("Video"))
        status = "passed"
        reason = "video generation done"
        documentA.close()
        documentV.close()
        DEFAULT_TAG = "citiusa"
        print("--- video file ID" + doc['correspondence']['videooutfile'])
        finlavideocloudpath=outFilename
    except OSError:
        documentA.close()
        documentV.close()
        status = "failed"
        reason = "video generation The Trigger  Failed"
        print("The Trigger  Failed:- " )

    return(status,reason,finlavideocloudpath,doc['correspondence']['jsonout'])


def sendsVideoCommunication(inputXML,properites):
    doc = xmltodict.parse(inputXML)
    clip = mp.VideoFileClip(properites['properties']['videooutdir']+doc['correspondence']['outVideoFilename']+properites['properties']['videoExtn'])
    clip_resized = clip.resize(height=200)  # make the height 360px ( According to moviePy documenation The width is then computed so that the width/height ratio is conserved.)
    clip_resized.write_videofile(properites['properties']['videooutdir'] + "whatsup/" + doc['correspondence']['outVideoFilename']+properites['properties']['videoExtn'],audio_codec="aac")
    VA_G = append_timestamp("VideoAudio")
    filePath = properites['properties']['videooutdir'] + "whatsup/" + doc['correspondence']['outVideoFilename']+properites['properties']['videoExtn']
    DEFAULT_TAG = "Try Kannan"
    print("--- Upload a local file")
    response = upload(filePath, resource_type= "video")
    dump_response(response)
    path=response["url"]
    print("",)
    message = "Welcome " + doc['correspondence']['firstName'] + " to Acke Retirement, This is your Video Statement for your retirement benefit::"
    to_whatsapp_number= doc['correspondence']['to_whatsapp_number']
    sendMessage(path,to_whatsapp_number,message)
    return( "Video Correspondance Sent to:"+ to_whatsapp_number )


# inputXML = "<correspondence><companyName>TCS</companyName>" \
# "<templateID>citi_statement_6789</templateID>" \
# "<customerID>ID123</customerID>" \
# "<to_whatsapp_number>whatsapp:+919789260241</to_whatsapp_number>" \
# "<accountName>D5467</accountName>" \
# "<productName>AKME POWER</productName>" \
# "<firstName>Faraday</firstName>" \
# "<lastName>Madasamy</lastName>" \
# "<agentName>Tom Smith</agentName>" \
# "<policyPeriod>2005-2035</policyPeriod>" \
# "<statementPeriod>June 2020</statementPeriod>" \
# "<statementDate>June 2020</statementDate>" \
# "<APR>21.3</APR>" \
# "<QPR>12.2</QPR>" \
# "<OPB>37.4 M</OPB>" \
# "<CLB>44.9 M</CLB>" \
# "<IGA>$900k in 30 days</IGA>" \
# "</correspondence>"

# properites=loadconfig()
# status,reason,finlavideocloudpath = generateVideoCommunication(inputXML,properites)
# print("status,reason,finlavideocloudpath",status,reason,finlavideocloudpath)
# exit(1)


def validateCredential(inputXML):
 
  status = "failed"
  reason = "Invalid XML Format"
  doc = xmltodict.parse(inputXML)
  reason = "Invalid userID password"
  #print("xmlFileNameTag", xmlFileNameTag)
  inputid = (doc['correspondence']['security']['id'])
  inputkey = (doc['correspondence']['security']['key'])
  if inputid == "citiusa" and  inputkey == "3GH7FGTUIOP8OIY":
    status = "passed"
    reason = "Valid userID password"
    #print (inputid,inputkey)
  elif inputid == "tcsdesign" and  inputkey == "4GH7XXXXOP8OIY":
    status = "passed"
    reason = "Valid userID password"
    #print (inputid,inputkey)
  else:
    status = "failed"
    reason = "Invalid userID password Please contact abc@tcs.com for setup"
  
  return (status, reason)

def validateTemplateDepoly(inputXML):
 
  status = "failed"
  reason = "Invalid Template password"
  doc = xmltodict.parse(inputXML)
  #print("xmlFileNameTag", xmlFileNameTag)
  inputtemplateid = (doc['correspondence']['templateID'])

  if inputtemplateid  == "CITI_POC_template_1001" or  inputtemplateid == "CITI_POC_template_1004":
    status = "passed"
    reason = "Valid Template ID"
    
  elif inputtemplateid == "CITI_POC_template_1002" or  inputtemplateid == "CITI_POC_template_1003":
    status = "passed"
    reason = "Valid userID password"
  else:
    status = "failed"
    reason = "Invalid Template ID , Please congifgure template before invoking"
  
  return (status ,reason)

def finaldump(videostatus,videoreason,videofilepath,processingfilename,jsonout):
    print("Final Dump",finalstatus,videofilepath + ":"  + videostatus + ":" + videoreason +  ":" + templateReason + ":" + credentialReason )
    dictionary={}
    dictionary["processingfilename"] = processingfilename
    dictionary["videostatus"] = videostatus
    dictionary["videoreason"] = videoreason
    dictionary["videofilepath"] = videofilepath
    print("dump json" +jsonout)
    with open(jsonout, "w") as outfile:
        json.dump(dictionary, outfile) 
    outfile.close()

def eCommsVideo(inputXML,processingfilename):
  
    try:
        videostatus,videoreason,videofilepath,jsonout = generateVideoCommunication(inputXML)
    except Exception as e:
        print("Oops!", e.__class__, "occurred.")
        print("Next entry.")
        print()  
        print ( " Trigger Failed" )
        finaldump("510",videostatus,videoreason,videofilepath,processingfilename,jsonout)

    finaldump("200",videostatus,videoreason,videofilepath,processingfilename,jsonout)
    return(videofilepath + ":"  + videostatus + ":" + videoreason +  ":" + processingfilename +  ":" + jsonout )

def generatevideo(inputXMLFile):
    if not inputXMLFile:
        return("Input XMLFile missing please provide input XML file usage generatevideo(inputXMLFile)");
    
    with open(inputXMLFile) as fd:
        inputXML = fd.read()
    fd.close()
     
    print ("start")
    
    base=os.path.basename(inputXMLFile)
    processingfilename=os.path.splitext(base)[0]
    print(eCommsVideo(inputXML,processingfilename))
    print("Completed")

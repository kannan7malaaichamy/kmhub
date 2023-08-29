import boto3
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler
import xmltodict
import time
import os
import sqlite3
configFile="C:/ezComms_Runtime/config/archiveS3/properties.xml"

def archiveS3(file,targetname):
    print("storing in AWS s3" + "........" )
    s3 = boto3.resource('s3')
    BUCKET = "publishing-platform"
    s3.Bucket(BUCKET).upload_file(file, targetname)
    

def loadconfig():
    with open(configFile) as fdxml:
        properites  = xmltodict.parse(fdxml.read())
    fdxml.close()
    #print("inside - config")
    return properites

def updateDB(x1,x2,x3,BUCKET,file) :
    try:
        sqllite_db_str='sqlite:////TCS Publishing Platform/DB/PP.db'
        conn = sqlite3.connect('C:/TCS Publishing Platform/DB/TCSPP.db')
        cursor = conn.cursor()
        #INSERT into ezArchiveS3(customerID,polcyID,LOB,S3BucketName,archiveFileName) values("32623","72815","LIFE","publishing-platform","COVERLETTER_8912_Tue04Jan2022091602_Tue04Jan2022091602.pdf");
        # Preparing SQL queries to INSERT a record into the database.
        #cur.execute("insert into contacts (name, phone, email) values (?, ?, ?)",(name, phone, email))
        cursor.execute('''INSERT INTO ezArchiveS3(customerID,polcyID,LOB,S3BucketName,archiveFileName) VALUES (?,?,?,?,?)''',(x1,x2,x3,BUCKET,file))
        conn.commit()
        conn.close()
    except Exception as e:
        print("DB Connetc filed")
        print(e)

patterns = "*"
ignore_patterns = ""
ignore_directories = False
case_sensitive = True
my_event_handler = PatternMatchingEventHandler(patterns, ignore_patterns, ignore_directories, case_sensitive)


def on_created(event):
    try:
        print(f" eComms BOT recevied  {event.src_path}!")
        properites = loadconfig()
        input_file = event.src_path
        #print("CallachiveS3  name only" )
        filename=os.path.basename(input_file)
        x = filename.split("_")
        #print("||||"+x[1]+x[3]+x[4])
        BUCKET="publishing-platform"
       
        archiveS3(input_file,filename)
        print("S3 archive done by the BOT! for the file:-" + input_file)
        updateDB(x[1],x[3],x[4],BUCKET,filename)
        
        
    
    except :
        print("Failed:- for the file")
        print(e)
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
watchdir  = properites['properties']['watchdir']
print("watchdir Folder " + watchdir)
go_recursively = True
my_observer = Observer()
my_observer.schedule(my_event_handler, watchdir, recursive=go_recursively)
my_observer.start()

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    my_observer.stop()
    my_observer.join()
except Exception as e:
    print("The Trigger  Failed:- for the file")
    print(e)
    pass

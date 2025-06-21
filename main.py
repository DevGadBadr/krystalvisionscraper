from selenium import webdriver
from myfunctions import *
import threading
from connect import connectAppToWebsocket

connectAppToWebsocket()

numberOfInstances = 4
chunkOfData = 10
tableName = 'krystalvision'

instances: list[dict] = []
createThreads: list[threading.Thread] = []
workingThreads: list[threading.Thread] = []
watchingThreads:list[threading.Thread] = []
mainDBConnection = connectDataBase()
mainDBConnection.autocommit = True
mainCursor = mainDBConnection.cursor()
mainCursor.execute(f"update {tableName} set istaken=false where isdone=false")
mainCursor.execute(f'select count (*) from {tableName}')
total = mainCursor.fetchall()[0][0]
mainCursor.execute(f'select count (*) from {tableName} where isdone=true')
done = mainCursor.fetchall()[0][0]

print(str(done) + ' / ' + str(total))

def createInstance(n):
    print(f'Making new webdriver window instance {n+1}...')
    newDriver= webdriver.Chrome(options=chrome_options)
    print(f'Making new connection to database instance {n+1}...')
    newConnection = connectDataBase()
    newConnection.autocommit = True
    newInstance = {"driver":newDriver,"connection":newConnection,"instanceID":n+1,"user":f"Perry{n+1}","workbook":f"KrystalVision{n+1}.xlsx"}
    instances.append(newInstance)
    
def watchDriver(instance):
    print(f'Watching Driver {instance["instanceID"]}')
    driver:WebDriver = instance["driver"]
    while True:
        screenShotName = f'./screenshots/driver-{instance["instanceID"]}-screenshot.png'
        try:
            driver.save_screenshot(screenShotName)
        except:
            print(f'Error taking screenshot in driver {instance["instanceID"]}')
        time.sleep(30)
            
def excuteMain(instance):
    driver:WebDriver = instance['driver']
    connection:ConnectionType = instance['connection']
    cursor = connection.cursor()
    returnData = chunkOfData
    cursor.execute(f'select count(*) from {tableName} where isdone=true')    

    sendMessage(f'Driver {instance["instanceID"]} - Krystal Vision Scraper Started')
    sendMessage(f'Driver {instance["instanceID"]} Logging In.')
    webProsLogIn(driver=driver,user=instance['user'])
    sendMessage(f'Driver {instance["instanceID"]} Entering Area.')
    enterTheArea(driver=driver)
            
    while returnData==chunkOfData:
        cursor.execute(f'with selected_records as (select * from {tableName} where istaken=false order by id asc limit {chunkOfData} for update skip locked) update {tableName} set istaken=true where id in (select id from selected_records) returning *')
        records = cursor.fetchall()
        records.sort()
        for record in records:
            print(f'Driver {instance["instanceID"]} is processing Invoice {record[1]}\n')
            time1 = time.time()
            (msg,sheet,row) = scrapeWebPros(driver=driver,record=record,workbook=instance['workbook'],instance=instance,cursor=cursor)
            time2 = time.time()
            timeElasped = round(time2-time1,2)
            if msg == 'success':
                cursor.execute(f'update {tableName} set isdone=true,isdoneimg=true,elapsedtime=%s,sheet=%s,row=%s where id=%s',(timeElasped,sheet,row,record[0]))
                print(f'Driver {instance["instanceID"]} - Values Updated For invoice {record[1]} at db row {record[0]}')
        returnData=len(records)
    print(f'Driver {instance["instanceID"]} Finished')
    
for n in range(numberOfInstances):
    newCereateThread = threading.Thread(target=createInstance,args=(n,))
    createThreads.append(newCereateThread)
    
for creatThread in createThreads:
    creatThread.start()
    
for creatThread in createThreads:
    creatThread.join()
    
for instance in instances:
    newWatchThread = threading.Thread(target=watchDriver,args=(instance,))
    watchingThreads.append(newWatchThread)
    
for thread in watchingThreads:
    thread.start()
    
for instance in instances:
    newWorkingThread = threading.Thread(target=excuteMain,args=(instance,))
    workingThreads.append(newWorkingThread)
    
for thread in workingThreads:
    thread.start()
    
for thread in workingThreads:
    thread.join()
    

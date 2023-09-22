# coding=utf-8
import os
import builtins
import xml.etree.ElementTree as ET
import csv, pyodbc
import threading
import tkinter as tk
from tkinter import filedialog

#access db connection
[x for x in pyodbc.drivers() if x.startswith('Microsoft Access Driver')]
DRV = '{Microsoft Access Driver (*.mdb, *.accdb)}'

#____________________________________________



#functions
def login(text):
    global log
    log.writelines(str(text)+"\r")
    print(str(text)+"\r")

def list_count(List,dev):
    try:
        is_present = List[0].index(dev)
        List[1][is_present] += 1
        login(str(List[1][is_present])+dev)
    except:
        List[0].append(dev)
        List[1].append(1)

def total_list(dev,out):      
    for i in range(len(dev[0])):
                dev_rep = find_replacement(dev[0][i],os.path.join("files","to_iq4_conv.csv"))
                out.writelines(str(dev[0][i])+","+str(dev[1][i]))
                if (dev_rep != ""): out.writelines(",replace by: "+dev_rep)
                out.writelines("\r")
  
def find_replacement(item,ref):
    #ref is a csv
    login("Looking for replacement controller")
    replacement = ""
    try:
        x = open(ref,'r')
    except OSError:
        err = "Could not open replacement list file"
        login(err)
    else:
        login("replacement list opened")
        reader = csv.DictReader(x)

        for row in reader:
            if (item == row["Original"] or item == row["Original"].upper() or item == row["Original"].lower()): 
                replacement = row["Replacement"]
                login("Replace by :"+replacement)
    finally:
        if replacement == "": login("No replacement was found")
        x.close()
        return replacement
    
def query_and_rec(query,cursor,output,total):
    login("SQL query : "+query)
    try:
        cursor.execute(query)
        result = cursor.fetchall()
    except:
        err = "Could not read database"
        login(err)
        output.writelines(err +"\r")
        nbresult = 0
    else:
        login(result)
        nbresult = len(result)
        login("Result obtained: "+str(nbresult))
        for i in range(nbresult):
            result[i] = str(result[i]).replace('(','').replace(')','')
            output.writelines(result[i]+"\r")
        if total : output.writelines("Total,"+str(nbresult)+"\r")
    return nbresult

#db is a string containing filepath + name of controller database 
def scan_controller(db,full_list,to_upgrade,outdir):

    login("Extracting database:"+db)
    
    upgraded = False
    
    #create sql connection
    # set up some constants
    MDB = db
    PWD = 'pw'
    
    # connect to db
    try:
        con = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV,MDB,PWD))
        cur = con.cursor()
    except:
        err = "Could not open database, make sure you have access and DB is not already in use by another program"
        login(err)
        device = "unknown"
    else:
        login("Logged in db")
    
        #get device info
        cur.execute('''SELECT lan FROM DeviceDetails;''')
        Lan = cur.fetchone()[0]
        login("Lan"+str(Lan))
            
        cur.execute('''SELECT node FROM DeviceDetails;''')
        OS = cur.fetchone()[0]
        login("OS"+str(OS))
            
        cur.execute('''SELECT device_name FROM DeviceDetails;''')
        dev_info = cur.fetchone()[0].split(" version ")
        login(dev_info)
        
        device  = dev_info[0]    
        version = dev_info[1] 
        login("Device: "+device)
        login("Version: "+version)
        
        #check replacement device
        up_dev = find_replacement(device,os.path.join("files","to_iq4_conv.csv"))
        if up_dev == "":
            replac_line = "No replacement found for controller" 
        else:
            replac_line = ("Replacement by "+up_dev+" suggested")
        login(replac_line)
            
        #in and out with absolute path
        output = os.path.join(outdir, "Lan"+str(Lan))
        login("Output controller info to: "+str(output))

        # Create output directory
        create_dir(output)
            
        #output
        out = os.path.join(output, "L"+str(Lan)+"n"+str(OS)+".csv")
        login(out)

        
        try:
            x = open(out,"w")
        except OSError:
            err = "Could not open controller output file, make sure it is not already in use by another program"
            login(err)
            x.writelines(err+"\r")
            controller_line = "Controller unknow"
        else:
            x.writelines("Device,"+device+",Version,"+version+","+replac_line+"\r")    
            

            #get sensor info
            x.writelines("Sensors list\r")   
            x.writelines("module_type,module_number,details,type\r")
            ai = query_and_rec('''SELECT * FROM IQSensor WHERE [type]=1 AND [details] IS NOT Null;''',cur,x,True)

            #get digin info
            x.writelines("Digital Inputs list\r")
            x.writelines("module_number ,type ,label_display ,label                          ,value_display ,value ,required_display ,required ,enabled_display ,enabled ,delay_display ,delay   ,status_alarm_group ,state_input ,hours_run_display ,starts_display ,state_text ,status_display ,subtype ,override_value ,override_value_enable ,override_value_display ,override_value_enable_display ,state_clear_enable\r")
            try:
                di = query_and_rec('''SELECT * FROM IQDiginType1 WHERE [subtype]=0 AND [label] IS NOT Null;''',cur,x,True)
            except:
                err = "No subtype property, check DI manually"
                login(err)
                x.writelines(err+"\r")
                di = query_and_rec('''SELECT * FROM IQDigin WHERE [type]=1 AND [details] IS NOT Null;''',cur,x,True)

            #get driver info
            x.writelines("Digital drivers list\r")
            x.writelines("module_type,module_number,details,type\r")
            do = query_and_rec('''SELECT * FROM IQDriver WHERE [type] = 1 AND [details] IS NOT Null;''',cur,x,True)
            x.writelines("Analog drivers list\r") 
            x.writelines("module_type,module_number,details,type\r")    
            ao = query_and_rec('''SELECT * FROM IQDriver WHERE [type] = 2 AND [details] IS NOT Null;''',cur,x,True)
            
            controller_line = (str(Lan)+","+str(OS)+","+device+","+version+","+str(di)+","+str(ai)+","+str(do)+","+str(ao)+","+up_dev+"\r") 
            
            #get IO Module info
            x.writelines("IO modules used\r")
            x.writelines("id,description,Module type\r")
            try:
                query_and_rec('''SELECT [id_number],[description],[io_type] FROM IOAssignment WHERE [io_type] IS NOT Null;''',cur,x,True)
            except:
                x.writelines("No IO modules possibility\r")          
        
        finally:
            #close output
            x.close()
            full_list.writelines(controller_line) 
            if not(device.startswith("IQ4") or device.startswith("iq4") or device.startswith("iqeco") or device.startswith("IQeco")) :
                to_upgrade.writelines(controller_line)
                upgraded = True       

            #close output
            cur.close()
            con.close()
            
    finally:
        login("Device "+device+" to be upgraded: "+str(upgraded))
        return [device,upgraded]

def controller_report(db,outdir):

    login("Extracting database:"+db)
    
    #create sql connection
    # set up some constants
    MDB = db
    PWD = 'pw'
    
    # connect to db
    try:
        con = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV,MDB,PWD))
        cur = con.cursor()
    except:
        err = "Could not open database, make sure you have access and DB is not already in use by another program"
        login(err)
        device = "unknown"
    else:
        login("Logged in db")
    
        #get device info
        cur.execute('''SELECT lan FROM DeviceDetails;''')
        Lan = cur.fetchone()[0]
        login("Lan"+str(Lan))
            
        cur.execute('''SELECT node FROM DeviceDetails;''')
        OS = cur.fetchone()[0]
        login("OS"+str(OS))
            
        cur.execute('''SELECT device_name FROM DeviceDetails;''')
        dev_info = cur.fetchone()[0].split(" version ")
        login(dev_info)
        
        device  = dev_info[0]    
        version = dev_info[1] 
        login("Device: "+device)
        login("Version: "+version)      
            
        #in and out with absolute path
        output = os.path.join(outdir, "Lan"+str(Lan))
        login("Output controller info to: "+str(output))

        # Create output directory
        create_dir(output)
            
        #output
        out = os.path.join(output, "L"+str(Lan)+"n"+str(OS)+".csv")
        login(out)

        
        try:
            x = open(out,"w")
        except OSError:
            err = "Could not open controller output file, make sure it is not already in use by another program"
            login(err)
            x.writelines(err+"\r")
            controller_line = "Controller unknow"
        else:
            x.writelines("Device,"+device+",Version,"+version+"\r")    
            

            #get sensor info
            x.writelines("Sensors list\r")   
            x.writelines("module_number,label,units,value,type\r")
            ai = query_and_rec('''SELECT [module_number],[label],[units],[value] FROM IQSensorType1 WHERE [type]=1 AND [label] IS NOT Null;''',cur,x,False)

            #get digin info
            x.writelines("Digital Inputs list\r")
            x.writelines("module_number,label,value,type\r")
            di = query_and_rec('''SELECT [module_number],[label],[value] FROM IQDiginType1 WHERE [subtype]=0 AND [label] IS NOT Null;''',cur,x,False)

            #get driver info
            x.writelines("Digital drivers list\r")
            x.writelines("module_number,label,value,type\r")
            do = query_and_rec('''SELECT [module_number],[label] FROM IQDriverType1 WHERE [type] = 1 AND [label] IS NOT Null;''',cur,x,False)
            x.writelines("Analog drivers list\r") 
            x.writelines("module_number,label,units,value,type\r") 
            ao = query_and_rec('''SELECT [module_number],[label] FROM IQDriverType2 WHERE [type] = 2 AND [label] IS NOT Null;''',cur,x,False)
            
            controller_line = (str(Lan)+","+str(OS)+","+device+","+version+","+str(di)+","+str(ai)+","+str(do)+","+str(ao)+"\r")         
        
        finally:
            #close output
            x.close()  
            cur.close()
            con.close()
            
    finally:
        pass
    
def create_dir(dirName):
    # Create output directory
    if not os.path.exists(dirName):
        os.mkdir(dirName)
        login("Directory " + dirName +  " Created ")
    else:    
        login("Directory " + dirName +  " already exists")

def check_dir(root):
    login ("testing root directory:" +rootdir)
    test = False
    for file in os.listdir(root):
        if file.endswith(".TSET"):
            test = True
    login(test)
    return test
    
def out_dir(dirName):
    dirName = os.path.normpath(rootdir).split(os.path.sep)
    dirName = dirName[len(dirName)-1]
    login(dirName)
    create_dir(dirName)
    return dirName
#____________________________________________
#GUI functions

def cancel():
    log.close()
    quit()

def db_scan():
    if check_rootfile(rootfile) is True:
        outdir = os.path.join(out_dir(rootfile),"Scan results")
        login(outdir)
        create_dir(outdir)
        try:
            open(os.path.join(outdir, "full_list.csv"),"w").close()                  
        except Exception as e:
            err = "Could not open list files, make sure you have all the required drivers installed (pyodbc, MS access 2010 framework distributabl)e it is not already in use by another program"
            login(err+"\r"+str(e))
        else:
            full = open(os.path.join(outdir, "full_list.csv"),"a")
            db_connect(rootfile,full,outdir)
            
            #list_heading = ("Lan,OS,Device,Version,DI,AI,DO,AO,Replace with\r")
            #full.writelines(list_heading)
            #tot_cont = 0
            #tot_up = 0
            #dev_list= [[],[]]
            #dev_up_list=[[],[]]

            #dev = scan_controller(rootfile,full,upgrade,outdir)   
            #device = dev[0].split()[0]                        
                
            #list_count(dev_list,device)
                    
                
            #full.writelines("Total:,"+str(tot_cont)+"\r")
            #total_list(dev_list,full)   
            #full.close()

            login('Scan complete')
    else:
        login("Not a valid mdb file")

def report():
    if check_rootfile(rootfile) is True:
        outdir = os.path.join(out_dir(rootfile),"Report starter")
        login(outdir)
        create_dir(outdir)
        # try:
            # open(os.path.join(outdir, "full_list.csv"),"w").close()                    
        # except Exception as e:
            # err = "Could not open list files, make sure it is not already in use by another program"
            # login(err+"\r"+str(e))
        # else:
        tot_cont = 0
        for subdir, dirs, files in os.walk(rootfile):
                for filename in files:
                    if (filename.endswith(".iq") or filename.endswith(".IQ")):#
                        tot_cont += 1
                        login("Reading controller:  "+str(tot_cont))
                        contr_db = os.path.join(subdir, filename)
                        login(contr_db)
                        controller_report(contr_db,outdir)
                        
        login('Reports complete')
    
def init_gui():

    #Tkinter GUI
    root = tk.Tk()
    root.title("Sigma DB scanner")
    root.minsize(400,100)
    root.geometry("480x100")
    
    # create the main sections of the layout, 
    # and lay them out
    top = tk.Frame(root)
    bottom = tk.Frame(root)
    top.pack(side=tk.TOP)
    bottom.pack(side=tk.BOTTOM, fill=tk.BOTH, expand=True)

    # create the widgets for the top part of the GUI,
    # and lay them out
    s = tk.Button(root, text="Select", width=10, height=2, command=select)
    c = tk.Button(root, text="Leave", width=10, height=2, command=cancel)
    e = tk.Button(root, text="Scan", wraplength=60, width=10, height=2, command=db_scan)
    #r = tk.Button(root, text="Start report", width=10, height=2, command=report)
    s.pack(in_=top, side=tk.LEFT)
    e.pack(in_=top, side=tk.LEFT)
    #r.pack(in_=top, side=tk.LEFT)
    c.pack(in_=top, side=tk.LEFT)

    # create the widgets for the bottom part of the GUI,
    # and lay them out
    global path
    path = tk.Label(root, text = "...", width=35, height=15)
    path.pack(in_=bottom, side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    w = tk.Label(root, text="Please choose SET Project file directory")
    w.pack()

    return root
    
def select():
    global rootfile
    login ("Please choose SET Project file directory")
    root.filename =  filedialog.askopenfile()
    rootfile = root.filename
#____________________________________________________________
#main code
global log
    
log = open("log.txt","w")

rootdir = ""
root = init_gui()
root.mainloop()


#-------------------------------------------------------------------------------
# This script is used for Parsing CM documents and Performing migration for ETL
# scripts
# Created: 16-Sep-2013
# Author: vvanarasi
#-------------------------------------------------------------------------------

import os
import getopt, sys
import csv
import win32com.client as win32
import subprocess
import getopt, sys
import cx_Oracle
import logging
import re
import getpass
import datetime
import xml.etree.cElementTree as ET
import paramiko
import time
import win32com.client
import shutil


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
## Add Logger

logger = logging.getLogger("test_logger")

timestr = time.strftime("%Y%m%d-%H%M%S")
global datestr
datestr = time.strftime("%Y%m%d")
logname = 'C:\Python27\examples\LOGS\CM_Automation_wrapper_'+timestr+'.log'

hdlr = logging.FileHandler(logname)
formatter = logging.Formatter('%(asctime)s %(levelname)s %(message)s')
hdlr.setFormatter(formatter)
logger.addHandler(hdlr) 
logger.setLevel(logging.DEBUG)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

def parse_doc():
    global cr_document
    global ddl_db_csv
    global unix_csv
    global informatica_csv
    global doc
    global cr_number    
    word = win32.Dispatch("Word.Application")
    word.Visible = 1
    word.Documents.Open(cr_document)
    doc = word.ActiveDocument
    cr_doc = os.path.basename(cr_document)
    cr_arr = cr_doc.split("_")
    cr_number = cr_arr[0]
    table = doc.Tables(3)
    cnt_rows = table.Rows.Count
    cnt_cols = table.Columns.Count

    outerarr = []
    for i in range(cnt_rows):
            innerarr = []
            for j in range(cnt_cols):
                    a=table.Cell(Row =i+1, Column =j+1).Range().encode('ascii','ignore').rstrip('\r\x07')
                    innerarr.append(a)
            outerarr.append(innerarr)


    ddl_db_csv = 'C:\\Python27\\examples\\CSV\\DDL_DB_'+ cr_number + '.csv'
    
    
    with open(ddl_db_csv, 'wb') as fp:
        a = csv.writer(fp, delimiter=',')
        a.writerows(outerarr)



    table = doc.Tables(9)
    cnt_rows = table.Rows.Count
    cnt_cols = table.Columns.Count

    ## Writing Informatica information to a csv file ##

    outerarr = []
    for i in range(cnt_rows):
            innerarr = []
            for j in range(cnt_cols):
                    a=table.Cell(Row =i+1, Column =j+1).Range().encode('ascii','ignore').rstrip('\r\x07')
                    innerarr.append(a)
            outerarr.append(innerarr)

    informatica_csv = 'C:\\Python27\\examples\\CSV\\INFORMATICA_' + cr_number + '.csv'
    

    with open(informatica_csv, 'wb') as fp:
        a = csv.writer(fp, delimiter=',')
        a.writerows(outerarr)


    table = doc.Tables(8)
    cnt_rows = table.Rows.Count
    cnt_cols = table.Columns.Count

    ## Writing Informatica information to a csv file ##

    outerarr = []
    for i in range(cnt_rows):
            innerarr = []
            for j in range(cnt_cols):
                    a=table.Cell(Row =i+1, Column =j+1).Range().encode('ascii','ignore').rstrip('\r\x07')
                    innerarr.append(a)
            outerarr.append(innerarr)

    
    unix_csv = 'C:\\Python27\\examples\\CSV\\UNIX_' + cr_number + '.csv'

    with open(unix_csv, 'wb') as fp:
        a = csv.writer(fp, delimiter=',')
        a.writerows(outerarr)


    

######################################################################
def validate(argv):
   global cr_document
   global envmt
   global env_var
   
   try:
     opts, args = getopt.getopt(argv,"hc:e:",["cr_document=","envmt=",])
   except getopt.GetoptError, err:
     print str(err) 
     print 'Usage: Automation_CM.py -c <cr_document>'
     sys.exit(2)
   for opt, arg in opts:
     if opt == '-h':
        print 'Usage: Automation_CM.py -c <cr_document>'
        sys.exit()
     elif opt in ("-c", "--cr_document"):
        print ""
        print "********************************** PARSING ARGUMENTS ****************************************"
        print ""
        print ""
        cr_document = arg
        print "ARGUMENTS PROVIDED:"
        print ""
        print "CR DOCUMENT :" + cr_document
     elif opt in ("-e", "--envmt"):
        envmt = arg
        env_var = envmt
        print ""
        print "ENVIRONMENT:" + envmt
     else :
           assert False, "unhandled option"
           

#####################################################################


def db_connect():
   global db
   global cursor
   global username
   global password
   global tnsname 
   global sqlfiles
   
   """ Connect to the database. """

   try:
       db = cx_Oracle.connect(username, password, tnsname)
   except cx_Oracle.DatabaseError as e:
       error, = e.args
       if error.code == 1017:
           print('Please check your credentials.')
           logger.error('Error with Credentials supplied')
       else:
           print('Database connection error: %s'.format(e))
           logger.error('Database connection error: %s',format(e))
           logger.error('Oracle Error Code:%n',error.code)
           logger.error('Oracle Error Message:%s',error.message)

       # Very important part!
       raise

   cursor = db.cursor()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
def db_disconnect():
   global db
   global cursor
   """
   Disconnect from the database. Handle exception during closure
   """

   try:
       cursor.close()
       db.close()
   except cx_Oracle.DatabaseError:
       logger.error('Oracle Error Occured while disconnect')
       sys.exit(1)
       


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   
def process_sql():
   global db
   global cursor
   
   ## Processing the SQL's Sequentially
   lst_sql = sqlfiles.split(",")

   for sql_file in lst_sql:
      f = open(sql_file)
      full_sql = f.read()
      sql_commands = full_sql.split(';')
   
      for item in sql_commands:
         if re.match('\s+',item):
            sql_commands.remove(item)


      
      for sql_command in sql_commands:
         print sql_command
         logger.info('SQL EXECUTED:%s',sql_command)
         if sql_command != '' :
            try:
               cursor.execute(sql_command)
            except cx_Oracle.DatabaseError as e:
               error, = e.args
               print sql_command
               print(error.code)
               print(error.message)
               print(error.context)
               logger.error('Oracle Error Code:%n',error.code)
               logger.error('Oracle Error Message:%s',error.message)
               logger.error('Oracle Error Context:%s',error.context)
               sys.exit(1)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

######################################################################
def call_sql():
    print ""
    print "***********************************************************************************************"
    print "*                                                                                             *"
    print "*                                    SQL MIGRATION                                            *"
    print "*                                                                                             *"
    print "***********************************************************************************************"

    
    global username
    global password
    global tnsname 
    global sqlfiles
    global sql_status
    sql_cnt = 0

    from collections import defaultdict
    columns = defaultdict(list) 
    
    with open(ddl_db_csv) as f:
        reader = csv.DictReader(f) 
        for row in reader: 
            for (k,v) in row.items(): 
                columns[k].append(v)

    for i in range(len(columns['Script name'])):
        sql_name = (columns['Script name'])[i]
        schema_name = (columns['Schema Name'])[i]
        database_name = (columns['Database'])[i]
        comments = (columns['DB Type'])[i]
        

        if (i > 0):
            prev_db_name = (columns['Database'])[i-1]
            
        if (database_name != '' and comments == 'Oracle'):
            sql_cnt += 1
            username = schema_name
            db_nm = database_name.split("/")
            
            if env_var == 'QA':
                row_db_nm = db_nm[0] + '.world'
            elif env_var == 'PROD':
                row_db_nm = db_nm[1] + '.world'
            else:
                assert False, "Invalid Data"
            
            if (i == 0):
                print username
                password = getpass.getpass()
                prev_passwd = password                
            elif ((i > 0) and (database_name != prev_db_name)):
                print username
                password = getpass.getpass()
                prev_passwd = password                
            else:
                password = prev_passwd

            tnsname = row_db_nm
            sqlfiles = sql_name
            logger.info('*****Opening DB Connection*****')
            logger.info('*****Processing SQL*****')
            try:
                db_connect()
            except:
                logger.info('*****Error in Opening DB Connection*****')
                sys.exit(1)

            try:
                process_sql()
            except:
                logger.info('*****Error in Processing SQL*****')
                sys.exit(1)

            try:
                db_disconnect()
                logger.info('*****Closing DB Connection *****')
            except:
                logger.info('*****Error in Closing DB Connection *****')
                sys.exit(1)

        else:
            logger.info('***********-----------******************')

    if sql_cnt >= 1:
        sql_status = '#5DC83C'
    else:
        sql_status = 'Yellow'

    return True


#####################################################################

def generate_xml():
    print ""
    print "***********************************************************************************************"
    print "*                                                                                             *"
    print "*                                    XML MIGRATION                                            *"
    print "*                                                                                             *"
    print "***********************************************************************************************"
    global varreplace 
    global varreuse 
    global varobjecttype
    global varlist
    global envmt
    global xml_name
    global xml_status
    global local_xml_log
    xml_cnt = 0

    varreplace = 'REPLACE'
    varreuse = 'REUSE'
    
    if envmt == 'DEV':
        var_envmt = ''
        unix_srvr = ''
    elif envmt == 'QA':
        var_envmt = ''
        unix_srvr = ''
    elif envmt == 'PROD':
        var_envmt = ''
        unix_srvr = ''
    else:
        assert False, "Invalid Environment"



    ### Start Building XML Dynamixally ####
    tree = ET.ElementTree(file='C:\Python27\examples\sample.xml')
    root = tree.getroot()
    vartemp = root.find( "RESOLVECONFLICT" )

    for mod in vartemp:
        if mod.tag == "SPECIFICOBJECT":
            mod.attrib['REPOSITORYNAME'] = ''
    
    table = doc.Tables(9)
    cnt_rows = table.Rows.Count
    cnt_cols = table.Columns.Count
    cr_number_dup = table.Cell(Row=2, Column=1).Range.Text

    if (cr_number_dup != ''):
        print ("***** Credentials to Connect to Informatica Workflow Manager *****")
        print ("Enter Username to Connect to ") +  var_envmt
        pmrep_username = raw_input("Enter Username in Upper Case:")
        pmrep_passwd = getpass.getpass()
        
    print ""
    port = 22        
    transport = paramiko.Transport((unix_srvr, port))
    transport.connect(username = unix_user_name, password = unix_password)
    sftp = paramiko.SFTPClient.from_transport(transport)

    for row in xrange(2, cnt_rows + 1):
        cr_num = table.Cell(Row=row, Column=1).Range.Text
        envmt = table.Cell(Row=row, Column=2).Range.Text
        order = table.Cell(Row=row, Column=3).Range.Text
        folder_name = table.Cell(Row=row, Column=4).Range().encode('ascii','ignore').rstrip('\r\x07')
        object_name = table.Cell(Row=row, Column=5).Range().encode('ascii','ignore').rstrip('\r\x07')
        object_mapping = table.Cell(Row=row, Column=6).Range().encode('ascii','ignore').replace("\r","").rstrip('\r\x07')
        tool_name = table.Cell(Row=row, Column=7).Range().encode('ascii','ignore').replace("\r","").rstrip('\r\x07')
        initial_rerun = table.Cell(Row=row, Column=8).Range().encode('ascii','ignore').rstrip('\r\x07')
        comments = table.Cell(Row=row, Column=9).Range.Text
        xml_name = object_name
        if (object_mapping != '' and tool_name == 'Repository Manager'):
            if ((initial_rerun != 'Completed')):
                xml_cnt += 1
                for child in root:
                    if child.tag == "FOLDERMAP":
                        child.attrib['TARGETREPOSITORYNAME'] = var_envmt
                        if child.attrib['SOURCEFOLDERNAME']=="SPIDER":
                                child.attrib['TARGETFOLDERNAME'] = folder_name
                                child.attrib['SOURCEFOLDERNAME'] = folder_name
                    vartemp = root.find( "RESOLVECONFLICT" )
                    for mod in vartemp:
                        if mod.tag == "SPECIFICOBJECT" and mod.attrib['FOLDERNAME'] == 'SPIDER':
                            mod.attrib['FOLDERNAME'] = folder_name
                object_mapping_srch1 = re.search('Shared(.*)Object folder',object_mapping)
                if object_mapping_srch1:
                    object_mapping_srch2 = object_mapping_srch1.group(1)
                    object_mapping_srch3 = re.search('Source definition: <Name of the objects>(.*)Target definition',object_mapping_srch2)
                    if object_mapping_srch3:
                        object_mapping_srch4 = object_mapping_srch3.group(1)
                        logger.info( "\n" )
                        logger.info ( 'Shared Objects---Source Definitions' )
                        logger.info ( object_mapping_srch4 ) ### between source and target definitions-----gives source definitions in shared objects
                        varlist = object_mapping_srch4.split(',')
                        for i in varlist:
                            if i != '':
                                ET.SubElement(vartemp, "SPECIFICOBJECT", {"RESOLUTION": varreplace,"REPOSITORYNAME":'REP_SVC_BI_DEV', "OBJECTTYPENAME":'Source definition' , "FOLDERNAME":'Shared' , "NAME":i})
                    object_mapping_srch5 = re.search('Target definition: <Name of the objects>(.*)MAPPING:  <Name of the objects>',object_mapping_srch2)
                    if object_mapping_srch5:
                        object_mapping_srch6 = object_mapping_srch5.group(1)
                        logger.info( "\n" )
                        logger.info ( 'Shared Objects---Target Definitions' )
                        logger.info ( object_mapping_srch6 ) ### between source and target definitions----- gives target definitions in shared objects
                        varlist = object_mapping_srch6.split(',')
                        for i in varlist:
                            if i != '':
                                ET.SubElement(vartemp, "SPECIFICOBJECT", {"RESOLUTION": varreplace,"REPOSITORYNAME":'REP_SVC_BI_DEV', "OBJECTTYPENAME":'Target definition' , "FOLDERNAME":'Shared' , "NAME":i})
                    object_mapping_srch27 = re.search('MAPPING:  <Name of the objects>(.*)',object_mapping_srch2)
                    if object_mapping_srch27:
                        object_mapping_srch28 = object_mapping_srch27.group(1)
                        logger.info( "\n" )
                        logger.info( 'Shared Objects---Mappings' )
                        logger.info( object_mapping_srch28 ) ###  gives Mappings in shared objects
                        varlist = object_mapping_srch28.split(',')
                        for i in varlist:
                            if i != '':
                                ET.SubElement(vartemp, "SPECIFICOBJECT", {"RESOLUTION": varreplace,"REPOSITORYNAME":'REP_SVC_BI_DEV', "OBJECTTYPENAME":'Mapping' , "FOLDERNAME":'Shared' , "NAME":i})
                object_mapping_srch_second = re.search('Object folder(.*)Source definition: <Name of the objects>',object_mapping)
                if object_mapping_srch_second:
                    object_mapping_srch_last = object_mapping_srch_second.group(1)
                    object_mapping_srch7 = re.search('WORKFLOW :(.*)WORKLET :',object_mapping_srch_last)
                    if object_mapping_srch7:
                        object_mapping_srch8 = object_mapping_srch7.group(1)
                        logger.info( "\n" )
                        logger.info( 'Object Folder---WORKFLOWS' )
                        logger.info( object_mapping_srch8 )
                        varlist = object_mapping_srch8.split(',')
                        for i in varlist:
                            if i != '':
                                ET.SubElement(vartemp, "SPECIFICOBJECT", {"RESOLUTION": varreuse,"REPOSITORYNAME":'REP_SVC_BI_DEV', "OBJECTTYPENAME":'Workflow' , "FOLDERNAME":folder_name , "NAME":i})
                    object_mapping_srch9 = re.search('WORKLET :(.*)SESSION :',object_mapping_srch_last)
                    if object_mapping_srch9:
                        object_mapping_srch10 = object_mapping_srch9.group(1)
                        logger.info(  "\n" )
                        logger.info(  'Object Folder---WORKLETS' )
                        logger.info(  object_mapping_srch10 )
                        varlist = object_mapping_srch10.split(',')
                        for i in varlist:
                            if i != '':
                                ET.SubElement(vartemp, "SPECIFICOBJECT", {"RESOLUTION": varreuse,"REPOSITORYNAME":'REP_SVC_BI_DEV', "OBJECTTYPENAME":'Worklet' , "FOLDERNAME":folder_name , "NAME":i})
                    object_mapping_srch11 = re.search('SESSION :(.*)MAPPING:',object_mapping_srch_last)
                    if object_mapping_srch11:
                        object_mapping_srch12 = object_mapping_srch11.group(1)
                        logger.info( "\n" )
                        logger.info( 'Object Folder---SESSIONS' )
                        logger.info( object_mapping_srch12 )
                        varlist = object_mapping_srch12.split(',')
                        for i in varlist:
                            if i != '':
                                ET.SubElement(vartemp, "SPECIFICOBJECT", {"RESOLUTION": varreuse,"REPOSITORYNAME":'REP_SVC_BI_DEV', "OBJECTTYPENAME":'Session' , "FOLDERNAME":folder_name , "NAME":i})
                    object_mapping_srch13 = re.search('MAPPING:(.*)MAPPLET:',object_mapping_srch_last)
                    if object_mapping_srch13:
                        object_mapping_srch14 = object_mapping_srch13.group(1)
                        logger.info( "\n" )
                        logger.info( 'Object Folder---MAPPINGS' )
                        logger.info( object_mapping_srch14 )
                        varlist = object_mapping_srch14.split(',')
                        for i in varlist:
                            if i != '':
                                ET.SubElement(vartemp, "SPECIFICOBJECT", {"RESOLUTION": varreuse,"REPOSITORYNAME":'REP_SVC_BI_DEV', "OBJECTTYPENAME":'Mapping' , "FOLDERNAME":folder_name , "NAME":i})
                    object_mapping_srch15 = re.search('MAPPLET:(.*)CONFIG:',object_mapping_srch_last)
                    if object_mapping_srch15:
                        object_mapping_srch16 = object_mapping_srch15.group(1)
                        logger.info( "\n" )
                        logger.info( 'Object Folder---MAPPLETS' )
                        logger.info( object_mapping_srch16 )
                        varlist = object_mapping_srch16.split(',')
                        for i in varlist:
                            if i != '':
                                ET.SubElement(vartemp, "SPECIFICOBJECT", {"RESOLUTION": varreuse,"REPOSITORYNAME":'REP_SVC_BI_DEV', "OBJECTTYPENAME":'Mapplet' , "FOLDERNAME":folder_name , "NAME":i})
                    object_mapping_srch17 = re.search('CONFIG:(.*)TASK:',object_mapping_srch_last)
                    if object_mapping_srch17:
                        object_mapping_srch18 = object_mapping_srch17.group(1)
                        logger.info( "\n" )
                        logger.info( 'Object Folder---CONFIG' )
                        logger.info( object_mapping_srch18 )
                        varlist = object_mapping_srch18.split(',')
                        for i in varlist:
                            if i != '':
                                ET.SubElement(vartemp, "SPECIFICOBJECT", {"RESOLUTION": varreuse,"REPOSITORYNAME":'REP_SVC_BI_DEV', "OBJECTTYPENAME":'Config' , "FOLDERNAME":folder_name , "NAME":i})
                    object_mapping_srch19 = re.search('TASK:(.*)TRANSFORMATION:',object_mapping_srch_last)
                    if object_mapping_srch19:
                        object_mapping_srch20 = object_mapping_srch19.group(1)
                        logger.info(  "\n" )
                        logger.info(  'Object Folder---TASK' )
                        logger.info(  object_mapping_srch20 )
                        varlist = object_mapping_srch20.split(',')
                        for i in varlist:
                            if i != '':
                                ET.SubElement(vartemp, "SPECIFICOBJECT", {"RESOLUTION": varreuse,"REPOSITORYNAME":'REP_SVC_BI_DEV', "OBJECTTYPENAME":'Task' , "FOLDERNAME":folder_name , "NAME":i})
                    object_mapping_srch21 = re.search('TRANSFORMATION:(.*)',object_mapping_srch_last)
                    if object_mapping_srch21:
                        object_mapping_srch22 = object_mapping_srch21.group(1)
                        logger.info( "\n" )
                        logger.info( 'Object Folder---TRANSFORMATION' )
                        logger.info( object_mapping_srch22 )
                        varlist = object_mapping_srch22.split(',')
                        for i in varlist:
                            if i != '':
                                ET.SubElement(vartemp, "SPECIFICOBJECT", {"RESOLUTION": varreuse,"REPOSITORYNAME":'REP_SVC_BI_DEV', "OBJECTTYPENAME":'Transformation' , "FOLDERNAME":folder_name , "NAME":i})
                    object_mapping_srch_third = re.search('Object folder(.*)',object_mapping)
                    if object_mapping_srch_third:
                        object_mapping_srch_middle = object_mapping_srch_third.group(1)
                        object_mapping_srch23 = re.search('Source definition: <Name of the objects>(.*)Target definition: <Name of the objects>',object_mapping_srch_middle)
                        if object_mapping_srch23:
                            object_mapping_srch24 = object_mapping_srch23.group(1)
                            logger.info( "\n" )
                            logger.info( 'Object Folder---Source Definitions' )
                            logger.info( object_mapping_srch24 )
                            varlist = object_mapping_srch24.split(',')
                            for i in varlist:
                                if i != '':
                                    ET.SubElement(vartemp, "SPECIFICOBJECT", {"RESOLUTION": varreuse,"REPOSITORYNAME":'REP_SVC_BI_DEV', "OBJECTTYPENAME":'Source definition' , "FOLDERNAME":folder_name , "NAME":i})
                        object_mapping_srch25 = re.search('Target definition: <Name of the objects>(.*)',object_mapping_srch_middle)
                        if object_mapping_srch25:
                            object_mapping_srch26 = object_mapping_srch25.group(1)
                            logger.info( "\n" )
                            logger.info( 'Object Folder---Target Definitions' )
                            logger.info( object_mapping_srch26 )
                            varlist = object_mapping_srch26.split(',')
                            for i in varlist:
                                if i != '':
                                    ET.SubElement(vartemp, "SPECIFICOBJECT", {"RESOLUTION": varreuse,"REPOSITORYNAME":'REP_SVC_BI_DEV', "OBJECTTYPENAME":'Target definition' , "FOLDERNAME":folder_name , "NAME":i})
                input_xml = 'C:\Python27\examples\XML\sample_'+timestr+'.xml'
                with open(input_xml, 'w') as f:
                    f.write('<?xml version="1.0" encoding="UTF-8" ?><!DOCTYPE IMPORTPARAMS SYSTEM "/infoapp/9.1.0/server/bin/impcntl.dtd">')
                    tree.write(f, 'utf-8')
                print ""    
                print "****************************************************************************************"
                print ""    
                print "Shared Objects specified in the CR document for REPLACE option:"
                if object_mapping_srch4 == "" and object_mapping_srch6 == "" and object_mapping_srch28 == "":
                    print "No Objects Specified"
                    print ""    
                else:
                    print object_mapping_srch4
                    print object_mapping_srch6
                    print object_mapping_srch28
                    print ""    
                print "Non Shared Objects specified in the CR document for REUSE option:"
                if object_mapping_srch8 == "" and object_mapping_srch10 == "" and object_mapping_srch12 == "" and object_mapping_srch14 == "" and object_mapping_srch16 == "" and object_mapping_srch18 == "" and object_mapping_srch20 == "" and object_mapping_srch22 == "" and object_mapping_srch24 == "" and object_mapping_srch26 == "":
                    print "No Objects Specified"
                    print ""    
                else:
                    print object_mapping_srch8
                    print object_mapping_srch10
                    print object_mapping_srch12
                    print object_mapping_srch14
                    print object_mapping_srch16
                    print object_mapping_srch18
                    print object_mapping_srch20
                    print object_mapping_srch22
                    print object_mapping_srch24
                    print object_mapping_srch26
                    print ""    
                print ""
                print ""
                print "****************************************************************************************"
                print ""
                print 'Connecting to Unix Server ' + unix_srvr + ' to copy the control file and the input cml file(s) '
                print ""
                control_xml = ''+cr_number+'_ctl_file_'+datestr+'.xml'
                print ""
                print "Files copied to unix server via sftp : "
                print ""
                print input_xml + "------>" + control_xml
                logger.info('***** Copying XML Control file to Unix Server *****')
                sftp.put(input_xml, control_xml)
                ftp_xml_file = object_name.split("\\").pop()
                ftp_xml = '' + ftp_xml_file
                logger.info('***** Copying Input XML file to Unix Server *****')
                print ""
                print object_name + "------>" + ftp_xml
                print ""
                sftp.put(object_name,ftp_xml)
                print 'Connecting to Repository Manager '
                print ""
                logger.info('***** Calling the Shell script - pmrep utility *****')
                print ""
                str = 'D:\plink.exe -ssh -pw ' +  unix_password + ' ' + unix_user_name + '@' + unix_srvr + ' sh utility_pmrep.sh ' + var_envmt + ' ' + pmrep_username + ' ' + pmrep_passwd + ' ' + ftp_xml + ' ' + control_xml + ' ' + cr_number
                os.system(str)
                xml_base_log_name = cr_number + '_pmrep_utility' + '.log'
                xml_log = '' + xml_base_log_name
                local_xml_log = 'C:\\Python27\\examples\\PMREP_LOGS\\' + xml_base_log_name
                print ""
                print ""
                sftp.get(xml_log,local_xml_log)
            else:
                logger.info('***** Not Processing becaue its rerun ****** ')
                print "Not Processing becaue its rerun"
                print ""
        elif (object_mapping != '' and tool_name == 'Workflow Manager'):
            if ((initial_rerun != 'Completed')):
                ## Start Building XML Dynamixally ####
                tree = ET.ElementTree(file='C:\Python27\examples\sample_wfm.xml')
                root = tree.getroot()
                for child in root:
                    if child.tag == "FOLDERMAP":
                        child.attrib['TARGETREPOSITORYNAME'] = var_envmt
                        ## Needs to be Modified after Testing - During the Actual Run - change source repository - start
                        #child.attrib['SOURCEREPOSITORYNAME'] = var_envmt
                        ## Needs to be Modified after Testing - During the Actual Run - change source repository - end
                        if child.attrib['SOURCEFOLDERNAME']=="FOLDER_NAME":
                                child.attrib['TARGETFOLDERNAME'] = folder_name
                                child.attrib['SOURCEFOLDERNAME'] = folder_name

                input_wfm_xml = 'C:\Python27\examples\XML\sample_wfm_'+timestr+'.xml'
                with open(input_wfm_xml, 'w') as f:
                    f.write('<?xml version="1.0" encoding="UTF-8" ?><!DOCTYPE IMPORTPARAMS SYSTEM "/infoapp/9.1.0/server/bin/impcntl.dtd">')
                    tree.write(f, 'utf-8')

                print 'Connecting to Unix Server ' + unix_srvr + ' to copy the control file and the input cml file(s) '
                print ""
                control_wfm_xml = ''+cr_number+'_ctl_file_wfm_'+datestr+'.xml'
                print ""
                print "Files copied to unix server via sftp : "
                print ""
                print input_wfm_xml + "------>" + control_wfm_xml
                logger.info('***** Copying XML Control file to Unix Server *****')
                sftp.put(input_wfm_xml, control_wfm_xml)
                ftp_xml_file = object_name.split("\\").pop()
                ftp_xml = '' + ftp_xml_file
                logger.info('***** Copying Input XML file to Unix Server *****')
                print ""
                print object_name + "------>" + ftp_xml
                print ""
                sftp.put(object_name,ftp_xml)
                print 'Connecting to Workflow Manager '
                print ""
                logger.info('***** Calling the Shell script - pmrep utility *****')
                print ""
                str = 'D:\plink.exe -ssh -pw ' +  unix_password + ' ' + unix_user_name + '@' + unix_srvr + ' sh utility_pmrep.sh ' + var_envmt + ' ' + pmrep_username + ' ' + pmrep_passwd + ' ' + ftp_xml + ' ' + control_wfm_xml + ' ' + cr_number
                os.system(str)
                xml_base_log_name = cr_number + '_pmrep_utility' + '.log'
                xml_log = '' + xml_base_log_name
                local_xml_log = 'C:\\Python27\\examples\\PMREP_LOGS\\' + xml_base_log_name
                print ""
                print ""
                sftp.get(xml_log,local_xml_log)

        else:
            logger.info('***** --------------- *****')

    sftp.close()
    if xml_cnt >= 1:
        xml_status = '#5DC83C'
    else:
        xml_status = 'Yellow'

    return True

#####################################################################

def unix_migration():
    global unix_status
    unix_cnt = 0
    print ""
    print "***********************************************************************************************"
    print "*                                                                                             *"
    print "*                                    UNIX MIGRATION                                           *"
    print "*                                                                                             *"
    print "***********************************************************************************************"


    from collections import defaultdict
    columns = defaultdict(list) 


    if envmt == 'DEV':
        var_envmt = ''
    elif envmt == 'QA':
        var_envmt = ''
    elif envmt == 'PROD':
        var_envmt = ''
    else:
        assert False, "Invalid Environment"

    port = 22
    transport = paramiko.Transport((var_envmt, port))
    transport.connect(username = unix_user_name, password = unix_password)
    sftp = paramiko.SFTPClient.from_transport(transport)
    print ""
    print 'Connecting to ' + env_var + ' Unix Server ' + var_envmt + ' to copy the specified file(s) in CR document'
    print "Files copied to unix server via sftp : "
    print ""

    with open(unix_csv) as f:
        reader = csv.DictReader(f) 
        for row in reader: 
            for (k,v) in row.items(): 
                columns[k].append(v)
    
    for i in range(len(columns['File Name'])):
        folder_file_name = (columns['File Name'])[i]
        permissions = (columns['File Permissions'])[i]
        comments = (columns['Comments'])[i]
        local_path = (columns['Folder Path/Name'])[i]
        path = folder_file_name
        path1 = path.split("/")
        path1.pop()
        remote_path = "/".join(path1)
        
        if  folder_file_name != '':
            unix_cnt += 1
            print local_path + "------>" + folder_file_name
            logger.info(local_path + "------>" + folder_file_name)
            try:
                sftp.chdir(remote_path)  # Test if remote_path exists
            except IOError:
                sftp.mkdir(remote_path)  # Create remote_path
                sftp.chdir(remote_path)
            sftp.put(local_path, folder_file_name)    # At this point, you are in remote_path in either case
            
        else:
            logger.info('****** Blank Lines in Unix Migration Table*****')

    sftp.close()
    if unix_cnt >= 1:
        unix_status = '#5DC83C'
    else:
        unix_status = 'Yellow'

    return True 
            

#####################################################################

def send_mail_via_com():
    olMailItem = 0x0
    obj = win32com.client.Dispatch("Outlook.Application")
    newMail = obj.CreateItem(olMailItem)
    newMail.Subject = 'CM Automation-[ ' + cr_number + ' ]-Env-' + env_var + ' -Auto Notification- ' + datestr
    attachment1 = logname
    newMail.Attachments.Add(attachment1)


    if local_xml_log != '':
        attachment2 = local_xml_log
        newMail.Attachments.Add(attachment2)
    if local_netezza_log != '':
        attachment3 = local_netezza_log
        newMail.Attachments.Add(attachment3)

    StrMailStart =	"<!DOCTYPE HTML><HTML><HEAD><STYLE>TABLE{BORDER-COLLAPSE:COLLAPSE;}" \
		"TABLE,TD{BORDER:1px SOLID BLACK;width:150} </STYLE></HEAD><BODY>" \
		"<p><FONT SIZE=2 FACE=CALIBRI><h3>Migration Summary:</h3></FONT></p>" \
		"<p><FONT SIZE=2 FACE=CALIBRI>Legend:</FONT></p>" \
		"<TABLE>" \
		"<TR>" \
		"<TD ALIGN=CENTER style=width:50><FONT SIZE=2 FACE=CALIBRI> Migration Status: </FONT></TD>" \
		"<TD ALIGN=CENTER style=width:50><FONT SIZE=2 FACE=CALIBRI> Successful </FONT></TD>" \
		"<TD ALIGN=CENTER style=width:50><FONT SIZE=2 FACE=CALIBRI> Failed </FONT></TD>" \
		"<TD ALIGN=CENTER style=width:50><FONT SIZE=2 FACE=CALIBRI> Not Applicable </FONT></TD>" \
		"</TR>" \
		"<TR>" \
		"<TD ALIGN=CENTER style=width:50 BGCOLOR=White><FONT SIZE=2 FACE=CALIBRI>Colour</FONT></TD>" \
		"<TD ALIGN=CENTER style=width:50 BGCOLOR=#5DC83C><FONT SIZE=2 FACE=CALIBRI>Green</FONT></TD>" \
		"<TD ALIGN=CENTER style=width:50 BGCOLOR=Red><FONT SIZE=2 FACE=CALIBRI>Red</FONT></TD>" \
		"<TD ALIGN=CENTER style=width:50 BGCOLOR=Yellow><FONT SIZE=2 FACE=CALIBRI>Yellow</FONT></TD>" \
		"</TR></TABLE><p><br></p>"
    StrMailMiddle =     		"<TABLE width = 300%>" \
		"<TR>" \
		"<TH ALIGN=CENTER style=width:50><FONT SIZE=2 FACE=CALIBRI> CR_NUMBER: </FONT></TH>" \
		"<TH colspan= 5 ALIGN=CENTER ><FONT SIZE=2 FACE=CALIBRI> CR12345 </FONT></TH>" \
		"</TR>" \
		"<TR>" \
		"<TD ALIGN=CENTER  style=width:150 BGCOLOR=White><FONT SIZE=2 FACE=CALIBRI>Migration Type</FONT></TD>" \
		"<TD ALIGN=CENTER  style=width:50 BGCOLOR=White><FONT SIZE=2 FACE=CALIBRI>Informatica</FONT></TD>" \
		"<TD ALIGN=CENTER  style=width:50 BGCOLOR=White><FONT SIZE=2 FACE=CALIBRI>Unix</FONT></TD>" \
		"<TD ALIGN=CENTER  style=width:50 BGCOLOR=White><FONT SIZE=2 FACE=CALIBRI>SQL</FONT></TD>" \
		"<TD ALIGN=CENTER  style=width:50 BGCOLOR=White><FONT SIZE=2 FACE=CALIBRI>Netezza</FONT></TD>" \
		"<TD ALIGN=CENTER  style=width:50 BGCOLOR=White><FONT SIZE=2 FACE=CALIBRI>Active Batch</FONT></TD>" \
		"</TR>" \
		"<TR>" \
		"<TD ALIGN=CENTER  BGCOLOR=White><FONT SIZE=2 FACE=CALIBRI>Status</FONT></TD>" \
		'<TD ALIGN=CENTER  BGCOLOR=' + xml_status + '><FONT SIZE=2 FACE=CALIBRI></FONT></TD>' \
		'<TD ALIGN=CENTER  BGCOLOR=' + unix_status + '><FONT SIZE=2 FACE=CALIBRI></FONT></TD>' \
		'<TD ALIGN=CENTER  BGCOLOR=' + sql_status + '><FONT SIZE=2 FACE=CALIBRI></FONT></TD>' \
		'<TD ALIGN=CENTER  BGCOLOR=' + netezza_status + '><FONT SIZE=2 FACE=CALIBRI></FONT></TD>' \
		'<TD ALIGN=CENTER  BGCOLOR=' + abatch_status + '><FONT SIZE=2 FACE=CALIBRI></FONT></TD>' \
		"</TR>" \
		"<TR>" \
		"<TD ALIGN=CENTER  BGCOLOR=White><FONT SIZE=2 FACE=CALIBRI>Errors</FONT></TD>" \
		"<TD ALIGN=CENTER colspan= 5 BGCOLOR=White><FONT SIZE=2 FACE=CALIBRI>NIL</FONT></TD>" \
		"</TR></TABLE><p><br></p>"
    StrMailEnd =	"</BODY></HTML>"

    newMail.HTMLBody = StrMailStart + StrMailMiddle + StrMailEnd
    newMail.To = ""
    newMail.CC = ""
    newMail.Send()

    return True

#####################################################################

def netezza_migration():
    global netezza_status
    global env_var
    global local_netezza_log
    
    netezza_cnt = 0
    print ""
    print "***********************************************************************************************"
    print "*                                                                                             *"
    print "*                                    NETEZZA MIGRATION                                        *"
    print "*                                                                                             *"
    print "***********************************************************************************************"


    from collections import defaultdict
    columns = defaultdict(list) 


    if env_var == 'DEV':
        var_envmt = ''
        netezza_srvr = ''
    elif env_var == 'QA':
        var_envmt = ''
        netezza_srvr = ''
    elif env_var == 'PROD':
        var_envmt = ''
        netezza_srvr = ''
    else:
        assert False, "Invalid Environment"

    print('In Netezza Migration:')
    print ""
    print 'Connecting to ' + env_var + ' Unix Server ' + var_envmt + ' to copy the specified SQL file(s) in CR document'
    port = 22
    transport = paramiko.Transport((var_envmt, port))
    transport.connect(username = unix_user_name, password = unix_password)
    sftp = paramiko.SFTPClient.from_transport(transport)

    
    with open(ddl_db_csv) as f:
        reader = csv.DictReader(f) 
        for row in reader: 
            for (k,v) in row.items(): 
                columns[k].append(v)
    print "Files copied to unix server via sftp : "
    print ""

    for i in range(len(columns['Script name'])):
        local_file = (columns['Script name'])[i]
        database_name = (columns['Database'])[i]
        schema_name = (columns['Schema Name'])[i]
        comments = (columns['DB Type'])[i]
        path = local_file
        path1 = path.split("\\")
        file_nm = path1.pop()
        remote_file = '' + file_nm
        remote_path = ''

        if (i > 0):
            prev_db_name = (columns['Database'])[i-1]

        
        if  (local_file != '' and comments == 'Netezza'):
            netezza_cnt += 1
            print local_file + "------>" + remote_file
            logger.info(local_file + "------>" + remote_file)
            try:
                sftp.chdir(remote_path)  # Test if remote_path exists
            except IOError:
                sftp.mkdir(remote_path)  # Create remote_path
                sftp.chdir(remote_path)
            sftp.put(local_file,remote_file)    # At this point, you are in remote_path in either case
            logger.info('***** Calling the Shell script - NZSQL *****')
            netezza_username = schema_name
            #netezza_passwd = getpass.getpass()

            if (i == 0):
                print netezza_username
                netezza_passwd = getpass.getpass()
                prev_netezza_passwd = netezza_passwd                
            elif ((i > 0) and (database_name != prev_db_name)):
                print netezza_username
                netezza_passwd = getpass.getpass()
                prev_netezza_passwd = netezza_passwd                
            else:
                password = prev_netezza_passwd

            
            print ""
            str = 'D:\plink.exe -ssh -pw ' +  unix_password + ' ' + unix_user_name + '@' + var_envmt + ' sh utility_netezza.sh ' + netezza_srvr + ' ' + database_name + ' ' + netezza_username + ' ' + netezza_passwd + ' ' + remote_file + ' ' + cr_number
            print ""
            os.system(str)
            netezza_base_log_name = cr_number + '_NETEZZA_LOAD_LOG' + '.log'
            netezza_log = '' + netezza_base_log_name
            local_netezza_log = 'C:\\Python27\\examples\\NETEZZA_LOGS\\' + netezza_base_log_name
            sftp.get(netezza_log,local_netezza_log)

            
        else:
            logger.info('****** ------ *****')

    if netezza_cnt >= 1:
        netezza_status = '#5DC83C'
    else:
        netezza_status = 'Yellow'

    sftp.close()
    return True 
            
#####################################################################

def unix_login():
    global unix_user_name
    global unix_password


    if env_var == 'DEV':
        unix_srvr = ''
    elif env_var == 'QA':
        unix_srvr = ''
    elif env_var == 'PROD':
        unix_srvr = ''
    else:
        assert False, "Invalid Environment"

    logger.info('****** Connecting to Unix Server Based on Environment *****')
    print ""
    print 'Please Provide Credentials for ' + env_var + ' Unix Server: ' + unix_srvr + ''
    unix_user_name = raw_input("Enter Username:")
    #print unix_user_name
    unix_password = getpass.getpass()

#####################################################################

def Active_Batch():
    print ""
    print "***********************************************************************************************"
    print "*                                                                                             *"
    print "*                                    ACTIVE BATCH MIGRATION                                   *"
    print "*                                                                                             *"
    print "***********************************************************************************************"

    abatch_count = 0
    global abatch_status
    
    table = doc.Tables(10)
    cnt_rows = table.Rows.Count
    cnt_cols = table.Columns.Count


    for row in xrange(2, cnt_rows + 1):
        cr_num = table.Cell(Row=row, Column=1).Range.Text
        envmt = table.Cell(Row=row, Column=2).Range().encode('ascii','ignore').rstrip('\r\x07')
        order = table.Cell(Row=row, Column=3).Range().encode('ascii','ignore').rstrip('\r\x07')
        job_name = table.Cell(Row=row, Column=4).Range().encode('ascii','ignore').rstrip('\r\x07')
        comments = table.Cell(Row=row, Column=5).Range().encode('ascii','ignore').rstrip('\r\x07')
        if (job_name != '' and envmt == env_var):
            abatch_count += 1
            if env_var == 'QA':
                dest_path = '\\\\\\QA\\'
            elif env_var == 'PROD':
                dest_path = '\\\\\\PROD\\'
            else:
                assert False, "Invalid Environment"

            print "Below Files are copied  : "
            shutil.copy2(job_name,dest_path)
            print job_name + "------>" + dest_path


    if abatch_count >= 1:
        abatch_status = '#5DC83C'
    else:
        abatch_status = 'Yellow'

    return True

#####################################################################

def main():
   global sql_status
   global xml_status
   global unix_status
   global netezza_status
   global abatch_status

   ## Validating Inputs
   logger.info('*****Parsing and Validating the Arguments*****')
   try:
       validate(sys.argv[1:])
   except:
       logger.info('*****Error Occured Validating the Input Arguments*****')
       sys.exit(1)
        
   ## Parsing the CR document
   logger.info('*****Parsing CR Document*****')
   try:
       parse_doc()
   except:
       logger.info('*****Error Occured While Parsing the CR document*****')
       sys.exit(1)

   ## Calling sub program for Unix
   logger.info('***** Gathering Unix Server Credentials *****')
   unix_login()
   
   ## Calling the SQL migration script
   logger.info('*****SQL MIGRATION START *****')
   sql_stat = call_sql()
   ## Closing DB Connection
   #logger.info('*****Closing DB Connection*****')
   #if sql_stat:
       #db_disconnect()

   logger.info('*****SQL MIGRATION END *****')
   if not sql_stat:
       logger.info('******SQL MIGRATION FAILURE *****')
       sql_status = 'Red'
       sys.exit(1)


   ## Unix Migration
   logger.info('******UNIX MIGRATION START****')
   unix_stat = unix_migration()
   logger.info('******UNIX MIGRATION END *****')
   if not unix_stat:
       logger.info('******UNIX MIGRATION FAILURE *****')
       unix_status = 'Red'
       sys.exit(1)


   ## Informatica Migration
   logger.info('******INFORMATICA MIGRATION START****')
   logger.info('******Generating XML Dynamically****')
   xml_stat = generate_xml()
   logger.info('******INFORMATICA MIGRATION END *****')
   if not xml_stat:
       logger.info('******INFORMATICA MIGRATION FAILURE *****')
       xml_status = 'Red'
       sys.exit(1)


   ## Netezza Migration
   logger.info('******NETEZZA MIGRATION START****')
   netezza_stat = netezza_migration()
   logger.info('******NETEZZA MIGRATION END *****')
   if not netezza_stat:
       logger.info('******NETEZZA MIGRATION FAILURE *****')
       netezza_status = 'Red'
       sys.exit(1)
       
   ## Actibe Batch Migration
   logger.info('******Active Batch MIGRATION START****')
   ab_stat = Active_Batch()
   logger.info('******Active Batch MIGRATION END *****')
   if not ab_stat:
       logger.info('******Active Batch MIGRATION FAILURE *****')
       abatch_status = 'Red'
       sys.exit(1)
       
   ## Sending Email
   logger.info('******Sending Mail *****')
   send_mail_stat = send_mail_via_com()
   if not send_mail_stat:
       logger.info('******SENDING MAIL FAILURE *****')
       sys.exit(1)

       
   
#####################################################################

# Call the Main Function #
if __name__ == "__main__":
   logger.info('########## STARTING SCRIPT EXECUTION ########## ')
   main()
   logger.info('########## ENDING SCRIPT EXECUTION ########## ')
   


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

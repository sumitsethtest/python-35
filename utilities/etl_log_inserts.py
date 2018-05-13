###########################################################################################################
#                                                                                                         #
# Python Script to Consolidate Log files and upload the resultant data from all files to Google Big Query #
#                                                                                                         #
###########################################################################################################

"""
Import Relevant Python Libraries
"""

import re
import glob
import time
import datetime
from google.cloud import bigquery


timestr = time.strftime("%Y%m%d-%H%M%S")
rows_to_insert=[]
rows_to_insert_err=[]
ml = []
client = bigquery.Client.from_service_account_json('C:\\upw\\gbigquery\\calcium-heaven-203803-381f525f1d39.json')
rg=re.compile('(\d+)\s+(\d{2}[-/]\d{2}[-/]\d{4})\s+(\d{2}[-/:]\d{2}[-/:]\d{2})\s+(\d+)\s+(\d{2}[-/:]\d{2}[-/:]\d{2})\s+(\d)\s+(\d+)\s+(\w*)\s+([+-]\d+)\s+(\d+)\s+(\d+)')


def process_single_log(logfilename):
    """ Function to Process One Log File."""
    global rows_to_insert

    with open(logfilename, 'r') as in_file:
        for _ in range(1):
            next(in_file)
        for line in in_file:
            modifiedline = re.sub(' +', ',', line)
            if re.match(r',', modifiedline):
                modifiedline = modifiedline[1:]
            mod_line_arr = modifiedline.split(",")
            if str(mod_line_arr[5]) == str('1'):
                mod_line_arr[6] = mod_line_arr[6][:1] + ',' + mod_line_arr[6][1:]
            #print(len(mod_line_arr))
            modifiedline = ','.join(str(e) for e in mod_line_arr)
            mod_line_arr = modifiedline.split(",")
            if len(mod_line_arr) == 12:
                #print(len(mod_line_arr))
                #print(mod_line_arr)
                mod_line_arr = mod_line_arr[:-1]
                mod_line_arr[0] = int(mod_line_arr[0])
                mod_line_arr[1] = datetime.datetime.strptime(mod_line_arr[1], '%m/%d/%Y').date()
                mod_line_arr[2] = datetime.datetime.strptime(mod_line_arr[2], '%H:%M:%S').time()
                mod_line_arr[3] = int(mod_line_arr[3])
                mod_line_arr[4] = datetime.datetime.strptime(mod_line_arr[4], '%H:%M:%S').time()
                mod_line_arr[5] = int(mod_line_arr[5])
                mod_line_arr[6] = int(mod_line_arr[6])
                mod_line_arr[9] = int(mod_line_arr[9])
                mod_line_arr[10] = int(mod_line_arr[10])
                rows_to_insert.append(tuple(mod_line_arr))
            else:
                mline = rg.search(line)
                ml.append(int(mline.group(1)))
                ml.append(datetime.datetime.strptime(mline.group(2), '%m/%d/%Y').date())
                ml.append(datetime.datetime.strptime(mline.group(3), '%H:%M:%S').time())
                ml.append(int(mline.group(4)))
                ml.append(datetime.datetime.strptime(mline.group(5), '%H:%M:%S').time())
                ml.append(int(mline.group(6)))
                ml.append(int(mline.group(7)))
                ml.append(mline.group(8))
                ml.append(mline.group(9))
                ml.append(int(mline.group(10)))
                ml.append(int(mline.group(11)))
                rows_to_insert.append(tuple(ml))

def process_all_log_files():
    """ Function to Process Multiple Log Files of a day using a Loop. Calls process_single_log function
        Get all Log files of current day and processes them.
    """
    fptn = 'CDR-' + str(datetime.datetime.now().strftime("%y%m%d")) + '*'
    all_log_files = glob.glob(fptn)
    for logfile in all_log_files:
        print(logfile)
        process_single_log(logfile)


def update_bigqry_table():
    """ Function to Load the Data from all Log Files to Google Big Query"""

    global rows_to_insert
    global client

    dataset_id = "shoreteldataset"
    table_id = "all_log_data"
    dataset_ref = client.dataset('shoreteldataset')
    table_ref = dataset_ref.table('all_log_data')
    table = client.get_table(table_ref)
    errors = client.insert_rows(table, rows_to_insert)

    print("Data Load completed")

    print("Rows Inserted Successfully : " + str(len(rows_to_insert)) + "")
    print("Rows Failed : " + str(len(errors)) + "")

if __name__ == "__main__":
    process_all_log_files()
    update_bigqry_table()
# -*- coding: utf-8 -*-

import csv
import time
import sys
import os

from geopy import geocoders

g_api_key = ''


def main():
    # default input file name is addresses.csv
    # or taking filename from command line
    csv_fname = 'name.csv'
    if len(sys.argv) > 1:
        csv_fname = sys.argv[1]

    g = geocoders.GoogleV3(api_key=g_api_key)

    # Create output.csv if not already exists
    if not os.path.exists('output.csv'):
        with open('output.csv', 'w') as f:
            fieldnames = ['address_ch', 'address_en', 'latitude', 'longitude']
            writer = csv.DictWriter(f, fieldnames = fieldnames)
            writer.writeheader()

    # Read all already processed addresses from output.csv
    with open('output.csv', 'r') as output_file:
        processed_addresses = [row[0] for row in csv.reader(output_file)]

    # Open input csv file 
    with open(csv_fname, 'rb') as input_file:
        input_csv = csv.reader(input_file, delimiter=',')

        with open('output.csv', 'a') as output_file:
            writer = csv.writer(output_file)

            for i, row in enumerate(input_csv, 1):
                #full_addy = row[0].split('   ')[1] # Read address from a row
                full_addy = '   '.join(str(elem) for elem in row)
                
                # Check if address is already processed
                if full_addy in processed_addresses:
                    print('{0}: {1} is already processed'.format(i, full_addy))
                else:
                    # Geocode address
                    r = g.geocode(full_addy, exactly_one=True, timeout=10)
                
                    if r:
                        result_row = (full_addy, r.address.encode('utf-8'), 
                                        r.latitude, r.longitude,)
                        print('{0}: {1} | {2} | ({3}, {4})'.format(i, *result_row))
                        writer.writerow(result_row)
                    else:
                        # If error occured, save address for later
                        if not os.path.exists('errors.txt'):
                            open('errors.txt', 'w').close()
                            
                        with open('errors.txt', 'a') as f:
                            f.write(full_addy + '\n')

                        print('{0}: {1} is not processed. Saved in "errors.txt"'.
                                format(i, full_addy))
                        
                    time.sleep(1)


if __name__ == '__main__':
    main()

# Request Admin Privileges
import ctypes
import requests
import json
import sys
import time
import os
from pptx import Presentation

# Check if Program is running as Admin
if not ctypes.windll.shell32.IsUserAnAdmin():
    print('Please run this program as Administrator.')
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, __file__, None, 1)


ipinfo = requests.get('http://ipinfo.io/json')
ipinfo.raise_for_status()

ipinfo = json.loads(ipinfo.text)

print('City: ' + ipinfo['city'])
print('Country: ' + ipinfo['country'])
print('Location: ' + ipinfo['loc'])
print('Region: ' + ipinfo['region'])
print('Time Zone: ' + ipinfo['timezone'])
print('Postal: ' + ipinfo['postal'])
print('Hostname: ' + ipinfo['hostname'])
print('Operating System:' + ipinfo['org'])
print('IP: ' + ipinfo['ip'])

# Ask the user if they want to save the output to a file
print('Do you want to save the output to a file? (y/n)')
response = input()
if response.lower() == 'y':
    # open file in write mode
    # ask user for the file format they want to save the output as
    print('What file format do you want to save the output as? (txt, json, xml, csv, html, pptx)')
    file_format = input()
    if file_format.lower() == 'txt':
        with open('C:\\ip-info.txt', 'w') as f:
            f.write('City: ' + ipinfo['city'] + '\n')
            f.write('Country: ' + ipinfo['country'] + '\n')
            f.write('Location: ' + ipinfo['loc'] + '\n')
            f.write('Region: ' + ipinfo['region'] + '\n')
            f.write('Time Zone: ' + ipinfo['timezone'] + '\n')
            f.write('Postal: ' + ipinfo['postal'] + '\n')
            f.write('Hostname: ' + ipinfo['hostname'] + '\n')
            f.write('Operating System:' + ipinfo['org'] + '\n')
            f.write('IP: ' + ipinfo['ip'] + '\n')
            f.close()
            print('Output saved to ip-info.txt in C: folder.')
            # Ask the user if he wants to open the file
            print('Do you want to open the file? (y/n)')
            open_file = input()
            if open_file.lower() == 'y':
                os.startfile('C:\\ip-info.txt')

            input('Press Enter to exit...')


    elif file_format.lower() == 'json':
        with open('C:\\ip-info.json', 'w') as f:
            json.dump(ipinfo, f)
            f.close()
            print('Output saved to ip-info.json in C: folder.')
            # Ask the user if he wants to open the file
            print('Do you want to open the file? (y/n)')
            open_file = input()
            if open_file.lower() == 'y':
                os.startfile('C:\\ip-info.json')
            input('Press Enter to exit...')


    elif file_format.lower() == 'xml':
        with open('C:\\ip-info.xml', 'w') as f:
            f.write('<?xml version="1.0" encoding="UTF-8"?>' + '\n')
            f.write('<ipinfo>' + '\n')
            f.write('\t' + '<city>' + ipinfo['city'] + '</city>' + '\n')
            f.write('\t' + '<country>' +
                    ipinfo['country'] + '</country>' + '\n')
            f.write('\t' + '<loc>' + ipinfo['loc'] + '</loc>' + '\n')
            f.write('\t' + '<region>' + ipinfo['region'] + '</region>' + '\n')
            f.write('\t' + '<timezone>' +
                    ipinfo['timezone'] + '</timezone>' + '\n')
            f.write('\t' + '<postal>' + ipinfo['postal'] + '</postal>' + '\n')
            f.write('\t' + '<hostname>' +
                    ipinfo['hostname'] + '</hostname>' + '\n')
            f.write('\t' + '<org>' + ipinfo['org'] + '</org>' + '\n')
            f.write('\t' + '<ip>' + ipinfo['ip'] + '</ip>' + '\n')
            f.write('</ipinfo>')
            f.close()
            print('Output saved to ip-info.xml in C: folder.')
            # Ask the user if he wants to open the file
            print('Do you want to open the file? (y/n)')
            open_file = input()
            if open_file.lower() == 'y':
                os.startfile('C:\\ip-info.xml')
            input('Press Enter to exit...')


    elif file_format.lower() == 'csv':
        with open('C:\\ip-info.csv', 'w') as f:
            f.write('City, Country, Location, Region, Time Zone, Postal, Hostname, Operating System, IP' + '\n')
            f.write(ipinfo['city'] + ',' + ipinfo['country'] + ',' + ipinfo['loc'] + ',' + ipinfo['region'] + ',' + ipinfo['timezone'] + ',' + ipinfo['postal'] + ',' + ipinfo['hostname'] + ',' + ipinfo['org'] + ',' + ipinfo['ip'] + '\n')
            f.close()
            print('Output saved to ip-info.csv in C: folder.')
            # Ask the user if he wants to open the file
            print('Do you want to open the file? (y/n)')
            open_file = input()
            if open_file.lower() == 'y':
                os.startfile('C:\\ip-info.csv')
            input('Press Enter to exit...')




    elif file_format.lower() == 'html':
        with open('C:\\ip-info.html', 'w') as f:
            # Write the HTML code and css on one line
            f.write('<html><head><style>body{font-family:Arial;}</style></head><body><h1>IP Info</h1><p>City: ' + ipinfo['city'] + '</p><p>Country: ' + ipinfo['country'] + '</p><p>Location: ' + ipinfo['loc'] + '</p><p>Region: ' + ipinfo['region'] + '</p><p>Time Zone: ' + ipinfo['timezone'] + '</p><p>Postal: ' + ipinfo['postal'] + '</p><p>Hostname: ' + ipinfo['hostname'] + '</p><p>Operating System: ' + ipinfo['org'] + '</p><p>IP: ' + ipinfo['ip'] + '</p></body></html>')
            f.close()
            print('Output saved to ip-info.html in C: folder.')
            # Ask the user if he wants to open the file
            print('Do you want to open the file? (y/n)')
            open_file = input()
            if open_file.lower() == 'y':
                os.startfile('C:\\ip-info.html')
            input('Press Enter to exit...')

    elif file_format.Lower() == 'pptx':
        prs = Presentation()
        blank_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank_slide_layout)
        left = top = width = height = Inches(1)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        p = tf.add_paragraph()
        p.text = 'IP Info'
        p.font.bold = True
        p.font.size = Pt(20)
        p = tf.add_paragraph()
        p.text = 'City: ' + ipinfo['city']
        p = tf.add_paragraph()
        p.text = 'Country: ' + ipinfo['country']
        p = tf.add_paragraph()
        p.text = 'Location: ' + ipinfo['loc']
        p = tf.add_paragraph()
        p.text = 'Region: ' + ipinfo['region']
        p = tf.add_paragraph()
        p.text = 'Time Zone: ' + ipinfo['timezone']
        p = tf.add_paragraph()
        p.text = 'Postal: ' + ipinfo['postal']
        p = tf.add_paragraph()
        p.text = 'Hostname: ' + ipinfo['hostname']
        p = tf.add_paragraph()
        p.text = 'Operating System: ' + ipinfo['org']
        p = tf.add_paragraph()
        p.text = 'IP: ' + ipinfo['ip']
        # Save the presentation to C: folder
        prs.save('C:\\ip-info.pptx')
        print('Output saved to ip-info.pptx in C: folder.')
        # Ask the user if he wants to open the file
        print('Do you want to open the file? (y/n)')
        open_file = input()
        if open_file.lower() == 'y':
            os.startfile('C:\\ip-info.pptx')
        input('Press Enter to exit...')
            


    else:
        print('Invalid file format.')
        input('Press Enter to exit...')

if response.lower() == 'n':
    print('Output was not saved.')
    input('Press Enter to exit...')

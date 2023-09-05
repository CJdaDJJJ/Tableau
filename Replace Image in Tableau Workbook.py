from enum import Enum
import os
from xml.etree.ElementTree import ElementTree
import zipfile

import pandas as pd
import pantab
import sys
from tableauhyperapi import Connection, CreateMode, HyperProcess, TableName, Telemetry

from pathlib import Path

#get user directory
user_directory = str(Path.home())

#create playground
if not os.path.isdir(user_directory + "\\BDG Dashboards Unpackaged"):
    os.makedirs(user_directory + "\\BDG Dashboards Unpackaged")



class WorkbookType(Enum):
    WORKBOOK = 1
    PACKAGED = 2
    UNKNOWN = 3

WorkbookType = Enum('WorkbookType', ['WORKBOOK', 'PACKAGED', 'UNKNOWN'])

class Workbook:
    type = None
    name = ''
    root_path = ''
    data_path = ''

    def set_type(self, type):
        self.type = type

    def set_name(self, name):
        self.name = name

    def set_root_path(self, root_path):
        self.root_path = root_path
    
    def set_data_path(self, data_path):
        self.data_path = data_path

#find the file most recently modified
import glob

list_of_files = glob.glob('\\\\rb.win.frb.org\\P1\\Shared\\Data Strategy\\BDG Dashboards\\BDG Dashboard Iterations\\*') # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getctime)
print(latest_file) 


example_input_file = latest_file
playground_directory = user_directory + "\\BDG Dashboards Unpackaged\\"

workbook = Workbook()

def get_workbook_type(input_file):
    extension = None
    try:
        extension = os.path.splitext(os.path.basename(input_file))[1]
    except:
        print('File not provided')

    if extension != None:
        if extension == '.twbx':
            return WorkbookType.PACKAGED
        elif extension == '.twb':
            return WorkbookType.WORKBOOK
        else:
            return WorkbookType.UNKNOWN
    else:
        return WorkbookType.UNKNOWN
    
def extract_twbx(input_file):
    basename = os.path.splitext(os.path.basename(input_file))[0]

    output_directory = playground_directory + basename

    workbook.set_root_path(os.path.abspath(output_directory))

    with zipfile.ZipFile(input_file, 'r') as zip_ref:
        zip_ref.extractall(output_directory)

def bootstrap_analysis(input_file):
    # Check if the workbook is TWBX
    if get_workbook_type(input_file=input_file) == WorkbookType.PACKAGED:
        extract_twbx(input_file=input_file)
        return True
    else:
        print('Error getting twbx. Incorrect file provided')
        return False

def analyze_twbx(workbook):
    print('Collecting data about workbook')
    image_path = str(workbook.root_path + '\\Image')
    print(image_path)
    # Does the workbook have images?
    exists = os.path.exists(image_path)
    if exists:
        print('This workbook uses external images, bundled into the workbook. They\'re located in the Images folder')


#Calling functions
print('Unpacking .twbx...')
if bootstrap_analysis(example_input_file):
    print('Successfully unpacked .twbx')
    analyze_twbx(workbook=workbook)
else:
    print('Error unpacking twbx')

#########################################################################################################
#Get MOR screenshots

# 3rd party
import win32com.client

# Environment setup
Application = win32com.client.Dispatch("PowerPoint.Application")
filenames = next(os.walk("\\\\rb.win.frb.org\\P1\\Shared\\Data Strategy\\BDG Dashboards\\MOR Summary Repository"), (None, None, []))[2]

# When given a filename, takes a screenshot of the PPT and exports the JPG
def screenshot_ppt(input):
    base = os.path.splitext(input)[0]
    extension = os.path.splitext(input)[1]
    if extension != '.pptx':
        return
    full_path = os.path.join("\\\\rb.win.frb.org\\P1\\Shared\\Data Strategy\\BDG Dashboards\\MOR Summary Repository\\", input)
    Presentation = Application.Presentations.Open(full_path)
    full_export_path = "\\\\rb.win.frb.org\\P1\\Shared\\Data Strategy\\BDG Dashboards\\MOR Summary Output\\" + base + ".jpg"
    #only prints the first slide
    Presentation.Slides[0].Export(full_export_path, "JPG")
    Presentation = None

# Cycles through all PPTs in Data Strategy folder
def get_ppt():
    for file in filenames:
        try:
            screenshot_ppt(file)
        except:
            print('Error trying to open file ' + file)

#Calling functions
if len(filenames) > 0:
    print('Found ' + ', '.join(filenames))
    get_ppt()
    Application.Quit()
else:
    print('No files found, quitting...')


####################################################################

import shutil

#List of images in the screenshots folder
path = "\\\\rb.win.frb.org\\P1\\Shared\\Data Strategy\\BDG Dashboards\\MOR Summary Output"
dir_list = os.listdir(path)
print(dir_list)

#List of images in the Tableau workbook
workbook_split = os.path.split(latest_file)
workbook_name = os.path.splitext(workbook_split[1])[0]
path2 = user_directory + "\\BDG Dashboards Unpackaged\\" + workbook_name + "\\Image"
dir_list2 = os.listdir(path2)
print(dir_list2)

def transfer_image(path1, path2):
    for filename in os.listdir(path1):
        if filename in os.listdir(path2):
            shutil.copy(path1 + "\\" + filename,path2)
        else:
            if filename != "Thumbs.db": 
                print("Please change the name of " + filename + " to \"MOR + your BDG\"")
                exit()


transfer_image(path,path2)

#####################################################
#Re-package the Tableau workbook


zipped_file = user_directory + "\\BDG Dashboards Unpackaged" + "\\Zipped " + workbook_name

dir_zip = user_directory + "\\BDG Dashboards Unpackaged\\" + workbook_name

zipped = shutil.make_archive(zipped_file, 'zip', dir_zip)

if os.path.exists(zipped):
   print(zipped) 
else: 
   print("ZIP file not created")

#change it into a Tableau packaged workbook

from pathlib import Path

# Specify the path of the file whose extension we want to change
file_path = Path(zipped_file + ".zip")

# Make a new path with the file's extension changed
new_file_path = file_path.with_suffix(".twbx")

# Rename the file's extension to that new path. Catch errors 
try:
    renamed_path = file_path.rename(new_file_path)
except FileNotFoundError:
    print(f"Error: could not find the '{file_path}' file.")
#except FileExistsError:
#    print(f"Error: the '{new_file_path}' target file already exists.")
else:
    print(f"Changed extension of '{file_path.name}' to '{renamed_path.name}'.")


########## upload from local drive to shared drive
shutil.copy(new_file_path, "\\\\rb.win.frb.org\\P1\\Shared\\Data Strategy\\BDG Dashboards\\BDG Dashboard Iterations")

shutil.rmtree(user_directory + "\\BDG Dashboards Unpackaged")

 
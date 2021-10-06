import looker_sdk
from looker_sdk import models
import pandas as pd
import os, zipfile, urllib3, requests

class tabbed_dashboard:
    def __init__(self, dir, zipped_files, output_name):
        self.dir = dir
        self.zipped_files = zipped_files
        self.output_name = output_name
        self.sdk = looker_sdk.init31(config_file='intradoint.ini') 
        
    def clean_folder(self):
        print('inside clean folder')
        if os.path.isdir(self.dir):
            print('Directory Exists')
            for f in os.listdir(self.dir):
                os.remove(os.path.join(self.dir, f))     
        else:
            print('Creating directory')
            os.mkdir(self.dir)   
    
    def unzip_files(self):
        with zipfile.ZipFile(self.zipped_files, 'r') as zip_ref:
           zip_ref.extractall(".")
    
    def write_files(self):
        writer = pd.ExcelWriter(os.path.join(self.dir,self.output_name), engine='xlsxwriter')
        for f in os.listdir(self.dir):
            df = pd.read_csv(os.path.join(self.dir,f))
            df.to_excel(writer, sheet_name=os.path.split(f)[1].split('.')[0][0:30], index=False)
        print('File written to: ', os.path.join(self.dir,self.output_name))
        writer.save()
    
    def main(self):
        self.clean_folder()
        self.unzip_files()
        self.write_files()

tabbed_dashboard('./dashboard-business_pulse', '../Downloads/dashboard-business_pulse.zip', 'tabbed_output.xlsx').main()

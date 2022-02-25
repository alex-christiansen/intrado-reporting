import looker_sdk
from looker_sdk import models
import pandas as pd
import os, zipfile, urllib3, requests

zipped_files_name = '../Downloads/dashboard-business_pulse.zip' # change this to the location of your zipped dashboard
output_file_name = 'tabbed_output.xlsx' # change output file name if you want

class tabbed_excel:
    def __init__(self, dir, zipped_files, output_name):
        self.dir = dir
        self.zipped_files = zipped_files
        self.output_name = output_name
        
    def clean_folder(self):
        if os.path.isdir(self.dir):
            print(self.dir, 'directory exists')
            for f in os.listdir(self.dir):
                os.remove(os.path.join(self.dir, f))     
        else:
            print('Creating directory', self.dir)
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

if __name__ == '__main__':
    db = tabbed_excel('./dashboard-business_pulse', zipped_files_name, output_file_name)
    db.main()
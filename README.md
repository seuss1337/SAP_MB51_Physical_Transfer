# MB51 Transfer Generator

## Intro
This script is used to rapidly create transfer documents required to be physically stored for auditing purposes. It is intended for situations where a large ammount of job records are missing and making them by hand will take a long time. It grabs excel files exported from SAP as the input and exports an Excel workbook with a transfer document per tab.

  *This script is intended to run under google colab so no need to install python and libraries on work computers.

## Setup
- Copy the MB51 folder from JP's shared folder to your google colab folder.

- There are 3 folders Input, Resource and output.


#### Resource Folder

- This folder will have 2 Files MASTER.xlsx and job_names.csv. MASTER is just the transfer sheet template, which you can modify. job_names.csv hold all the job numbers job names and storage location information, it is used to fill out the header on the transfer sheet. It is important that job_names.csv has 3 columns named Job Number, Job Name and Storage Location otherwise the script won't run.

- Open the Resource folder and modify the MASTER.xlsx file with your office number. Also look through the inventory item CM codes and add the ones that your office uses the most and remove the ones that you don't use. These are usually the handgrinding / speedline pads, removal tooling and joint/patch material.

- Open SAP and run Y_DEV_2900016, this is your backlog code. Under output options select Excel Format. Click on box that says "Click for 'Excel', '6 Month Exception' and 'E-Mail' options." Remove everything from Selected fields and add Sales Order/Job #, Job Name, and Storage Location. Press the check mark and press F8 to exceute the report. ( Remove WBS Status not TECO for older jobs). Once Backlog is excecuted press CTRL + SHIFT + F7 to open excel. In the excel box click on File and save the file as a CSV File. Open the CSV file on your computer and check that you have 3 columns. Change the Sales Order (cell A1) to Job number. Save the file and upload it to you Resouce folder.

#### Input Folder

  This folder contains excel file that you add for the script to process. Each excel file is the saved output of mb51. You need 1 file per job number.

- To generate this file go to transaction MB51, enter your Plant Number and the storage location for the job you want to process and press F8. Pres SHIFT F4 and save the file. Upload the file to your Input folder on your google colab and repeat for however many files you want to process. It is a good idea to name the file after the storage location to make keeping tack of multiple files easier.

## Running

  This script is meant to run on google colab to avoid having to install python and libraries on work computers. Follow setup instructions and open untiled.ipynb with google colab. Click on the first cell and press SHIFT + Enter. Once that cell excecutes click on the second cell and press SHIFT + ENTER. Script can take a while to run depending on how many files you're processing. Once the script output says "Script Complete" you can go to your output folder and download your files.


### Printing tips

To print open your output file and press CTRL while selecting all your tabs. Click on file then print and under scaling options select "Fit Sheet on One Page" scroll down the print preview and to make sure it's been applied to all sheets and click print.

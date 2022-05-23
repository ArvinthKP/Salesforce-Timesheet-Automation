# Salesforce-Timesheet-Automation

Automated Timesheet Filling For Sales Force

Format of Excel File Attached

![image](https://user-images.githubusercontent.com/18065155/169826158-b19b1e6d-5f1b-4d32-b8b8-713c7cee3fc5.png)


Need to manually Login. Automated Login Not Working

After all the tasks are added. We need to manually click submit for approval.
this part is not automated because to manually verify the values entered automatically

If your login takes more time please update the values at line 132 
eg time.sleep(100) for avoiding any issues


URL = 'https://****.my.salesforce.com/'
EXCEL_FILE_PATH = "./Tasks.xlsx"

Change the url and the excel file path to your location

Chrome Web Driver is required Please download web driver based on your chrome version.

Chrome Web Driver Download Link
https://chromedriver.chromium.org/downloads

Automated timesheet submission for these fields in the image 
![image](https://user-images.githubusercontent.com/18065155/169825269-f8f815c4-4259-4b60-b8d6-a7c589e518d0.png)



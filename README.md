# DMARC-to-Google-sheet
In Gmail: Save attachments that are DMARC-reports and insert values from them in Google sheet. The reports are in XML-format. 

## Preparations in Google Drive
1. Create a folder and call it DMARC
2. Create a sub-folder of DMARC and call it DMARC-reports, note the folder id
3. Create a sub-folder of DMARC-reports and call it DMARC-archive, note the folder id
4. Create a copy of Google sheet dmarc-report and put it in the DMARC-folder, note the file id. 
https://docs.google.com/spreadsheets/d/10crRusmMwmcTuiUMn0jwFRHSotbXXrNTRTPcx4hJI4U/copy

## Create the script
Go to https://script.google.com/home
Click *New project* 

![bilde](https://user-images.githubusercontent.com/50026860/120083438-f87f0c00-c0c8-11eb-8398-d2543a135e13.png)

Give the project a name (change the Untitled project)
Copy the code from the GAS-file and paste it there.

## Trigger the script
Go to Triggers and Add Trigger

![bilde](https://user-images.githubusercontent.com/50026860/120083982-da1b0f80-c0cc-11eb-87ba-e992c19d5919.png)

Change it settings so the script is trigged for instance once a day. 

![bilde](https://user-images.githubusercontent.com/50026860/120084012-2b2b0380-c0cd-11eb-8589-9271cc9f9d7e.png)


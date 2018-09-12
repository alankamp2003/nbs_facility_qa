Instructions for use:
  1. Start RStudio.
  
  
  2. Install the required packages:
     a. Locate the "Console" pane in RStudio. It's usually in the bottom left corner.
     b. In that pane, at the prompt that looks like ">", type or paste this command - 
        install.packages(c("shiny", "shinyjs", "shinyalert", "readxl", "rmarkdown", "knitr", "kableExtra", "png", dplyr", "mailR", "futile.logger", "properties"))
     c. Press Enter. This will take a few seconds. You should see several messages telling you that a certain package has been successfully unpacked etc. The last message should say something like -
       The downloaded binary packages are in
          C:\...\downloaded_packages.
          Here "..." is a directory path.
  
  
  3. Install tinytex:
     a. In the console pane type or paste - 
        install.packages(c("tinytex", "rmarkdown"))
        tinytex::install_tinytex()
     b. Press Enter. This will take some time. After it's done, you should see messages similar to Step 2.c.
  
  
  4. Set the default directory:
     a. Go to the Tools menu.
     b. Select Global Options.
     c. Set the "Global working directory (when not in a project):" to wherever this code has been unzipped, i.e. the folder that has this file.
  
  
  5. Check and install Java if needed:
     a. Open a command prompt and type "java -version" (no quotes).
     b. If you see an error message saying something like "'java' is not recognized as an internal or external command", then you need to install Java.
     c. If you need to install Java:
        I. Go to http://www.oracle.com/technetwork/pt/java/javase/downloads/jdk8-downloads-2133151.html.
       II. Look for the version for your operating system, e.g. Windows x64, in the table that says "Java SE Development Kit 8u181".
      III. Check "Accept License Agreement" and click the link in front of your operating system's name. 
       IV. This should show a window that asks you whether you want to save a file that ends in ".exe". Click "Yes".
        V. After the file is downloaded, click it to the start installing Java on your computer. 
       VI. You may be asked whether you want to allow this program to make changes to your computer. Click "Allow" or "Yes" as appropriate.
      VII. Follow the instructions from there without making any changes to the default options provided, by clicking "Next" or "Yes" as appropriate.
     VIII. At some point you will be shown where java will be installed e.g. "C:\Program Files\Java". Please make a note of this, as it may be needed later.
       IX. At the end you should see a message saying something like "Java successfully installed". Click "OK" or "End" as appropriate to quit.
  
  
  6. Close and then restart RStudio.
Steps 2 through 6 need to be executed only once.

  7. If "app.R" from this folder isn't open in RStudio:
     a. Go to the File menu.
     b. Select "Open File...".
     c. Choose that file from this folder.
  
  
  8. Run the app:
     a. With "app.R" open, click on "Run App". It should be a button towards the top right corner of the tab that is showing the file.
     b. This should pop up a window that shows the user interface(UI).
  
  
  9. Use the app:
     a. Upload the Excel (.xslx) file containing facility QA data using the "Browse..." button under "Choose facility data file".
     b. Specify the date range for which the data was generated, under "Enter date range".
     c. Click "Generate Reports" to generate the reports from the uploaded file. 
        I. If the facility QA data file has not been uploaded or date range has not been entered, an error message will be shown on the right and reports won't be generated. 
       II. If the two things are present, the QA reports will start getting generated and a progress bar and various messages in the bottom right corner of the window will be shown.
      III. After the reports have been generated, a message will be shown on the right, telling you where they have been saved on your computer.
     d. Upload the Excel (.xslx) file containing facility email addresses, using the "Browse..." button under "Choose facility email file".
     e. Specify the directory where COIIN reports have been saved by entering its full path under "Specify COIIN directory".
     f. Enter your email address and password under "Specify 'from' email address" and "Password" respectively. 
     g. Send emails by clicking "Send emails". To be able to send emails:
        I. Facility QA data file needs to be uploaded.
       II. Date range needs to be entered.
      III. The facility QA reports need to be already present in the directory specified in step 9.c.III.
       IV. The email address and password need to be entered and the email address has to have either "@gmail.com" or "@uiowa.edu".
        V. If any of these conditions is not true, an error message will be shown and the emails will not be sent.
       VI. If the directory for COIIN reports has not been specified, a popup window will be shown asking you whether you would like to send emails without COIIN reports.
        V. If you click "OK" on that window, the emails will be sent. On clicking "Cancel" they won't be sent.
      VII. If the emails are being sent, a progress bar and various messages in the bottom right corner of the window will be shown.
     VIII. After the emails have been sent, a message will be shown that says "Sent emails. See 'xyz' for any warnings or errors.". Here 'xyz' is the name of a log file on your computer, that contains any warnings for missing emails or COIIN reports.
  
  
  7. Stopping the app:
     a. Go back to RStudio.
     b. Locate a hexagonal "STOP" button right above the "Console" pane mentioned in step 2.a and click it. 

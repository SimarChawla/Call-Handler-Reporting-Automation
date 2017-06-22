# Call-Handler-Reporting-Automation
Written for my workplace, this repository contains the files to automate a monthly reporting process that regularly takes 3 days to complete.

The automation is broken into three proccesses.

### 1. Downloading Files
Using java to access the selenium webdriver, the files are searched for and downloaded from the cisco unity connection site.

### 2. Naming and moving the files
Once the files are downloaded, they are appropriately named and moved into their respective folders.

### 3. Extracting data
Using the itextsharp library, the pdf's are parsed and relevant data is extracted.

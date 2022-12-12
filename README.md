### Website 
https://ambiente.messefrankfurt.com/frankfurt/en/exhibitor-search.html

This Scraper imports All exhibitors from the above website into an excel file. 

Fields being scraped:
 - Name of the brand
 - Email
 - Phone Number
 - Country

### Procedure to run the scraper
#### Create and Activate the virtual environment(windows)
``virtualenv venv``

``venv\scripts\venv``

#### Install Requirements
``pip install -r requirements.txt``

#### Run the scraper
``python script.py``


### Convert the script to executable(.exe) file
``pyinstaller --onefile script.py``

This command creates an executable of the script in a dist folder. 

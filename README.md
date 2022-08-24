# mailsender

This script allows you to send the same message to multiple emails at the same time. The list of emails should be provided in the .xlsx format in the *data* GUI field.  
  
To run the script run that line in the directory of your project:  
`python main.py`  
In some cases  
`python3 main.py`  

To allocate the name of a user you are sending an email, put {name} in any place you need in the *body* GUI field and script automatically get the name from the spreadsheet.  

Mappings of column names and variables are provided in the utils.py. Also you will need to put your email and your application password provided by google to run the application.

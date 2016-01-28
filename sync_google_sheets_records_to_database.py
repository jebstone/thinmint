# Retrieves user signups & optouts written to Google Sheets by Squarespace and loads them to a local database
# Logs all batches to a batch table with counts

import gspread
import pyodbc
 
thinmint_db = {
    "Driver": "{SQL Server Native Client 11.0}",
    "Server": "NEPTUNE\SQLEXPRESS",
    "Integrated Security": "True",
    "Trusted_Connection": "yes",
    "Database": "ThinMint"
}
 
datasheets = {
    "Signups": "ThinMintSignups",
    "Resigns": "ThinMintResigns",
    "OptOuts": "ThinMintOptOuts",
    "ClickThrus": "ThinMintOptOuts"
}
 
def log_me(origin, destination, result, record_count=0):
""" Logs the occurrence of a batch update to the database """

    # Set up database connection
    db_connection = pyodbc.connect(**thinmint_db)
    db_cursor = db_connection.cursor()
 
    # Load all new signups to database
    db_cursor.execute("execute thinmint.dbo.LogMe ?,?,?,?", origin, destination, result, record_count)
    db_cursor.commit()
    db_connection.close()
 
 
def move_new_signups():
""" Retrieve new Signups from Google Sheets, load them to a registration database, and log the batch when done """

    # Login with your Google account
    gc = gspread.login('me@gmail.com','password')
 
    spreadsheet = gc.open(datasheets["Signups"])
    worksheet = spreadsheet.sheet1
 
    # Retrieve all signups as a list of lists
    new_signups = worksheet.get_all_values()
 
    # Remove the Header record
    if (new_signups[0][0] == "Submitted On"):
        new_signups.pop(0)
    signup_count = len(new_signups)
 
    for record in new_signups:
        record.append("website")
    
    if (signup_count > 0):
        # Set up database connection
        db_connection = pyodbc.connect(**thinmint_db)
        db_cursor = db_connection.cursor()
 
        # Load all new signups to database
        db_cursor.executemany("execute thinmint.dbo.AddNewUser ?,?,?,?", new_signups)
        db_cursor.commit()
        db_connection.close()
 
    log_me(datasheets["Signups"], 'dbo.Users', 'Success', str(signup_count))
    # Delete users from Google Spreadsheet
    #    print(row)
 
 
def optout_users():
""" Retrieve opt-out users from Google Sheets, load them to an opt-out table, and log the batch when done """
    # Login with your Google account
    gc = gspread.login('me@gmail.com','password')
 
    spreadsheet = gc.open(datasheets["OptOuts"])
    worksheet = spreadsheet.worksheet("OptOuts")
 
    # Retrieve all signups as a list of lists
    optouts = worksheet.get_all_values()
 
    # Trim off all the Google Spreadsheet reporting header records
    optouts = optouts[15:]
 
    optout_count = len(optouts)
    optouts_clean = []
 
    for optout in optouts:
        optouts_clean.append(optout[0:5])
 
    if (optout_count > 0):
        # Set up database connection
        db_connection = pyodbc.connect(**thinmint_db)
        db_cursor = db_connection.cursor()
 
        # Load all new signups to database
        db_cursor.executemany("execute thinmint.dbo.OptOutUser ?,?,?,?,?", optouts_clean)
        db_cursor.commit()
        db_connection.close()
 
    log_me(datasheets["OptOuts"], 'dbo.OptOuts', 'Success', str(optout_count))
 
 
def resign_users():
""" Retrieve resigned users from Google Sheets, load them to an resigns table, and log the batch when done """

    # Login with your Google account
    gc = gspread.login('me@gmail.com','password')
 
    spreadsheet = gc.open(datasheets["Resigns"])
    worksheet = spreadsheet.sheet1
 
    # Retrieve all signups as a list of lists
    resigns = worksheet.get_all_values()
 
    # Remove the Header record
    if (resigns[0][0] == "Submitted On"):
        resigns.pop(0)
    resigns_count = len(resigns)
    
    if (resigns_count > 0):   
 
        for record in resigns:
            resigns[2] = resigns[2][0:min(len(resigns[2]),250)]
        # Set up database connection
        db_connection = pyodbc.connect(**thinmint_db)
        db_cursor = db_connection.cursor()
 
        # Load all new signups to database
        db_cursor.executemany("execute thinmint.dbo.ResignUser ?,?,?", resigns)
        db_cursor.commit()
        db_connection.close()
 
    log_me(datasheets["Resigns"], 'dbo.Resigns', 'Success', str(resigns_count))
 
 
def load_clickthrus():
""" Retrieve users who clicked through the email from Google Sheets, load them to a tracking table, and log the batch when done """
    # Login with your Google account
    gc = gspread.login('me@gmail.com','password')
 
    sheets = []
    total_records = 0
 
    number_of_sheets = 3
    for i in range(1, number_of_sheets+1):
        sheets.append("ClickThrus" + str(i))
 
    for sheet in sheets:
        spreadsheet = gc.open(datasheets["ClickThrus"])
        worksheet = spreadsheet.worksheet(sheet)
 
        # Retrieve all signups as a list of lists
        clicks = worksheet.get_all_values()
 
        # Trim off all the Google Spreadsheet reporting header records
        clicks = clicks[15:]
 
        click_count = len(clicks)
        total_records += click_count
        clicks_clean = []
 
        for click in clicks:
            clicks_clean.append(click[0:5])
 
        if (len(clicks_clean) > 0):
            # Set up database connection
            db_connection = pyodbc.connect(**thinmint_db)
            db_cursor = db_connection.cursor()
 
            # Load all new signups to database
            db_cursor.executemany("execute thinmint.dbo.LoadClickThrus ?,?,?,?,?", clicks_clean)
            db_cursor.commit()
            db_connection.close()
 
    log_me(datasheets["ClickThrus"], 'dbo.ClickThrus', 'Success', str(total_records))
 
move_new_signups()
#optout_users()
#resign_users()
#load_clickthrus()

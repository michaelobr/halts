print("Importing module dependencies...")
import os
import xlwings as xw
import datetime
import time
import pandas as pd
sleeptime = 1

def setup():
    print("Running setup process...")

    # Initialize global scope
    global now
    global today
    global month_dict
    
    # Setup time and date
    now = datetime.datetime.now()

   # This if/else is for proper MM/DD/YYYY formatting to ensure that MM is always 2 digits
    if now.month < 10:
        month = "0" + str(now.month)
    else:
        month = str(now.month)

    if now.day < 10:
        day = "0" + str(now.day)
    else:
        day = str(now.day)

    today = month + "." + day + "." + str(now.year)

    # This is used mostly for filepaths and also setting up dirs
    month_dict = {
        1: "JAN",
        2: "FEB",
        3: "MAR",
        4: "APR",
        5: "MAY",
        6: "JUN",
        7: "JUL",
        8: "AUG",
        9: "SEP",
        10: "OCT",
        11: "NOV",
        12: "DEC"
    }

    # Check to make sure the year dir is setup (for spectras)
    if os.path.isdir("U://***//Daily Spectras//" + str(now.year)) == True:
        pass
    else:
        print("New year detected. Creating a new year directory in [Daily Spectras]...")
        os.mkdir("U://***//Daily Spectras//" + str(now.year))

    # Check to make sure the month dir is setup within the year dir (for spectras)
    if os.path.isdir("U://***//Daily Spectras//" + str(now.year) + "//" + month_dict[now.month]) == True:
        pass
    else:
        print("New month detected. Creating a new month directory in [Daily Spectras]...")
        os.mkdir("U://***//Daily Spectras//" + str(now.year) + "//" + month_dict[now.month])

    # Check to make sure the year dir is setup (for holdings check)
    if os.path.isdir("U://***//Holdings Check//" + str(now.year)) == True:
        pass
    else:
        print("New year detected. Creating a new year directory in [Holdings Check]...")
        os.mkdir("U://***//Holdings Check//" + str(now.year))

    # Check to make sure the month dir is setup within the year dir (for holdings check)
    if os.path.isdir("U://***//Holdings Check//" + str(now.year) + "//" + month_dict[now.month]) == True:
        pass
    else:
        print("New month detected. Creating a new month directory in [Holdings Check]...")
        os.mkdir("U://***//Holdings Check//" + str(now.year) + "//" + month_dict[now.month])

    # Check to make sure the year dir is setup (for daily sheets)
    if os.path.isdir("U://***//Daily Sheet//" + str(now.year)) == True:
        pass
    else:
        print("New year detected. Creating a new year directory in [Daily Sheet]...")
        os.mkdir("U://***//Daily Sheet//" + str(now.year))

    # Check to make sure the month dir is setup within the year dir (for daily sheets)
    if os.path.isdir("U://***//Daily Sheet//" + str(now.year) + "//" + month_dict[now.month]) == True:
        pass
    else:
        print("New month detected. Creating a new month directory in [Daily Sheet]...")
        os.mkdir("U://***//Daily Sheet//" + str(now.year) + "//" + month_dict[now.month])

    # Check to make sure the year dir is setup (for ***)
    if os.path.isdir("U://***//***//" + str(now.year)) == True:
        pass
    else:
        print("New year detected. Creating a new year directory in [***]...")
        os.mkdir("U://***//***//" + str(now.year))

    # Check to make sure the month dir is setup within the year dir (for ***)
    if os.path.isdir("U://***//***//" + str(now.year) + "//" + month_dict[now.month]) == True:
        pass
    else:
        print("New month detected. Creating a new month directory in [***]...")
        os.mkdir("U://***//***//" + str(now.year) + "//" + month_dict[now.month])

def find_last_row(idx, workbook, col=1):
    """ Find the last row in the worksheet that contains data.

    idx: Specifies the worksheet to select. Starts counting from zero.

    workbook: Specifies the workbook

    col: The column in which to look for the last cell containing data.
    """

    ws = workbook.sheets[idx]

    lwr_r_cell = ws.cells.last_cell      # lower right cell
    lwr_row = lwr_r_cell.row             # row of the lower right cell
    lwr_cell = ws.range((lwr_row, col))  # change to your specified column

    if lwr_cell.value is None:
        lwr_cell = lwr_cell.end('up')    # go up untill you hit a non-empty cell

    return lwr_cell.row

def newest(path):
    files = os.lis***ir(path)
    paths = [os.path.join(path, basename) for basename in files]
    tester = max(paths, key=os.path.getctime)
    if os.path.isdir(tester) == True:
        return newest(tester)
    else:    
        return tester

def daily_sheet():
    print("Creating new daily sheet for today...")
    
    # Open the most recent daily sheet 
    daily_sheet = xw.Book(newest("U://***//Daily Sheet//"))
    app = xw.apps.active

    # Grab all the holdings from column F (it's the 6th column of the sheet)
    last_row = find_last_row("Daily Sheet", daily_sheet, 6)
    global column_F
    column_F = xw.Range("F2:F" + str(last_row)).value
    column_F = list(filter(None, column_F))

    # Save the sheet as today's so we can work with it later
    daily_sheet.save(path="U://***//Daily Sheet//" + str(now.year) + "//" + month_dict[now.month] + "//" + today + " Daily Sheet.xlsm")
    app.quit()
    time.sleep(sleeptime)

def ***_spectra(): 
    print("Running *** spectra...")

    # This will open the undated ***/VE spectra
    ***_spectra = xw.Book(r"U:\***\*** Spectra.xlsm")  
    app = xw.apps.active

    # First let's create a sheet object for the param table
    param_sheet = ***_spectra.sheets["Param_table"]

    # To keep the dates currnet we must manually set the formula
    ***_spectra.sheets["Param_table"].activate()
    xw.Range("G5").formula = "=today()"
    xw.Range("G6").formula = "=today()"

    # Remove *** from Param table for now but store its values (we'll put it back after we run ***)
    ***_params = param_sheet.range("A6:K6").value
    param_sheet.range("A6:K6").clear_contents()

    # Clear and run spectra on *** Holdings Tab
    ***_spectra.sheets["*** Holdings"].activate()
    clear_spectra = ***_spectra.macro("Clear")
    clear_spectra()
    ***_macro = ***_spectra.macro("Run_Spectra")
    ***_macro()

    # We need to get rid of the leading 0s for the account numbers
    last_row = find_last_row("*** Holdings", ***_spectra)
    account_numbers = xw.Range("C11:C" + str(last_row)).value
    leading_removed = [s.lstrip("0") for s in account_numbers]
    new_list = [int(x) for x in leading_removed]
    xw.Range("C11:C" + str(last_row)).options(transpose=True).value = new_list

    # Get range.values of *** for *** file
    global ***_data_for_***
    ***_data_for_*** = ***_spectra.sheets["*** Holdings"].range("A11:A" + str(last_row)).value

    # Put *** back in param table, replacing ***
    ***_params = param_sheet.range("A5:K5").value
    param_sheet.range("A5:K5").value = ***_params

    # Clear and run ***
    ***_spectra.sheets["*** Holdings"].activate()
    clear_spectra = ***_spectra.macro("Clear")
    clear_spectra()
    ***_macro = ***_spectra.macro("Run_Spectra")
    ***_macro()

    # We need to get rid of the leading 0s for the account numbers
    last_row = find_last_row("*** Holdings", ***_spectra)
    account_numbers = xw.sheets["*** Holdings"].range("C11:C" + str(last_row)).value
    leading_removed = [s.lstrip("0") for s in account_numbers]
    new_list = [int(x) for x in leading_removed]
    xw.sheets["*** Holdings"].range("C11:C" + str(last_row)).options(transpose=True).value = new_list

    # Get range.values of *** for *** file
    global ***_data_for_***
    ***_data_for_*** = ***_spectra.sheets["*** Holdings"].range("A11:A" + str(last_row)).value

    # Set the param table back to its original state
    param_sheet.range("A5:K5").value = ***_params
    param_sheet.range("A6:K6").value = ***_params

    # Run holdings check macro
    ***_spectra.sheets["*** Holdings"].activate()
    holdings_check = ***_spectra.macro("HoldingsCheck")
    holdings_check()

    # We will then save a dated copy and close the workbook
    ***_spectra.save(path="U://***//Daily Spectras//" + str(now.year) + "//" + month_dict[now.month] + "//" + today + " - *** Spectra.xlsm")
    app.quit()
    time.sleep(sleeptime)

def daily_spectra():
    print("Running daily spectra...")
    
    # This will open the undated daily big main spectra
    daily_spectra = xw.Book(r"U:\***\Daily Spectra.xlsm")  
    app = xw.apps.active

    # Then we will execute the macro 
    run_spectra = daily_spectra.macro("Run_Spectra")
    run_spectra()
    
    # We need to get rid of the leading 0s for the account numbers
    last_row = find_last_row("Spectra", daily_spectra)
    account_numbers = xw.Range("I11:I" + str(last_row)).value
    leading_removed = [s.lstrip("0") for s in account_numbers]
    new_list = [int(x) for x in leading_removed]
    xw.Range("I11:I" + str(last_row)).options(transpose=True).value = new_list

    # Let's go ahead and sat Y to N in the param table for later  
    sheet_paramtable = daily_spectra.sheets["Param_table"]
    sheet_paramtable.range("M2").value = "N"

    # Gather the holdings to later be added to Holdings Backup.xlsx (non-dated)
    global all_holdings_sans_***
    all_holdings_sans_*** = xw.Range("A11:L" + str(last_row)).value

    # Remove duplicates
    remove_duplicates = daily_spectra.macro("RemoveDuplicates")
    remove_duplicates()

    # Convet column A to all text format
    conversion = daily_spectra.macro("ConvertToText")
    conversion()

    # Get data for *** file
    global main_data_for_***
    main_data_for_*** = xw.Range("A11:A" + str(last_row)).value

    # We will then save a dated copy and close the workbook
    daily_spectra.save(path="U://***//Daily Spectras//" + str(now.year) + "//" + month_dict[now.month] + "//" + today + " - Daily Spectra with ***.xlsm")
    app.quit()
    time.sleep(sleeptime)

def ***_spectra():
    # This assumes the *** spectra is today's spectra
    print("Fetching data from *** spectra...")
    
    # Check to ensure it's today's spectra
    while True:
        time_of_file = os.path.getmtime(r"U:\***\*** Halts Multi-Spectra.xlsm")
        formatted = datetime.datetime.utcfromtimestamp(time_of_file).strftime("%m.%d.%Y")
        if formatted == today:
            print("*** Spectra up to date. Pulling data now...")
            break
        else: 
            print("*** Spectra out of date. Please drag to SDM halts folder. Refreshing in 60 seconds...")
            time.sleep(60)

    # Open the spectra
    ***_spectra = xw.Book(r"U:\***\*** Halts Multi-Spectra.xlsm")
    app = xw.apps.active

    # Convert account numbers
    last_row = find_last_row("Security Distribution", ***_spectra)
    account_numbers = xw.Range("I11:I" + str(last_row)).value
    leading_removed = [s.lstrip("0") for s in account_numbers]
    new_list = [int(x) for x in leading_removed]
    xw.Range("I11:I" + str(last_row)).options(transpose=True).value = new_list

    # Select text
    global ***_holdings
    ***_holdings = xw.Range("A11:L" + str(last_row)).value

    # Get data for *** file
    global ***_data_for_***
    ***_data_for_*** = xw.Range("A11:A" + str(last_row)).value

    # Save and exit
    ***_spectra.save()
    app.quit()
    time.sleep(sleeptime)

def add_***_to_big():
    print("Pasting *** holdings over to main spectra...")
    
    # Open the dated copy that we already ran and saved with today's holdings
    daily_spectra = xw.Book("U://***//Daily Spectras//" + str(now.year) + "//" + month_dict[now.month] + "//" + today + " - Daily Spectra with ***.xlsm")
    app = xw.apps.active

    # Paste *** holdings below
    last_row = find_last_row("Spectra", daily_spectra)
    xw.Range("A" + str(last_row + 1)).value = ***_holdings

    # Save and close
    daily_spectra.save()
    app.quit()
    time.sleep(sleeptime)

def holdings_check():
    print("Running holdings check...")
    
    # Open the undated Holdings Backup.xlsm file
    holdings_backup = xw.Book(r"U:\***\Holdings Backup.xlsm")
    app = xw.apps.active
    
    # Paste over the all holdings (this is everything that was run from the daily spectra)
    xw.Range("A2").value = all_holdings_sans_***

    # Paste over *** holdings below daily spectra holdings
    last_row = find_last_row("ALL Holdings", holdings_backup)
    xw.Range("A" + str(last_row + 1)).value = ***_holdings

    # Do the data string in column M
    last_row = find_last_row("ALL Holdings", holdings_backup) # we need to find the last row again because we just pasted over *** holdings
    formula = xw.Range("M1").formula
    xw.Range("M1:M" + str(last_row)).formula = formula

    # Paste over column F (sans blanks) from the daily sheet to column N
    xw.Range("N2").options(transpose=True).value = column_F

    # Apply formula in column O and drag down just until the last row of data in column N
    last_row = find_last_row("ALL Holdings", holdings_backup, 14) 
    formula = xw.Range("O2").formula
    xw.Range("O2:O" + str(last_row)).formula = formula

    # Apply N/A filter to column O (the securities that remain are the securities that are no longer held)
    filterNA = holdings_backup.macro("FilterNA")
    filterNA()

    # Save a dated copy to correct directory
    holdings_backup.save(path="U://***//Holdings Check//" + str(now.year) + "//" + month_dict[now.month] + "//" + today + " - Holdings Check.xlsm")
    app.quit()
    time.sleep(sleeptime)

def ***():
    # Open the *** file
    ***_file = xw.Book(r"U:\***\*** Data.xlsm")
    app = xw.apps.active

    # Combine all holdings into one master list
    total_holdings = main_data_for_*** + ***_data_for_*** + ***_data_for_*** + ***_data_for_***
    total_holdings = list(filter(None, total_holdings))

    # Clear the sheet
    ***_file.sheets["All Holdings"].clear_contents()

    # Make column A text format so data is pasted over correctly
    convert = ***_file.macro("ConvertToText")
    convert()

    # Paste over total_holdings
    xw.Range("A2").options(transpose=True).value = total_holdings

    # Remove duplicates
    remove_duplicates = ***_file.macro("RemoveDuplicates")
    remove_duplicates()

    # Sort the file to add *** tags
    tag = ***_file.macro("Sort")
    tag()

    # Save and close
    ***_file.save(path="U://***//***//" + str(now.year) + "//" + month_dict[now.month] + "//" + today + " - *** Upload.xlsm")
    app.quit()
    time.sleep(sleeptime)

def decompose():
    print("Running end of day decomposition...")
    
    # Open the dated *** spectra
    ***_spectra = xw.Book("U://***//Daily Spectras//" + str(now.year) + "//" + month_dict[now.month] + "//" + today + " - *** Spectra.xlsm")  
    app = xw.apps.active
    
    # Clear *** holdings
    ***_spectra.sheets["*** Holdings"].activate()
    clear_holdings = ***_spectra.macro("Clear")
    clear_holdings()

    ## Clear and set active sheet to *** holdings 
    ***_spectra.sheets["*** Holdings"].activate()
    clear_holdings = ***_spectra.macro("Clear")
    clear_holdings()

    # Save it as the nondated *** spectra overwriting it
    ***_spectra.save(path=r"U:\***\*** Spectra.xlsm") 
    app.quit()
    time.sleep(sleeptime)

if __name__ == "__main__":
    setup()
    daily_sheet()
    ***_spectra()
    daily_spectra()
    ***_spectra()
    add_***_to_big()
    holdings_check()
    ***()
    #decompose()
    print("Finished.")
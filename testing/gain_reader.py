import pandas as pd
import os
import datetime
import calendar

def newest(path):
    files = os.listdir(path)
    paths = [os.path.join(path, basename) for basename in files]
    tester = max(paths, key=os.path.getctime)
    if os.path.isdir(tester) == True:
        return newest(tester)
    else:    
        return tester

def setup():
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

    # Monday = 0, Friday = 4, if it's less than 5 than means it's a workday and we should run the program
    weekday_count = calendar.weekday(now.year, now.month, now.day)

    if weekday_count <= 4:
        # Check to make sure the year dir is setup 
        if os.path.isdir("U://***//Daily Folder//" + str(now.year)) == True:
            pass
        else:
            print("New year detected. Creating a new year directory in [Daily Folder]...")
            os.mkdir("U://***//Daily Folder//" + str(now.year))

        # Check to make sure the month dir is setup within the year dir
        if os.path.isdir("U://***//Daily Folder//" + str(now.year) + "//" + month_dict[now.month]) == True:
            pass
        else:
            print("New month detected. Creating a new month directory in [Daily Folder]...")
            os.mkdir("U://***//Daily Folder//" + str(now.year) + "//" + month_dict[now.month])

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
        
        return True
    else:
        print("Setup failure. Not a business day.")
        return False
setup()

def get_different_rows(source_df, new_df):
    """Returns just the rows from the new dataframe that differ from the source dataframe"""
    merged_df = source_df.merge(new_df, indicator=True, how='outer')
    changed_rows_df = merged_df[merged_df['_merge'] == 'right_only']
    return changed_rows_df.drop('_merge', axis=1)

morning_download = pd.read_csv("U://***//GDP//" + str(now.year) + "//" + month_dict[now.month] + "//" + today + " - GDP Download AM.csv")
afternoon_download = pd.read_csv("U://***//GDP//" + str(now.year) + "//" + month_dict[now.month] + "//" + today + " - GDP Download PM.csv")

# Remove unneccesary columns we don't need
morning_download = morning_download.drop(["BBG_RETURN_CODE", "BBG_NUMBER_OF_FIELDS", "ID_BB_GLOBAL"], axis=1)
afternoon_download = afternoon_download.drop(["BBG_RETURN_CODE", "BBG_NUMBER_OF_FIELDS", "ID_BB_GLOBAL"], axis=1)

# Filter the data frame for only the halt codes we need
halt_codes = ["ACQU", "AHLT", "DLST", "HALT", "SUSP", "UNLS"]
morning_download = morning_download.query("MARKET_STATUS in @halt_codes")
afternoon_download = afternoon_download.query("MARKET_STATUS in @halt_codes")

# This will find out which halt codes have changed from the morning 
with pd.option_context('display.max_rows', None, 'display.max_columns', None):
    changed_tags = get_different_rows(morning_download, afternoon_download)
    print(changed_tags)

# Load in the daily sheet 
dsdf = pd.read_csv(r"U:\***\prod\dailysheet.csv")

# Get rid of columns we don't need 
dsdf = dsdf.drop(["SECURITY DESCRIPTION", "USERBANK & CLIENT", "REASON"], axis=1)

# Print results
with open("dsdf.csv", "w") as f:
    dsdf.to_csv(f, index=False)
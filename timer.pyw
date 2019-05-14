import schedule
import time
import halts
refresh_time = 20

def job():
    halts.setup()
    halts.daily_sheet()
    halts.***_spectra()
    halts.daily_spectra()
    halts.***_spectra()
    halts.add_***_to_big()
    halts.holdings_check()
    halts.***()
    #halts.decompose()
    return

schedule.every().monday.at("09:30").do(job)
schedule.every().tuesday.at("09:30").do(job)
schedule.every().wednesday.at("09:30").do(job)
schedule.every().thursday.at("09:30").do(job)
schedule.every().friday.at("09:30").do(job)

while True:
    schedule.run_pending()
    time.sleep(refresh_time)
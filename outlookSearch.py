import pandas as pd
import win32com.client
import os
import time
from datetime import datetime, timedelta

def outlookSearch():
    selections = [1, 2, 3]
    choice = 6
    print('Please select a mail box to search from the options below:')
    print('1. Inbox')
    print('2. Sent')
    print('3. Deleted')
    while True:
        try:
                selection = int(input('\nPlease make a selection: '))
                if selection in selections:
                    if selection == 1:
                        print('Searching inbox...')
                    if selection == 2:
                        print('Searching sent emails..')
                        choice = 5
                    if selection == 3:
                        print('Searching deleted emails...')
                        choice = 3
                    break
                else:
                    print('Command not recognized. Please enter a number from the options above.')
        except ValueError:
            print('Command not recognized. Please enter a number from the options above.')


    dateFrame = pd.read_excel('input.xlsx', engine='openpyxl')
    dateFrame['Email Match'] = pd.NA
    outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
    emailBox = outlook.GetDefaultFolder(choice)
    messages = emailBox.Items
    cutoffDate = (datetime.now() - timedelta(days=45)).strftime('%m/%d/%Y %H:%M %p')
    filter = f"[SentOn] >= '{cutoffDate}'"
    messages = emailBox.Items.Restrict(filter)
    messages.Sort("[SentOn]", True)

    print(f"Searching emails sent after {cutoffDate}...")
    startTime = time.time()
    for idx, searchId in enumerate(dateFrame['SearchID']):
        searchIdString = str(searchId)
        match = False
        for message in messages:
            try:
                if searchIdString in message.Subject or searchIdString in message.Body:
                    match = True
                    break
            except AttributeError:
                continue

        dateFrame.at[idx, 'Email Match'] = 'Match' if match else pd.NA

    dateFrame.to_excel('output.xlsx', index=False)
    endTime = time.time()
    executionTime = endTime - startTime
    print(f"Done! Results saved to {os.path.abspath('output.xlsx')}, taking {executionTime:.2f} seconds")

def main():
    outlookSearch()

if __name__ == "__main__":
    main()

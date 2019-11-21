from O365 import Account, FileSystemTokenBackend, MSGraphProtocol
import time
import datetime
import re

CLIENT_ID = '85b6aa4d-103c-48f4-b0da-d766f18c7518'
SECRET_ID = 'edBgd-L0FAMELEYuaZ:t8=S8F9DoJhK]'

credentials = (CLIENT_ID, SECRET_ID)
protocol = MSGraphProtocol(default_resource='USCHIconfroom1@nemera.net') 

token_backend = FileSystemTokenBackend(token_path='auth_token', token_filename='auth_token.txt')
scopes = ['Calendars.Read.Shared','offline_access']
account = Account(credentials, protocol=protocol, token_backend=token_backend)
account.authenticate(scopes=scopes)

schedule = account.schedule()

calendar = schedule.get_default_calendar()

start_time = float(time.time())
end_time = float(time.time())

flag = 0

def parse_event_string(event):
    event_string = str(event)

    start_index = event_string.find('from:') + 6
    end_index = event_string.find('to:') - 1 
    start_meeting_time = event_string[start_index:end_index]

    start_obj = datetime.datetime.strptime(start_meeting_time, '%H:%M:%S')
    now = datetime.datetime.now().strftime('%H:%M:%S')
    now = datetime.datetime.strptime(now, '%H:%M:%S')

    time_diff_min = ((start_obj - now).total_seconds())/60

    return time_diff_min

while(1):
    
    end_time = float(time.time())

    if (end_time - start_time) > 60:


        start_time = float(time.time())
        print("it's been one minute, I'm refreshing the auth token")

        #account.connection.refresh_token()

        today = datetime.date.today()
        tomorrow = today + datetime.timedelta(days=1)

        q = calendar.new_query('start').greater_equal(today)
        q.chain('and').on_attribute('end').less_equal(tomorrow)

        new_event = calendar.get_events(query=q, include_recurring=True)

        for event in new_event:

            time_diff_min = parse_event_string(event)
        
            if (time_diff_min < 10) and (time_diff_min > 0):
                print(event)





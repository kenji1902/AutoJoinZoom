import pandas as pd
import os
from datetime import datetime as date
from datetime import time
import webbrowser as wb
import sched
import time
from win10toast import ToastNotifier

pd.set_option('display.max_colwidth', None)

class Zoom:
    def __init__(self,schedulePath,sheet):
        self.schedulePath = schedulePath
        self.sheet = sheet
        self.weekDay = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']

    def joinMeeting(self):
        Meetings = self.loadSchedule()
        print("\n\nMeetings for Today: ")
        print(Meetings[['Course','Time_In','Time_Out']])
        Course,link,TimeIn,timeOut = self.getCurrMeeting(Meetings)

        if link is not None:
            print(f"\nNext Iteration: {timeOut}")
            wb.open(link)
            notify = ToastNotifier()
            notify.show_toast(Course, f"from {TimeIn} to {timeOut}\nlink: {link}", duration=25, icon_path = "Assets/images/panic.ico")
            return timeOut
        return 5.0

    def loadSchedule(self):
        dataFrame = pd.read_excel(self.schedulePath + self.sheet)
        currWeekDay = date.today().weekday()
        return dataFrame.loc[dataFrame['Day'] == self.weekDay[currWeekDay]]

    def getCurrMeeting(self, Meetings):
        now = date.now().time()
        now = now.strftime('%H:%M:%S')
        now = date.strptime(now,'%H:%M:%S').time()
        Meetings = Meetings.loc[Meetings['Time_In'] <= now ]
        Meetings = Meetings.loc[Meetings['Time_Out'] > now ]
        print("\n\n\nCurrent Meeting: ")
        print(Meetings)
        if not Meetings.empty:
            return (Meetings['Course'].head(1).to_string(index=False),
                    Meetings['Zoom_Link'].head(1).to_string(index=False),
                    date.strptime(Meetings['Time_In'].head(1).to_string(index=False),'%H:%M:%S').time(), 
                    date.strptime(Meetings['Time_Out'].head(1).to_string(index=False),'%H:%M:%S').time())
        return None, None, None, None

def getTimeNow():
    now = date.now().time()
    now = now.strftime('%H:%M:%S')
    return date.strptime(now,'%H:%M:%S').time()

def subtractTime(time1, time2):
    dateObj = date(1, 1, 1)
    time1 = date.combine(dateObj, time1)
    time2 = date.combine(dateObj, time2)
    return time1 - time2
def clear():
   # for windows
   if os.name == 'nt':
       _ = os.system('cls')

   # for mac and linux
   else:
       _ = os.system('clear')

def run(function):
    while True:
        clear()
        nextTime = function()
        if not isinstance(nextTime, float):
            nextTime = subtractTime(nextTime,getTimeNow())
            nextTime = nextTime.total_seconds()
        time.sleep(nextTime)


if __name__ == '__main__':
    sheet = 'schedule.xlsx'
    schedulePath = f'Assets/excel/'

    zoom = Zoom(schedulePath,sheet)
    schedule = sched.scheduler(time.time,time.sleep)

    if not os.path.exists(schedulePath):
        exit()
    
    run(zoom.joinMeeting)
        
        
import time
from datetime import datetime
import win32com.client as wincl
from plyer import notification
import psutil

class Batterycheaker:

    #this function is convert battery reaming time into hour and minutes
    def convertTime(self, seconds):
        minutes, seconds = divmod(seconds, 60)
        hours, minutes = divmod(minutes, 60)
        return "%d hour %02d minutes %02d" % (hours, minutes, seconds)

    #this function will return the actuall time period AM or PM and help to say correct welcome message
    def currentTimeperiod(self):
        timePeriod = datetime.today().strftime("%I:%M %p")
        if "AM" in timePeriod:
            return "AM"
        else:
            return "PM"

    #function is to speak
    def speak(self, str):
        speaker_number = 1
        spk = wincl.Dispatch("SAPI.SpVoice")
        vcs = spk.GetVoices()
        spk.Voice
        spk.Rate = -1
        spk.SetVoice(vcs.Item(speaker_number))
        spk.Speak(str)

    #function to battery Notification
    def Battery_Notification(self):
        battery = psutil.sensors_battery()
        self.battery_percent = battery.percent;
        battery_plugged = battery.power_plugged
        message = str(self.battery_percent) + "% Battery Remaining \n"+"power plugged is "+ str(battery_plugged)+"\n"+str(self.convertTime(battery.secsleft))
        notification.notify(
            app_name="Agent Diana",
            title="Battery Percentage",
            message=message,
            timeout=10
        )

    #function to detect your reamainig battery and tell you
    def batteryNotify(self):
        self.Battery_Notification()
        if(int(self.battery_percent) >= 90):
            self.speak("sir your battery is fully charged plugged out your charger")
        else:
            if (int(self.battery_percent) <= 25):
                self.speak("sir your battery is low plugged in your charger")
            else:
                self.speak("sorry for disturbing you sir")


#and here is main function where yout program execution startup
if __name__ == "__main__":
    batteryCheaker_Obj = Batterycheaker()
    if batteryCheaker_Obj.currentTimeperiod() == "AM":
        batteryCheaker_Obj.speak("welcome sir Diana here the is "+datetime.today().strftime("%I:%M %p")+" now that's why i'm saying good night")
    else:
        batteryCheaker_Obj.speak("welcome sir Diana here the is " + datetime.today().strftime("%I:%M %p") + " now and good morning")
    while(True):
        batteryCheaker_Obj.batteryNotify()
        time.sleep(60*15)

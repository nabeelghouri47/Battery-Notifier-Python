import time
from datetime import datetime
import win32com.client as wincl
from plyer import notification
import psutil


class Batterycheaker:
    # this function is convert battery reaming time into hour and minutes
    def convertTime(self, seconds):
        minutes, seconds = divmod(seconds, 60)
        hours, minutes = divmod(minutes, 60)
        return "%d hour %02d minutes %02d second" % (hours, minutes, seconds)

    # this function will return the actuall time period AM or PM and help to say correct welcome message
    def partofDay(self, x):
        if (x > 4) and (x <= 8):
            return 'Early Morning'
        elif (x > 8) and (x <= 12):
            return 'Morning'
        elif (x > 12) and (x <= 16):
            return 'After Noon'
        elif (x > 16) and (x <= 20):
            return 'Good Evening'
        elif (x > 20) and (x <= 24):
            return 'Good Night'
        else:
            return 'Late Night'
    # function is to speak
    def speak(self, str):
        speaker_number = 1
        spk = wincl.Dispatch("SAPI.SpVoice")
        vcs = spk.GetVoices()
        spk.Voice
        spk.Rate = -1
        spk.SetVoice(vcs.Item(speaker_number))
        spk.Speak(str)

    # function to battery Notification
    def Battery_Notification(self):
        self.battery = psutil.sensors_battery()
        self.battery_percent = self.battery.percent;
        self.battery_plugged = self.battery.power_plugged
        message = str(self.battery_percent) + "% Battery Remaining \n" + "power plugged is " + str(
            self.battery_plugged) + "\n" + str(self.convertTime(self.battery.secsleft))
        notification.notify(
            app_name="Agent Diana",
            title="Battery Percentage",
            message=message,
            timeout=10
        )

    # function to detect your reamainig battery and tell you
    def batteryNotify(self):
        self.Battery_Notification()
        if (int(self.battery_percent) >= 90):
            self.speak("sir your battery "+str(self.battery_percent)+" is approxtimately full charged so plugged out your charger")
        else:
            if (int(self.battery_percent) <= 20):
                self.speak("sir your battery is low plugged in your charger")
            else:
                self.speak(
                    "your remaining battery is " + str(self.battery_percent) + " percent which you can use for " + str(
                        self.convertTime(self.battery.secsleft)) + " left")

    def checkBatteryplugged(self):
        bateryPlugged = False
        if(self.battery_plugged == True):
            bateryPlugged = True
        return bateryPlugged
# and here is main function where yout program execution startup
if __name__ == "__main__":
    batteryCheaker_Obj = Batterycheaker()
    currentTime = datetime.now()
    hour = currentTime.hour
    batteryCheaker_Obj.speak(batteryCheaker_Obj.partofDay(hour) + " Code 47 diana here time is: "+datetime.today().strftime("%I:%M:%p"))

    while (True):
        batteryCheaker_Obj.batteryNotify()
        time.sleep(60 * 15)

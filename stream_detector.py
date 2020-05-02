import webbrowser
import wmi
from twitchAPI.twitch import Twitch
from time import sleep
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume

target_stream = 'bobross'
offline_sleep_time = 60*5  # Rest time in seconds if RL not live
online_sleep_time = 60*20  # Increase time between live-checks if RL is live
sleep_time = offline_sleep_time  # Start in offline state
twitch = Twitch('wjzimsnic43dkim6vqxyu2345wjfyp',
                'ei8qdwqk48ofnngec517qi66wmnha9')

while True:
    print('Checking if live')
    api_result = twitch.get_streams(user_login=target_stream)
    is_live = len(api_result['data']) > 0
    if is_live:
        print("Channel is live!")
        browser_open = len(wmi.WMI().Win32_Process(name='chrome.exe')) > 0
        if not browser_open:
            webbrowser.open('https://twitch.tv/' + target_stream)
            muted = False
            while not muted:
                sessions = AudioUtilities.GetAllSessions()
                for session in sessions:
                    if session.Process and session.Process.name() == "chrome.exe":
                        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                        volume.SetMasterVolume(0.01, None)
                        muted = True
        sleep_time = online_sleep_time
    else:
        print("Not live")
        sleep_time = offline_sleep_time
    sleep(sleep_time)

for title in pygetwindow.getAllTitles():
    print(title)
# import wikipedia as w
# import pyaudio as pa
import speech_recognition as sr
import win32com.client
import webbrowser as w
import os
import datetime

# import openai
# from config import apikey

speaker = win32com.client.Dispatch("SAPI.spVoice")
shell = win32com.client.Dispatch("WScript.Shell")
wmp = win32com.client.Dispatch("WMPlayer.OCX")

browsing = [["youtube", "https://www.youtube.com"],
            ["google", "https://www.google.com/"],
            ["instagram", "https://www.instagram.com"],
            ["youtube music", "https://music.youtube.com"],
            ["wikipedia", "https://www.wikipedia.com"],
            ["whatsapp", "https://www.whatsapp.com"]
            ]
sys_apps = [["notepad", "notepad.exe"],
            ["calculator", "calc.exe"],
            ["chrome", "chrome.exe"],
            ["microsoft edge", "msedge.exe"],
            ]
apps = [["camera", f"start microsoft.windows.camera:"],
        ["zoom", "%APPDATA%\\Zoom\\bin\\Zoom.exe"]]

songs = [["amplifier", "F:\\DATA\\Music\\Bajan\\Music\\new songs\\01 Amplifier (Imran Khan) n.j.mp3"],
         ["blue eyes", "F:\\DATA\\Music\\Bajan\\Music\\new songs\\_Songs-24 - Copy.mp3"],
         ["baarish",
          "F:\\DATA\\Music\\Bajan\\Music\\new songs\\01 - Baarish - Half Girlfriend  DJMaza.Life  - Copy.mp3"],
         ["love dose", "F:\\DATA\\Music\\Bajan\\Music\\new songs\\02 Love Dose - Yo Yo Honey Singh.mp3"],
         ["woofer", "F:\\DATA\\Music\\Bajan\\Music\\new songs\\02 Woofer (Clean) Dr Zeus n Zora 320Kbps.mp3"],
         ["long drive", "F:\\DATA\\Music\\Bajan\\Music\\new songs\\03 Long Drive (Khiladi 786)(1).mp3"],
         ["thodi der", "F:\\DATA\\Music\\Bajan\\Music\\new songs\\02 - Thodi Der - Half Girlfriend  DJMaza.Life .mp3"],
         ["proper patola", "F:\\DATA\\Music\\Bajan\\Music\\new songs\\01 - Proper Patola (feat. Badshah) - Dow.mp3"],
         ["dance hall", "C:\\Users\\HP\\Downloads\\whip-afro-dancehall-music-110235.mp3"],
         ]


# todo: introduce open ai
# def ai(prompt):
#     openai.api_key = apikey
#     response = openai.ChatCompletion.create(
#         model="gpt-4",  # or "gpt-3.5-turbo"
#         messages=[{"role": "user", "content": prompt}],
#         max_tokens=150,
#         temperature=0.7
#     )
#     return response.choices[0].message['content']

def take_command():
    t = sr.Recognizer()
    with sr.Microphone() as source:
        # t.pause_threshold = 1
        audio = t.listen(source)
        try:
            print("Recognizing..")
            query = t.recognize_google(audio, language="en-in")
            print(f"User Said..:{query}")
            return query
        except Exception as e:
            return "some error occurred"


# todo: to add chat feature through ai
# def chat(query):


if __name__ == '__main__':
    print('hello i am jarvis')
    # speaker.Speak(" hello i am jarvis AI")

    time = datetime.datetime.now().strftime("%H")
    print(time)
    if time in range(4, 12):
        speaker.Speak(" hello,Good Morning, i am jarvis AI")
    elif time in range(12, 16):
        speaker.Speak(" hello,Good afternoon, i am jarvis AI")
    elif time in range(16, 20):
        speaker.Speak(" hello,Good evening, i am jarvis AI")
    elif time in range(20, 24):
        speaker.Speak(" hello,Good night, i am jarvis AI")
    # print("initializing..")
    # query = take_command()
    # speaker.Speak(query)

    # for female voice //LUNA
    # voices = speaker.GetVoices()
    # if "hello luna".lower() in query.lower():
    #     for i in range(voices.Count):
    #         voice = voices.Item(i)
    #         if voice.GetAttribute("Gender") == "Female":
    #             speaker.Voice = voice
    #             print(f"Selected voice: {voice.GetAttribute('Name')}")
    #             break
    #     con = 1
    con = 1
    while con == 1:
        speaker.Speak(" how may i help you")
        print("listening...")
        query = take_command()
        print(query)
        speaker.Speak(query)
        # todo: add more features
        # to browse any site
        for site in browsing:
            # print(site[0])
            if f"open {site[0]}".lower() in query.lower():
                speaker.Speak(f"opening {site[0]} mam....")
                w.open(site[1])
        # todo: add more songs
        # to play songs from desktop
        for song in songs:
            if f"play {song[0]}".lower() in query.lower():
                if os.path.exists(song[1]):
                    # F:\DATA\Music\Bajan\Music\new
                    # songs
                    music_file = song[1]
                    speaker.Speak(f"playing {song[0]} mam....")
                    os.system(f'start "" "{music_file}"')
                else:
                    print("File not found!")
        # to open sys_apps
        for app in sys_apps:
            if f"open {app[0]}".lower() in query.lower():
                speaker.Speak(f"opening {app[0]} mam....")
                command = app[1]
                shell.Run(command)
        # to open apps
        for app in apps:
            if f"open {app[0]}".lower() in query.lower():
                speaker.Speak(f"opening {app[0]} mam....")
                os.system(app[1])
        # to show current time
        if "what is the time now".lower() in query.lower():
            time = datetime.datetime.now().strftime("%H:%M:%S")
            speaker.Speak(f"mam the time is {time} ")
            # hr = datetime.datetime.now().strftime("%H")
            # min = datetime.datetime.now().strftime("%M")
            # sec = datetime.datetime.now().strftime("%S")
            # speaker.Speak(f"mam the time is {hr} baj ke{min} minutes aur{sec}seconds ")

        # user_input = "Write a post caption for a birthday."
        # ai_response = get_openai_response(user_input)
        # print("AI Response:", ai_response)

        # for ai
        # if "openai".lower() in query.lower():
        #     ai(prompt=query)

        # todo : to terminate jarvis's listening
        if "that's all jarvis".lower() in query.lower():
            speaker.Speak("okay, let me know if i can help you with anything else. ")
            con = 0
    # else:
    #     print("unable to recognize")

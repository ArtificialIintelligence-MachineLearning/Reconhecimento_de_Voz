import win32com.client

voice = win32com.client.Dispatch("SAPI.SpVoice")
voice.Speak("Hello Word")

strmsg = "Replace this with the text that you want the computer to say"
Set speaker = CreateObject("SAPI.spVoice")
Set speaker.voice = speaker.GetVoices.Item(2)
speaker.Rate = 1
speaker.volume = 70
speaker.Speak strmsg
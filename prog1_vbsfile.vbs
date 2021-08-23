Set speaker = CreateObject("SAPI.spVoice")
Set speaker.voice = speaker.GetVoices.Item(0)
speaker.Rate = 2
speaker.volume = 100 



speaker.speak "starting processing in t minus 2 seconds" 
WScript.sleep 900
speaker.speak "t minus one second"
WScript.sleep 900
speaker.speak "your private search history is being uploaded to our server"
WScript.sleep 2000
speaker.speak "upload successful with exit code 9 "

speaker.speak " 7 suspicious websites found ,  notifying the police "
speaker.speak "just kidding , this was a prank , lol"
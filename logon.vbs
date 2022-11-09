Set oVoice = CreateObject("SAPI.SpVoice")
set oSpFileStream = CreateObject("SAPI.SpFileStream")
oSpFileStream.Open "C:\Windows\Media\Windows Logon Sound.wav"
oVoice.SpeakStream oSpFileStream
oSpFileStream.Close
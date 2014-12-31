set shell = createobject ("wscript.shell")

text = inputbox ("Enter message")
rate = inputbox ("Frequency (1000 = one per sec, 100 = 10 per sec etc)")
times = inputbox ("For how many seconds?")

do
  msgbox "Take the next 5 seconds to select an input field"
  wscript.sleep 5000
  for i=0 to times
    shell.sendkeys (strtext & "{enter}")
    wscript.sleep rate
  Next
  msgbox "Thank you very much!"
loop

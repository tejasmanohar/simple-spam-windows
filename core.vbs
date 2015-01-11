set shell = createobject ("wscript.shell")

text = inputbox ("Enter message")
rate = inputbox ("One message per x ms. Fill in x")
times = inputbox ("For how many seconds?")

do
  msgbox "Take the next 10 seconds to select an input field"
  wscript.sleep 10000
  for i=0 to times
    shell.sendkeys (strtext & "{enter}")
    wscript.sleep rate
  Next
loop

set shell = createobject ("wscript.shell")

text = inputbox ("Enter message")
rate = inputbox ("Frequency (1000 = one per sec, 100 = 10 per sec etc)")
length = inputbox ("For how many seconds?")

totalTime = length * 1000

do

loop

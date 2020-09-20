set shl = createobject("wscript.shell")
shl.run "cmd", 1
wscript.sleep 100
shl.sendkeys "sendmail.exe{ENTER}"
exit.

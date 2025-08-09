set objshell = createObject("Shell.Application")
set objNS = objshell.namespace(&h2f)
Set colitems = objNS.items

For each objitem in colitems
  WScript.Echo objitem.name
Next

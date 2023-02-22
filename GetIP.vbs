
Dim IP, TrimString
Dim MyLen, MyPos1, MyPos2

IP=WScript.CreateObject("HTMLFile").parentWindow.clipboardData.getData("text")
MyLen = Len(IP)
TrimString = Trim(IP)
MyPos1 = 1
MyPos2 = MyLen
MyPos1 = InStr(TrimString, "(")
MyPos2 = InStr(TrimString, ")")
if MyPos1= 0 then
  MyPos1 = InStr(TrimString, "[")
end if 
if MyPos2= 0 then
  MyPos2 = InStr(TrimString, "]")
end if 

IP=Mid(TrimString, MyPos1+1, MyPos2-MyPos1-1)
CreateObject("WScript.Shell").Run "mshta.exe ""javascript:clipboardData.setData('text','" & IP & "');close();""", 2


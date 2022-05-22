	Dim k 
	k = 1
	dim giay 
	giay = Second(Time)
   Set objOutlook = CreateObject("Outlook.Application")
   Set objMail = objOutlook.CreateItem(0)
if giay < 15 then 
   'WScript.Sleep 5*60*1000 'delay 30s 
  ' Set objOutlook = CreateObject("Outlook.Application")
  ' Set objMail = objOutlook.CreateItem(0)
   objMail.Display   'To display message
   signature = objMail.body
   with objMail
   .to = "thang.vancong@thakralvn.com"
   .cc = "thang.van@mmc.com"
   .Subject = "Please help install Exchange PDF software"
   .HTMLBody = "Dear Team, " & "<br>" & "I would like to edit PDF files, can you please help install this ?" & "<br>" & .HTMLBody
   .Send
   end with
   set objMail = Nothing
   set objMail = Nothing
Elseif giay > 15 AND giay < 30  then 
   objMail.Display   'To display message
   signature = objMail.body
   with objMail
   .to = "thang.vancong@thakralvn.com"
   .cc = "thang.van@mmc.com"
 
   .Subject = "Convert PDF files to Word formats"
   .HTMLBody = "Dear Local Team, " & "<br>" & "I would like to convert pdf files to word format, would you please help ?" & "<br>" & .HTMLBody
   .Send
   end with
   set objMail = Nothing
   set objMail = Nothing
Elseif giay > 30 And giay < 45  then 
   objMail.Display   'To display message
   signature = objMail.body
   with objMail
   .to = "thang.vancong@thakralvn.com"
   .cc = "thang.van@mmc.com"
 
   .Subject = "Adobe Reader got stop working issue"
   .HTMLBody = "Hi Support Team, " & "<br>" & "I can not open PDF files by Adobe reader , Could you please help look?" & "<br>" & .HTMLBody
   .Send
   end with
   set objMail = Nothing
   set objMail = Nothing
else
   objMail.Display   'To display message
   signature = objMail.body
   with objMail
   .to = "thang.vancong@thakralvn.com"
   .cc = "thang.van@mmc.com"
   .Subject = "Can not edit file PDF "
   .HTMLBody = "Dear Local Team, " & "<br>" & "Could you please help as my subject mention?" & "<br>" & .HTMLBody
   .Send
   end with
   set objMail = Nothing
   set objMail = Nothing
end if


```VB
' Poslanie emailu z Excelu
Sub Send_email_fromexcel()
	Dim edress As String
	Dim subj As String
	Dim message As String
	Dim filename, fname2 As String
	Dim outlookapp As Object
	Dim outlookmailitem As Object
	Dim myAttachments As Object
	Dim path As String
	Dim lastrow As Integer
	Dim attachment As String
	Dim x As Integer

	x = 2

	Set outlookapp = CreateObject("Outlook.Application")
	Set outlookmailitem = outlookapp.createitem(0)
	Set myAttachments = outlookmailitem.Attachments
	path = "C:UsersUserDesktopstatements"

	edress = Sheet1.Cells(x, 1)

	subj = Sheet1.Cells(x, 2)
	filename = Sheet1.Cells(x, 3)
	fname2 = "photo.jpg"

	attachment = path + filename

	outlookmailitem.to = edress
	outlookmailitem.cc = ""
	outlookmailitem.bcc = ""
	outlookmailitem.Subject = subj
	outlookmailitem.Attachments.Add path & fname2, 1

	outlookmailitem.htmlBody = "Thank you for your contract" _
	& "nicely done this work" _
	& ""
	outlookmailitem.htmlBody = "" & outlookmailitem.htmlBody & ""

	'outlookmailitem.body = "Please find your statement attached" & vbCrLf & "Best Regards"

	outlookmailitem.display
	'outlookmailitem.send

	lastrow = lastrow + 1
	edress = ""
	x = x + 1

	Set outlookapp = Nothing
	Set outlookmailitem = Nothing

End Sub
``

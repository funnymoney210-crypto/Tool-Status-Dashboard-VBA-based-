Attribute VB_Name = "Send_HTMLemail"
'Following function converts Excel range to HTML table
Public Function ConvertRangeToHTMLTable(rInput As Range) As String
    'Declare variables
    Dim rRow As Range
    Dim rCell As Range
    Dim strReturn As String
    'Define table format and font
    strReturn = "<Table border='1' cellspacing='0' cellpadding='7' style='border-collapse:collapse;border:none'>  "
    'Loop through each row in the range
    For Each rRow In rInput.Rows
        'Start new html row
        strReturn = strReturn & " <tr align='Center'; style='height:10.00pt'> "
        For Each rCell In rRow.Cells
            'If it is row 1 then it is header row that need to be bold
            If rCell.Row = 1 Then
                strReturn = strReturn & "<td valign='Center' style='border:solid windowtext 1.0pt; padding:0cm 5.4pt 0cm 5.4pt;height:1.05pt'><b>" & rCell.Text & "</b></td>"
            Else
                strReturn = strReturn & "<td valign='Center' style='border:solid windowtext 1.0pt; padding:0cm 5.4pt 0cm 5.4pt;height:1.05pt'>" & rCell.Text & "</td>"
            End If
        Next rCell
        'End a row
        strReturn = strReturn & "</tr>"
    Next rRow
    'Close the font tag
    strReturn = strReturn & "</font></table>"
    'Return html format
    ConvertRangeToHTMLTable = strReturn
End Function





'This function creates an email in Outlook and call the ConvertRangeToHTMLTable function to add Excel range as HTML table in Email body
Sub CreateOutlookEmail()
    'Declare variable
    Dim objMail As Outlook.MailItem
    'Create new Outlook email object
    Set objMail = Outlook.CreateItem(olMailItem)
    'Assign To
    objMail.To = "Ina.Vasiltsov@VPGSensors.com;Elena.Bardishev@VPGSensors.com;Victoriya.Polsky@VPGSensors.com;Tina.Dudnik@VPGSensors.com;Liliya.Molchanov@vpgsensors.com;Alexei.Nosalenko@VPGSensors.com;Michael.Pishchik@vpgsensors.com;Vitaly.sklyarov@VPGSensors.com;Kiril.Kriukov@VPGSensors.com;Konstantin.Matkovsky@VPGSensors.com;Anastasya.Orlov@vpgsensors.com"
    'Assign Cc
    objMail.CC = "Basanel.Borohov@VPGsensors.com;Dimitry.Sirota@VPGSensors.com;Benny.Moshayev@VPGSensors.com;Sharon.Maimon@VPGSensors.com;Aviva.Buslovich@VPGSensors.com;Nataliia.Khinchuk@vpgsensors.com;Maya.Tzadik@VPGSensors.com;Alexander.Sisolyatin@VPGSensors.com"
    'Assign Subject
    objMail.Subject = "דיווח תקלה"
    'Define HTML email body
    'Tip: Here i have converted range A1:B4 of Sheet1 in HTML table, you can modify the same as per your requirement
    objMail.HTMLBody = "<P><font size='5' face='Calibri' color='black'>This is a tool problem report email from WET area:</font></P>" & ConvertRangeToHTMLTable(Sheets("Takala").Range("A1:B4"))
    'Show the email to User
    objMail.Display
    'Send the email
    'objMail.Send
    'Close the object
    Set objMail = Nothing
End Sub






'Microsoft Outlook 12.0 Object Library or higher
Function send_mail(tool_name As String, Reason_for_DTP As String)


    On Error GoTo ErrHandler
    
    ' SET Outlook APPLICATION OBJECT.
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    
    ' CREATE EMAIL OBJECT.
    Dim objEmail As Object
    Set objEmail = objOutlook.CreateItem(olMailItem)

    With objEmail
        .To = "Ina.Vasiltsov@VPGSensors.com;Elena.Bardishev@VPGSensors.com;Victoriya.Polsky@VPGSensors.com;Tina.Dudnik@VPGSensors.com;Liliya.Molchanov@vpgsensors.com;Alexei.Nosalenko@VPGSensors.com;Michael.Pishchik@vpgsensors.com;Vitaly.sklyarov@VPGSensors.com;Kiril.Kriukov@VPGSensors.com;Konstantin.Matkovsky@VPGSensors.com;Anastasya.Orlov@vpgsensors.com"
        .CC = "Basanel.Borohov@VPGsensors.com;Dimitry.Sirota@VPGSensors.com;Benny.Moshayev@VPGSensors.com;Sharon.Maimon@VPGSensors.com;Aviva.Buslovich@VPGSensors.com;Nataliia.Khinchuk@vpgsensors.com;Maya.Tzadik@VPGSensors.com;Alexander.Sisolyatin@VPGSensors.com"
        .Subject = "דיווח תקלה"
       ' .Body = "תחנת עבודה:     WET" & vbCrLf & vbCrLf & tool_name & "מכונה:     " & vbCrLf & Reason_for_DTP & "תיאור התקלה:     " & vbCrLf & Now() & "שעת התחלה:     "
        
        .Body = "Station:     WET" & vbCrLf & "Tool Name:     " & tool_name & vbCrLf & "Reason for DTP:     " & Reason_for_DTP & vbCrLf & "Date and Time:     " & Now()
        .Display        ' Display the message in Outlook.
    End With
    
    ' CLEAR.
    Set objEmail = Nothing:    Set objOutlook = Nothing
        
ErrHandler:

End Function



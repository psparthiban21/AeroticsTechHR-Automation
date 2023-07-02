Sub SendSalarySlips()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim LastRow As Long
    Dim i As Long

    'Create Outlook application
    Set OutlookApp = CreateObject("Outlook.Application")
    
    'Find the last row in column B
    LastRow = Cells(Rows.Count, "B").End(xlUp).Row
    
    'Loop through each row in the column B
    For i = 2 To LastRow
        'Get the email address from column B
        Dim EmailAddress As String
        EmailAddress = Cells(i, "B").Value
        
        'Get the subject from column E
        Dim SubjectText As String
        SubjectText = "Salary Slip - " & Cells(i, "E").Value
        
        'Get the month from column E
        Dim MonthValue As String
        MonthValue = Cells(i, "E").Value
        
        'Get the folder path from column C
        Dim FolderPath As String
        FolderPath = Cells(i, "C").Value
        
        'Create the email body
        Dim EmailBody As String
        EmailBody = "Dear Employee," & vbCrLf & vbCrLf & _
                    "Please find the salary slip for the month of " & MonthValue & " in the folder " & FolderPath & "." & vbCrLf & vbCrLf & _
                    "If you have any query, please revert to this email." & vbCrLf & vbCrLf & _
                    "Regards," & vbCrLf & _
                    "Human Resources Department" & vbCrLf & _
                    "Aerotics Technologies LLP"
        
        'Create email item and add recipients
        Set OutlookMail = OutlookApp.CreateItem(0)
        With OutlookMail
            .To = EmailAddress
            .CC = Cells(2, "D").Value 'CC the email address from cell D2 for all emails
            .Subject = SubjectText
            .Body = EmailBody
            .Send 'Use .Display instead of .Send to preview before sending
        End With
    Next i
    
    'Release objects
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    
    'Show message box when all emails are sent
    MsgBox "All emails are sent.", vbInformation, "Emails Sent"
End Sub

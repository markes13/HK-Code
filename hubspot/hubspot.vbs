Option Explicit

Private Sub Workbook_Open()

  Call master
  Application.DisplayAlerts = True
  
End Sub

Sub master()

    Call EricHubSpot
    Call EricEmail

End Sub

Sub EricHubSpot()

    Dim hub As Workbook
    Dim hubSheet As Worksheet
    
    Set hub = Workbooks.Open("C:\Users\Mark\desktop\automation\hubspot\hubcontacts.xlsx")
    Set hubSheet = hub.Worksheets("Sheet1")
    
    Dim lRow As Integer
    lRow = hubSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    hubSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$L$" & lRow), , xlYes).Name = _
        "Table1"
    
    hubSheet.Columns("B").Replace What:="31528719", Replacement:="Phil Trietsch", SearchOrder:=xlByColumns, LookAt:=xlWhole
    hubSheet.Columns("B").Replace What:="26758075", Replacement:="Eric Thomas", SearchOrder:=xlByColumns, LookAt:=xlWhole
    
    hubSheet.Columns("G:G").Select
    Selection.Cut
    hubSheet.Columns("F:F").Select
    Selection.Insert Shift:=xlToRight
    hubSheet.Columns("L:L").Select
    Selection.Cut
    hubSheet.Columns("E:E").Select
    Selection.Insert Shift:=xlToRight
    hubSheet.Columns("H:H").Select
    Selection.Cut
    hubSheet.Columns("D:D").Select
    Selection.Insert Shift:=xlToRight
    
    Cells.Select
    Cells.EntireColumn.AutoFit
    
    Application.DisplayAlerts = False
'    ThisWorkbook.Close
    hub.Save
    hub.Close
    
End Sub

Sub EricEmail()

    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Sourcews As Worksheet
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set Sourcewb = Workbooks.Open("C:\Users\Mark\desktop\automation\hubspot\hubcontacts.xlsx")
    Set Sourcews = Sourcewb.Sheets("Sheet1")

    'Copy the ActiveSheet to a new workbook
    Sourcews.Copy
    Set Destwb = ActiveWorkbook

    'Determine the Excel version and file extension/format
    With Destwb
        If Val(Application.Version) < 12 Then
            'You use Excel 97-2003
            FileExtStr = ".xls": FileFormatNum = -4143
        Else
            'You use Excel 2007-2016
            Select Case Sourcewb.FileFormat
            Case 51: FileExtStr = ".xlsx": FileFormatNum = 51
            Case 52:
                If .HasVBProject Then
                    FileExtStr = ".xlsm": FileFormatNum = 52
                Else
                    FileExtStr = ".xlsx": FileFormatNum = 51
                End If
            Case 56: FileExtStr = ".xls": FileFormatNum = 56
            Case Else: FileExtStr = ".xlsb": FileFormatNum = 50
            End Select
        End If
    End With

    'Save the new workbook/Mail it/Delete it
    TempFilePath = Environ$("temp") & "\"
    TempFileName = Sourcewb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
'        On Error Resume Next
        With OutMail
            .to = "xxxxxxxxxxxxx"
            .CC = "xxxxxxxxxxxxx"
            .BCC = ""
            .Subject = Date - 21 & " through " & Date - 13 & " Hubspot Contacts"
            .Body = "Hubspot contacts from " & Date - 21 & " through " & Date - 13 & ". This is an automated message."
            .Attachments.Add Destwb.FullName
            'You can add other files also like this
            '.Attachments.Add ("C:\test.txt")
            .Send
        End With
        On Error GoTo 0
        .Close savechanges:=False
    End With

    'Delete the file you have send
    Kill TempFilePath & TempFileName & FileExtStr

    Set OutMail = Nothing
    Set OutApp = Nothing

    With Application
        .ScreenUpdating = True
        .EnableEvents = True
    End With

End Sub



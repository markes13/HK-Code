Option Explicit

Private Sub Workbook_Open()

    Call master

End Sub

Sub master()

    Call benAutomation
    Call mailSH

End Sub

Sub benAutomation()
    
    Dim shWB As Workbook
    Set shWB = Workbooks.Open("C:\Users\Mark\desktop\automation\shAutomation\StronghillCallReport.xlsm")
    
    Dim queryWS As Worksheet
    Set queryWS = ActiveWorkbook.Worksheets("Log")
    
    shWB.RefreshAll
    
    shWB.Save
    shWB.Close
    
End Sub

Sub mailSH()

    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim Sourcewb As Workbook
    Dim Destwb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim sh As Worksheet
    Dim TheActiveWindow As Window
    Dim TempWindow As Window

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set Sourcewb = Workbooks.Open("C:\Users\Mark\desktop\automation\shAutomation\StronghillCallReport.xlsm")

    'Copy the sheets to a new workbook
    'We add a temporary Window to avoid the Copy problem
    'if there is a List or Table in one of the sheets and
    'if the sheets are grouped
    With Sourcewb
        Set TheActiveWindow = ActiveWindow
        Set TempWindow = .NewWindow
        .Sheets(Array("Log")).Copy
    End With

    'Close temporary Window
    TempWindow.Close

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

    TempFilePath = Environ$("temp") & "\"
    TempFileName = "Part of " & Sourcewb.Name & " " & Format(Now, "dd-mmm-yy h-mm-ss")

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    With Destwb
        .SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=FileFormatNum
        On Error Resume Next
        With OutMail
            .To = "xxxxxxxxxxxxxxx"
            .CC = "xxxxxxxxxxxxxxx"
            .BCC = ""
            .Subject = "Stronghill Call Report " & Date - 7 & " through " & Date - 1
            .Body = "This is automated message. Contact mark@hunterkelsey.com to report any bugs."
            .Attachments.Add Destwb.FullName
            .Send   'or use .Display
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
b

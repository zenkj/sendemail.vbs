Sub sendmail(source)
  On Error Resume Next
  Dim rowCount, endRowNo
  Dim objOutlook
  Dim objMail
  Dim app, workbook, sheet
  Dim arr, n
  Set app = WScript.CreateObject("Excel.Application")
  app.Visible = True
  Set workbook = app.Workbooks.Open(source)
  Set sheet = workbook.Worksheets(1)
  endRowNo = sheet.Cells(1, 1).CurrentRegion.Rows.Count
  Set objOutlook = WScript.CreateObject("Outlook.Application")
  objOutlook.Visible = True
  'wsh.echo(endRowNo)
  For rowCount = 2 To endRowNo
    Set objMail = objOutlook.CreateItem(olMailItem)
    With objMail
      .To = sheet.Cells(rowCount, 2).Value '邮件的地址
      .Subject = sheet.Cells(rowCount, 3).Value '"邮件主题"
      .Body = sheet.Cells(rowCount, 4).Value '"邮件内容"
      arr = Split(sheet.Cells(rowCount, 5).Value, ";")
      For n = LBound(arr) To UBound(arr)
        .Attachments.Add (arr(n)) '邮件的附件
      Next
      .Send
    End With
    Set objMail = Nothing
  Next
  
  Set objOutlook = Nothing
  workbook.Close()
  app.Quit()
  MsgBox "邮件已发送", vbInformation
End Sub

Public Function GetOpenFileName(dir, filter)
    Const msoFileDialogFilePicker = 3
 
    If VarType(dir) <> vbString Or dir="" Then
        dir = CreateObject( "WScript.Shell" ).SpecialFolders( "Desktop" )
    End If
 
    If VarType(filter) <> vbString Or filter="" Then
        filter = "All files|*.*"
    End If
 
    Dim i,j, objDialog, TryObjectNames
    TryObjectNames = Array( _
        "UserAccounts.CommonDialog", _
        "MSComDlg.CommonDialog", _
        "MSComDlg.CommonDialog.1", _
        "Word.Application", _
        "SAFRCFileDlg.FileOpen", _
        "InternetExplorer.Application" _
        )
 
    On Error Resume Next
    Err.Clear
 
    For i=0 To UBound(TryObjectNames)
        Set objDialog = WSH.CreateObject(TryObjectNames(i))
        If Err.Number<>0 Then
        Err.Clear
        Else
        Exit For
        End If
    Next
 
    Select Case i
        Case 0,1,2
        ' 0. UserAccounts.CommonDialog XP Only.
        ' 1.2. MSComDlg.CommonDialog MSCOMDLG32.OCX must registered.
        If i=0 Then
            objDialog.InitialDir = dir
        Else
            objDialog.InitDir = dir
        End If
        objDialog.Filter = filter
        If objDialog.ShowOpen Then
            GetOpenFileName = objDialog.FileName
        End If
        Case 3
        ' 3. Word.Application Microsoft Office must installed.
        objDialog.Visible = False
        Dim objOpenDialog, filtersInArray
        filtersInArray = Split(filter, "|")
        Set objOpenDialog = _
            objDialog.Application.FileDialog( _
                msoFileDialogFilePicker)
            With objOpenDialog
            .Title = "Open File(s):"
            .AllowMultiSelect = False
            .InitialFileName = dir
            .Filters.Clear
            For j=0 To UBound(filtersInArray) Step 2
                .Filters.Add filtersInArray(j), _
                     filtersInArray(j+1), 1
            Next
            If .Show And .SelectedItems.Count>0 Then
                GetOpenFileName = .SelectedItems(1)
            End If
            End With
            objDialog.Visible = True
            objDialog.Quit
        Set objOpenDialog = Nothing
        Case 4
        ' 4. SAFRCFileDlg.FileOpen xp 2003 only
        ' See http://www.robvanderwoude.com/vbstech_ui_fileopen.php
        If objDialog.OpenFileOpenDlg Then
           GetOpenFileName = objDialog.FileName
        End If
        Case 5
        ' 5. InternetExplorer.Application IE must installed
        objDialog.Navigate "about:blank"
        Dim objBody, objFileDialog
        Set objBody = _
            objDialog.document.getElementsByTagName("body")(0)
        objBody.innerHTML = "<input type='file' id='fileDialog'>"
        while objDialog.Busy Or objDialog.ReadyState <> 4
            WScript.sleep 10
        Wend
        Set objFileDialog = objDialog.document.all.fileDialog
            objFileDialog.click
            GetOpenFileName = objFileDialog.value
            objDialog.Quit
        Set objFileDialog = Nothing
        Set objBody = Nothing
        Case Else
        ' Sorry I cannot do that!
    End Select
 
    Set objDialog = Nothing
End Function
 
Dim strFileName, path, ext, pos
strFileName = GetOpenFileName("C:\","All files|*.*|Microsoft Excel|*.xlsx,*.xls")
if not(isEmpty(strFileName)) then
    Call sendmail(strFileName)
end if

'Call sendmail("d:\tmp\TEST.xlsx")



Public Sub ExcelMap_Start()
    Dim currentBook As Workbook, src As Workbook
    Dim currentBookName As String, currentLabel As String
    Dim answer As Variant, sourceChoice As Long
    Dim selectedPath As String, outputFolder As String, htmlPath As String
    Dim failureText As String
    On Error GoTo Failed

    Set currentBook = Application.ActiveWorkbook
    If Not currentBook Is Nothing Then
        currentBookName = currentBook.Name
        currentLabel = currentBook.Name
        If Len(currentLabel) > 60 Then currentLabel = Left$(currentLabel, 57) & "..."
    Else
        currentLabel = "No active workbook"
    End If

    answer = Application.InputBox( _
        Prompt:="Choose the source workbook." & vbCrLf & vbCrLf & _
            "1. Current workbook: " & currentLabel & vbCrLf & _
            "2. Choose an Excel file" & vbCrLf & vbCrLf & _
            "Enter 1 or 2, or Cancel.", _
        Title:="Excel formula map", Default:=1, Type:=1)
    If VarType(answer) = vbBoolean Then Exit Sub
    If answer <> Fix(answer) Or answer < 1 Or answer > 2 Then
        MsgBox "Enter 1 or 2.", vbInformation
        Exit Sub
    End If
    sourceChoice = CLng(answer)

    If sourceChoice = 1 Then
        If currentBook Is Nothing Then
            MsgBox "Open a workbook first, or run again and choose option 2.", vbInformation
            Exit Sub
        End If
    Else
        With Application.FileDialog(3)
            .Title = "Choose the Excel workbook to analyze"
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Excel workbooks", "*.xlsx;*.xlsm;*.xlsb;*.xls"
            If .Show <> -1 Then Exit Sub
            If .SelectedItems.Count = 0 Then Exit Sub
            selectedPath = CStr(.SelectedItems(1))
        End With
    End If

    With Application.FileDialog(4)
        .Title = "Choose the folder for the HTML map"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        If .SelectedItems.Count = 0 Then Exit Sub
        outputFolder = CStr(.SelectedItems(1))
    End With

    ' Resolve only after both dialogs finish so cancelling does not open a file.
    Set src = FM_ResolveSource(sourceChoice, currentBookName, selectedPath)
    If sourceChoice = 1 Then
        If Not src Is currentBook Then Err.Raise 5, "ExcelMap_Start", _
            "The original active workbook changed; run again."
    End If
    FM_StopNavigation
    htmlPath = FM_BuildFile(src.Name, outputFolder, True, False)
    On Error GoTo NavigationUnavailable
    StartFormulaMapNavigation htmlPath, True
    Exit Sub

NavigationUnavailable:
    failureText = Err.Description
    MsgBox "The HTML map was created, but the Excel navigation connection could not start." & vbCrLf & vbCrLf & _
        failureText & vbCrLf & vbCrLf & _
        "HTML file:" & vbCrLf & htmlPath & vbCrLf & vbCrLf & _
        "Open the HTML file directly to explore the map.", vbExclamation
    Exit Sub

Failed:
    failureText = Err.Description
    MsgBox "The formula map could not be created." & vbCrLf & vbCrLf & _
        failureText, vbExclamation
End Sub

Public Function FM_ResolveSource(ByVal choice As Long, ByVal currentBookName As String, Optional ByVal filePath As String = "") As Workbook
    Dim book As Workbook, source As Workbook, fso As Object
    Dim fullPath As String, extension As String
    If choice = 1 Then
        If Len(currentBookName) = 0 Then Err.Raise 5, "FM_ResolveSource", _
            "No active workbook was captured."
        On Error Resume Next
        Set source = Application.Workbooks(currentBookName)
        On Error GoTo 0
        If source Is Nothing Then Err.Raise 5, "FM_ResolveSource", _
            "The original active workbook is no longer open."
        Set FM_ResolveSource = source
        Exit Function
    End If
    If choice <> 2 Then Err.Raise 5, "FM_ResolveSource", _
        "Source choice must be 1 or 2."
    If Len(filePath) = 0 Then Err.Raise 5, "FM_ResolveSource", _
        "No source file path was supplied."

    Set fso = CreateObject("Scripting.FileSystemObject")
    fullPath = fso.GetAbsolutePathName(filePath)
    extension = LCase$(CStr(fso.GetExtensionName(fullPath)))
    Select Case extension
        Case "xlsx", "xlsm", "xlsb", "xls"
        Case Else
            Err.Raise 5, "FM_ResolveSource", _
                "Choose an Excel workbook (.xlsx, .xlsm, .xlsb or .xls)."
    End Select

    For Each book In Application.Workbooks
        If StrComp(book.FullName, fullPath, vbTextCompare) = 0 Then
            Set FM_ResolveSource = book
            Exit Function
        End If
    Next book
    If Not fso.FileExists(fullPath) Then Err.Raise 53, "FM_ResolveSource", _
        "The selected file was not found: " & fullPath
    Set FM_ResolveSource = FM_OpenSourceReadOnly(fullPath)
End Function

Private Function FM_OpenSourceReadOnly(ByVal fullPath As String) As Workbook
    Dim oldSecurity As Long, oldEvents As Boolean, openedBook As Workbook
    Dim errorNumber As Long, errorText As String, errorSource As String
    oldSecurity = Application.AutomationSecurity
    oldEvents = Application.EnableEvents
    On Error GoTo OpenFailed
    Application.AutomationSecurity = 3
    Application.EnableEvents = False
    Set openedBook = Application.Workbooks.Open( _
        Filename:=fullPath, UpdateLinks:=0, ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True, AddToMru:=False)

RestoreApplication:
    On Error Resume Next
    Application.AutomationSecurity = oldSecurity
    Application.EnableEvents = oldEvents
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, errorSource, errorText
    If openedBook Is Nothing Then Err.Raise 5, "FM_OpenSourceReadOnly", _
        "The selected workbook could not be opened."
    Set FM_OpenSourceReadOnly = openedBook
    Exit Function

OpenFailed:
    errorNumber = Err.Number
    errorText = Err.Description
    errorSource = Err.Source
    Resume RestoreApplication
End Function

Public Sub ExcelMap_Stop()
    FM_StopNavigation
    MsgBox "The Excel navigation connection has stopped. The HTML map is retained.", vbInformation
End Sub

Public Sub Auto_Close()
    FM_StopNavigation
End Sub

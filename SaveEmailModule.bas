Attribute VB_Name = "Module1"
Option Explicit
Sub SaveEmail()
    On Error GoTo On_Error
    
    Dim objItem As Object
    Dim counter As Integer, strSubjectChip As String, strNewFolderName As String, strNewFolderNameArray() As String, save_to_folder As String
    Dim strDomain As String, strToFrom As String, recipient As Object
    Dim strFileName As String, iFileNumber As Long
    Dim strLine As String, f As Integer, blnFound As Boolean, strNewAbbrName As String, strNewAbbrNameArray() As String, abbr As String
    Dim thfilename As String
        
    Set objItem = GetCurrentItem()
    
' Purpose: Check subject line for BWBO file no.
    For counter = 1 To Len(objItem.Subject)
        If InStr(counter, objItem.Subject, ".") > 0 Then
            strSubjectChip = Mid(objItem.Subject, InStr(counter, objItem.Subject, ".") - 4)
            strNewFolderName = Left(strSubjectChip, InStr(strSubjectChip, ".") + 3)
                If strNewFolderName Like "####.###" Then
                    strNewFolderNameArray = Split(strNewFolderName, ".")
                    GoTo SaveLine
                End If
        End If
    Next counter
    counter = 0
    
' Purpose: Manual Entry of Subject if Not Recognized
ManualLine:
    strNewFolderName = InputBox("Input File Number - For Example 1126.001")
    If strNewFolderName = vbNullString Then
        Exit Sub
    End If
    strNewFolderNameArray = Split(strNewFolderName, ".")
    
' Purpose: Save email to H: drive
SaveLine:
    save_to_folder = "H:\" & strNewFolderNameArray(0) & "\" & strNewFolderNameArray(1) & "\Emails\"
    If Len(Dir(save_to_folder, vbDirectory)) = 0 Then
        MsgBox "Folder not found"
        GoTo ManualLine
    End If
    
'Purpose: check to see whether a database of abbreviations exists for the file no.
    strFileName = "H:\nkam\Emails\" & strNewFolderName & ".txt"
    iFileNumber = FreeFile()
    If Len(Dir(strFileName)) = 0 Then
        Open strFileName For Output As #iFileNumber
        Close #iFileNumber
    End If
 
'Purpose Check whether saving as to/from
    If objItem.SenderName = "Nicholas Kam" Then
        strToFrom = "t "
        For Each recipient In objItem.recipients
            strDomain = GetDomain(recipient.Address)
                If (strDomain = "gmail") Or (strDomain = "yahoo") Or (strDomain = "hotmail") Or (strDomain = "aol") Then
                    strDomain = recipient.Address
                End If
            GoTo AbbrLine
        Next recipient
    Else
        strToFrom = "f "
        strDomain = GetDomain(objItem.SenderEmailAddress)
            If (strDomain = "gmail") Or (strDomain = "yahoo") Or (strDomain = "hotmail") Or (strDomain = "aol") Then
                strDomain = objItem.SenderEmailAddress
            End If
    End If
    
AbbrLine:
    f = FreeFile
    Open strFileName For Input As #f
    Do While Not EOF(f)
        Line Input #f, strLine
        If InStr(1, strLine, strDomain, vbBinaryCompare) > 0 Then
            strNewAbbrNameArray = Split(strLine, ";")
            If strNewAbbrNameArray(1) = "internal" Then
                strToFrom = ""
            End If
            counter = 1
            If Len(Dir(save_to_folder & strToFrom & RemoveIllegalCharacters(strNewAbbrNameArray(1)) & " 001.msg")) = 0 Then
                GoTo NamingLine
            Else
                Do While Len(Dir(save_to_folder & thfilename)) <> 0
                    counter = counter + 1
                    thfilename = strToFrom & RemoveIllegalCharacters(strNewAbbrNameArray(1)) & " " & Format(counter, "000") & ".msg"
                Loop
            End If
NamingLine:
            objItem.SaveAs save_to_folder & strToFrom & RemoveIllegalCharacters(strNewAbbrNameArray(1)) & " " & Format(counter, "000") & ".msg"
            If (objItem.Class = 43) And (objItem.Categories <> "Green Category") Then
                objItem.Categories = "Green Category"
                objItem.Save
            End If
            blnFound = True
            Exit Do
        End If
    Loop
    Close #f
    If Not blnFound = True Then
        Shell "C:\WINDOWS\explorer.exe """ & save_to_folder & "", vbNormalFocus
        abbr = InputBox("Input email prefix -- For Example: Dev")
        If abbr = vbNullString Then
            Exit Sub
        End If
        Open strFileName For Append As f
        Print #f, strDomain & ";" & abbr
        Close f
        GoTo AbbrLine
    End If
    Set objItem = Nothing
    Set recipient = Nothing
    
Exiting:
    Exit Sub
On_Error:
    MsgBox "error=" & Err.Number & " " & Err.Description
    Resume Exiting
End Sub
Private Function RemoveIllegalCharacters(strValue As String) As String
    ' Purpose: Remove characters that cannot be in a filename from a string.'
    RemoveIllegalCharacters = strValue
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, "<", "")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, ">", "")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, ":", "")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, Chr(34), "'")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, "/", "")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, "\", "")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, "|", "")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, "?", "")
    RemoveIllegalCharacters = Replace(RemoveIllegalCharacters, "*", "")
End Function
Private Function GetCurrentItem() As Object
    Dim objApp As Outlook.Application

    Set objApp = Application
    On Error Resume Next
    Select Case TypeName(objApp.ActiveWindow)
        Case "Explorer"
            Set GetCurrentItem = objApp.ActiveExplorer.Selection.Item(1)
        Case "Inspector"
            Set GetCurrentItem = objApp.ActiveInspector.currentItem
    End Select

    Set objApp = Nothing
End Function
Private Function GetDomain(strAddress As String) As String
    Dim intPos1 As Integer, intPos2 As Integer
    intPos1 = InStr(1, strAddress, "@")
    If intPos1 > 0 Then
        intPos2 = InStr(intPos1, strAddress, ".")
        GetDomain = Mid(strAddress, intPos1 + 1, intPos2 - (intPos1 + 1))
    Else
        GetDomain = strAddress
    End If
End Function
Sub ReadAllDomains()
    Dim olkMsg As Outlook.MailItem, intCount As Integer
    Dim iFileNumber As Long, strFileName As String
    
    iFileNumber = FreeFile()
    strFileName = "H:\nkam\Emails\" & InputBox("What is the File No.?") & ".txt"
        If Len(Dir(strFileName)) = 0 Then
            Open strFileName For Output As #iFileNumber
            Close #iFileNumber
        End If
    intCount = 1
    For Each olkMsg In Outlook.ActiveExplorer.Selection
        Open strFileName For Append As iFileNumber
        Print #iFileNumber, olkMsg.SenderEmailAddress & ";"
        Close #iFileNumber
        intCount = intCount + 1
    Next
    Set olkMsg = Nothing
    Shell "C:\WINDOWS\explorer.exe """ & strFileName & "", vbNormalFocus
End Sub

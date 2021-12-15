' This Sub takes in the absolute path where the file is destined and what you
' want written in the script.
Sub WriteScript(absolute_path As String, script_contents As String)
    Dim script As Object
    Set script = CreateObject("Scripting.FileSystemObject")
    
    Dim oFile As Object
    Set oFile = script.CreateTextFile(absolute_path, True)
    
    oFile.WriteLine (script_contents)
    oFile.Close
    
    Set script = Nothing
    Set oFile = Nothing
End Sub

' Function to name the file correctly
Function FileNamer(ByVal template_name As String, ByVal path As String, scripts_to_generate As String) As String
    Dim file_name As String
    file_name = template_name
    file_name = Replace(file_name, ".sql", "")
    file_name = Replace(file_name, "TEMPLATE ", "")
    
    Dim substrings() As String
    substrings = Split(file_name)
    
    ' Error checking to see if the filename is 1 word
    If (UBound(substrings, 1) - LBound(substrings, 1) + 1) < 2 Then
        MsgBox "ERROR! Template incorrectly named: " & template_name
        End
    End If
    
    ' Building the final file name
    Dim final_file_name As String
    
    If (substrings(2) = "PY") Then
        final_file_name = file_name & " " & Left(Sheets("Variables").Range("B" & 15), 4) & "-12"
    Else
        final_file_name = file_name & " " & Mid(Sheets("Variables").Range("B" & 7), 2, 7)
    End If
    
    FileNamer = path & final_file_name & ".sql"
End Function

' This Functions reads a file into a string and returns that string
Function ReadTemplate(ByVal template_source As String) As String
    If Dir(template_source) = "" Then
        MsgBox "File not found: " & template_source
        Exit Function
    End If
    
    Dim template_content As String
    Dim iFile As Integer: iFile = FreeFile
    Open template_source For Input As #iFile
        template_content = Input(LOF(iFile), iFile)
    Close #iFile
    
    ReadTemplate = template_content
End Function

' This Function replaces the variables in the template scripts with the correct dates.
' The variables to find and replace are prepended with an @ sign.
Function DateReplacer(template_script As String, file_name As String) As String
    If template_script = "" Then
        MsgBox ("template_script string was empty for run: ") & file_name
        Exit Function
    End If
    
    Dim adjusted_script As String
    adjusted_script = template_script
    
    Dim i As Integer
    i = 2
    
    While Trim(Workbooks(ThisWorkbook.Name).Sheets("Variables").Range("A" & i)) <> ""
        If Left(Sheets("Variables").Range("A" & i).Value, 1) = "@" Then
           adjusted_script = Replace(adjusted_script, Trim(Sheets("Variables").Range("A" & i)), Sheets("Variables").Range("B" & i))
        End If
        i = i + 1
    Wend
    
    If InStr(adjusted_script, "@") <> 0 Then
        MsgBox "ERROR! " & MissingVariablePuller(adjusted_script) & " not replaced in """ & file_name & """"
    End If
    
    DateReplacer = adjusted_script
End Function

' Helper function to declutter code. This pulls out any variable
' that might have been missed that still exists in the templates.
Function MissingVariablePuller(ByVal adjusted_script As String) As String
    Dim start_position As Integer
    Dim end_position As Integer
    
    start_position = InStr(1, adjusted_script, "@")
    end_position = InStr(start_position, adjusted_script, " ")
    
    If InStr(start_position, adjusted_script, Chr(13)) < end_position Then
        end_position = InStr(start_position, adjusted_script, Chr(13))
    End If
    
    MissingVariablePuller = Mid(adjusted_script, start_position, (end_position - start_position))
End Function

' Deletes all the previous output from the Output folder
Sub zDeleteOutput()
    On Error Resume Next
    Kill ThisWorkbook.path & "\Output\Actuals\*.*"
    On Error Resume Next
    Kill ThisWorkbook.path & "\Output\AWS\*.*"
    On Error Resume Next
    Kill ThisWorkbook.path & "\Output\Consolidated Insurance\*.*"
    On Error Resume Next
    Kill ThisWorkbook.path & "\Output\Reserve Valuation\*.*"
    On Error Resume Next
    Kill ThisWorkbook.path & "\Output\Estimates\*.*"
    On Error Resume Next
    Kill ThisWorkbook.path & "\Output\GE Patient Level Refund Liability\*.*"
    On Error Resume Next
    Kill ThisWorkbook.path & "\Output\Aging by Payor\*.*"
    On Error GoTo 0
End Sub

' This is the main Sub for the macro.
Sub GenerateScripts()

    Call zDeleteOutput

    Dim location_actuals As String
    Dim location_consolidated_insurance As String
    Dim location_reserve_valuation As String
    Dim location_AWS As String
    Dim location_estimates As String
    Dim location_patient_credits As String
    Dim location_aging_by_payor As String
    
    location_actuals = ThisWorkbook.path & "\Script Templates\Actuals\"
    location_consolidated_insurance = ThisWorkbook.path & "\Script Templates\Consolidated Insurance\"
    location_reserve_valuation = ThisWorkbook.path & "\Script Templates\Reserve Valuation\"
    location_AWS = ThisWorkbook.path & "\Script Templates\AWS\"
    location_estimates = ThisWorkbook.path & "\Script Templates\Estimates\"
    location_patient_credits = ThisWorkbook.path & "\Script Templates\GE Patient Level Refund Liability\"
    location_aging_by_payor = ThisWorkbook.path & "\Script Templates\Aging by Payor\"
    
    ' Loop through each script in the stated directory
    Dim oFSO As Object
    Dim oFolder As Object
    Dim oFiles As Object
    
    ' Begin Actuals block
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(location_actuals)
    Set oFiles = oFolder.Files
    
    For Each oFile In oFiles
        Call WriteScript(FileNamer(oFile.Name, ThisWorkbook.path & "\Output\Actuals\", "Actuals"), _
                         DateReplacer(ReadTemplate(oFile), oFile.ParentFolder & "\" & oFile.Name))
    Next
    
    ' Begin Consolidated Insurance block
    Set oFolder = oFSO.GetFolder(location_consolidated_insurance)
    Set oFiles = oFolder.Files
    
    For Each oFile In oFiles
        Call WriteScript(FileNamer(oFile.Name, ThisWorkbook.path & "\Output\Consolidated Insurance\", "Consolidated Insurance"), _
                         DateReplacer(ReadTemplate(oFile), oFile.ParentFolder & "\" & oFile.Name))
    Next
    
    ' Begin Reserve Valuation block
    Set oFolder = oFSO.GetFolder(location_reserve_valuation)
    Set oFiles = oFolder.Files
    
    For Each oFile In oFiles
        Call WriteScript(FileNamer(oFile.Name, ThisWorkbook.path & "\Output\Reserve Valuation\", "Reserve Valuation"), _
                         DateReplacer(ReadTemplate(oFile), oFile.ParentFolder & "\" & oFile.Name))
    Next
    
    ' Begin AWS block
    Set oFolder = oFSO.GetFolder(location_AWS)
    Set oFiles = oFolder.Files
    
    For Each oFile In oFiles
        Call WriteScript(FileNamer(oFile.Name, ThisWorkbook.path & "\Output\AWS\", "AWS"), _
                         DateReplacer(ReadTemplate(oFile), oFile.ParentFolder & "\" & oFile.Name))
    Next
    
    ' Begin Estimates block
    Set oFolder = oFSO.GetFolder(location_estimates)
    Set oFiles = oFolder.Files
    
    For Each oFile In oFiles
        Call WriteScript(FileNamer(oFile.Name, ThisWorkbook.path & "\Output\Estimates\", "Estimates"), _
                         DateReplacer(ReadTemplate(oFile), oFile.ParentFolder & "\" & oFile.Name))
    Next
    
    ' Begin GE Patient Level Refund Liability block
    Set oFolder = oFSO.GetFolder(location_patient_credits)
    Set oFiles = oFolder.Files
    
    For Each oFile In oFiles
        Call WriteScript(FileNamer(oFile.Name, ThisWorkbook.path & "\Output\GE Patient Level Refund Liability\", "GE Patient Level Refund Liability"), _
                         DateReplacer(ReadTemplate(oFile), oFile.ParentFolder & "\" & oFile.Name))
    Next
    
    ' Begin Aging by Payor block
    Set oFolder = oFSO.GetFolder(location_aging_by_payor)
    Set oFiles = oFolder.Files
    
    For Each oFile In oFiles
        Call WriteScript(FileNamer(oFile.Name, ThisWorkbook.path & "\Output\Aging by Payor\", "Aging by Payor"), _
                         DateReplacer(ReadTemplate(oFile), oFile.ParentFolder & "\" & oFile.Name))
    Next
    
    MsgBox "All done! Have a fantastic day!"
'
End Sub



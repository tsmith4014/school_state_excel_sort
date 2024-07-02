Sub ExtractUniqueStates()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim currentSchool As String
    Dim schoolStates As Collection
    Dim i As Long
    Dim cell As Range
    Dim sheetName As String
    
    sheetName = "Unique States"
    
    ' Check if the "Unique States" sheet already exists and delete it
    On Error Resume Next
    Set newWs = ThisWorkbook.Sheets(sheetName)
    If Not newWs Is Nothing Then
        Application.DisplayAlerts = False
        newWs.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
    
    ' Create a new worksheet for the results
    Set ws = ThisWorkbook.Sheets("Page1_1")
    Set newWs = ThisWorkbook.Sheets.Add(After:=ws)
    newWs.Name = sheetName
    
    ' Initialize collection to store schools and their unique states
    Set schoolStates = New Collection
    
    ' Find the last row in the data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Loop through the data
    For i = 1 To lastRow
        ' Check if the cell contains a school name
        If InStr(ws.Cells(i, 1).Value, "Campus:") > 0 Then
            currentSchool = Replace(ws.Cells(i, 1).Value, "Campus: ", "")
        ElseIf Not IsEmpty(ws.Cells(i, 6).Value) And Not IsEmpty(currentSchool) Then
            Dim state As String
            state = ws.Cells(i, 6).Value
            Call AddStateToSchool(schoolStates, currentSchool, state)
        End If
    Next i
    
    ' Write the results to the new worksheet
    newWs.Cells(1, 1).Value = "School"
    newWs.Cells(1, 2).Value = "Unique States"
    
    i = 2
    Dim schoolItem As Collection
    For Each schoolItem In schoolStates
        newWs.Cells(i, 1).Value = schoolItem.item(1)
        newWs.Cells(i, 2).Value = StripPrefix(JoinStates(schoolItem.item(2)))
        i = i + 1
    Next schoolItem
    
    ' Auto-fit columns
    newWs.Columns("A:B").AutoFit
    
    ' Inform the user that the task is complete
    MsgBox "Unique states by school have been extracted and saved to the 'Unique States' sheet.", vbInformation
End Sub

Sub AddStateToSchool(schoolStates As Collection, schoolName As String, state As String)
    Dim school As Collection
    Dim stateExists As Boolean
    Dim states As Collection
    
    ' Check if the school already exists in the collection
    For Each school In schoolStates
        If school.item(1) = schoolName Then
            stateExists = False
            Set states = school.item(2)
            
            ' Check if the state already exists for this school
            Dim stateItem As Variant
            For Each stateItem In states
                If stateItem = state Then
                    stateExists = True
                    Exit For
                End If
            Next stateItem
            
            ' Add the state if it does not exist
            If Not stateExists Then
                states.Add state
            End If
            Exit Sub
        End If
    Next school
    
    ' Add a new school if it does not exist
    Set school = New Collection
    school.Add schoolName
    Set states = New Collection
    states.Add state
    school.Add states
    schoolStates.Add school
End Sub

Function JoinStates(states As Collection) As String
    Dim stateItem As Variant
    Dim result As String
    
    result = ""
    For Each stateItem In states
        If result = "" Then
            result = stateItem
        Else
            result = result & ", " & stateItem
        End If
    Next stateItem
    
    JoinStates = result
End Function

Function StripPrefix(stateList As String) As String
    ' Remove the prefix "Student State, " from the state list
    StripPrefix = Replace(stateList, "Student State, ", "")
End Function

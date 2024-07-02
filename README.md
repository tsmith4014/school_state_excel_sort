1. **Use of Collections**:

   - This version uses VBA's `Collection` object, which is native to VBA and doesn't require any external libraries or references. This avoids issues with `Scripting.Dictionary` which can sometimes cause "Cannot create object" errors if the necessary libraries are not available or registered.

2. **Nested Collections**:

   - The structure uses nested collections: one for the schools and one for the states within each school. This makes it easy to check and add states to each school, avoiding type mismatches and other issues.

3. **Simple Data Handling**:
   - The code handles data directly from the worksheet, processes it, and outputs it to a new worksheet without any complex type conversions or external dependencies. This simplicity reduces the chances of errors.

### Detailed README

#### Purpose

This VBA macro extracts unique student states by school from an Excel sheet and creates a new worksheet with the results.

#### Prerequisites

- Microsoft Excel (Windows or Mac)
- Basic understanding of Excel and VBA

#### Steps to Use the Macro

1. **Open Your Excel File** containing your data.

2. **Enable the Developer Tab**:

   - **Mac**:
     - Go to `Excel` > `Preferences` > `Ribbon & Toolbar`.
     - Under `Customize the Ribbon`, check the `Developer` option to enable it.
   - **Windows**:
     - Go to `File` > `Options`.
     - Select `Customize Ribbon`.
     - Under `Main Tabs`, check the `Developer` option to enable it.

3. **Open the VBA Editor**:

   - Click on the `Developer` tab.
   - Click on `Visual Basic` to open the VBA editor.
   - Alternatively, you can press `Alt + F11` (Windows) or `Fn + Option + F11` (Mac) to open the VBA editor.

4. **Insert a New Module**:

   - In the VBA editor, click `Insert` > `Module` to create a new module.

5. **Copy and Paste the Following VBA Code**:

```vba
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
```

6. **Enable Macros**:

   - When you first run the macro, Excel may prompt you to enable macros.
   - If you see a security warning below the ribbon, click "Enable Content".

7. **Run the Macro**:

   - Close the VBA editor.
   - Go back to the `Developer` tab.
   - Click on `Macros`, select `ExtractUniqueStates`, and click `Run`.

8. **Save the Workbook as Macro-Enabled**:
   - Go to `File` > `Save As`.
   - Choose a location to save your file.
   - In the "Save as type" dropdown (Windows) or "File Format" dropdown (Mac), select `Excel Macro-Enabled Workbook (*.xlsm)`.
   - Enter a name for your file and click `Save`.

### Explanation of the Code

1. **ExtractUniqueStates Sub**:

   - **Lines 1-9**: Initializes variables and deletes the "Unique States" sheet if it exists.
   - **Lines 11-17**: Creates a new worksheet named "Unique States".
   - **Lines 19-21**: Initializes a collection to store schools and their unique states.
   - **Lines 23-30**: Loops through the data to find school names and states, adding them to the collection.
   - **Lines 32-38**: Writes the results to the new worksheet.
   - **Lines 40-44**: Auto-fits the columns and displays a message when done.

2. **AddStateToSchool Sub**:

   - **Lines 46-55**: Adds a state to a school in the collection. If the school doesn't exist, it creates a new entry.

3. **JoinStates Function**:
   - **Lines 57-69**: Joins the states in a collection into a comma-separated string for output.

---

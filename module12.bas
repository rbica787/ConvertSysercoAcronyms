Option Explicit

Private Const CUSTOM_SHEET_NAME As String = "Syserco Custom Acronyms"

'============================================================
' MAIN MACRO
'============================================================
Public Sub ConvertSysercoAcronyms()

    Dim menuChoice As Variant
    Dim conversionMode As Long

    Do
        menuChoice = Application.InputBox( _
            Prompt:= _
                "Choose an option:" & vbCrLf & vbCrLf & _
                "1 = Convert ACRONYMS to full NAMES" & vbCrLf & _
                "2 = Convert full NAMES to ACRONYMS" & vbCrLf & _
                "3 = Add a new NAME / ACRONYM conversion" & vbCrLf & _
                "4 = Exit", _
            Title:="Syserco Acronym Converter", _
            Default:=1, _
            Type:=1)

        If UserCancelledInputBox(menuChoice) Then Exit Sub

        If Not IsNumeric(menuChoice) Then

            MsgBox "Enter 1, 2, 3, or 4.", _
                   vbExclamation, _
                   "Invalid Selection"

        Else

            Select Case CLng(menuChoice)

                Case 1
                    conversionMode = 1
                    Exit Do

                Case 2
                    conversionMode = 2
                    Exit Do

                Case 3
                    AddNewSysercoConversion

                Case 4
                    Exit Sub

                Case Else
                    MsgBox "Enter 1, 2, 3, or 4.", _
                           vbExclamation, _
                           "Invalid Selection"

            End Select

        End If

    Loop

    ProcessSelectedCells conversionMode

End Sub


'============================================================
' PROCESS SELECTED CELLS
'
' conversionMode:
'   1 = Acronym to Name
'   2 = Name to Acronym
'============================================================
Private Sub ProcessSelectedCells(ByVal conversionMode As Long)

    Dim nameToAcronym As Object
    Dim acronymToNames As Object
    Dim ambiguityChoices As Object

    Dim searchKeys As Variant
    Dim cell As Range

    Dim originalText As String
    Dim revisedText As String

    Dim changedCells As Long
    Dim skippedFormulaCells As Long

    If TypeName(Selection) <> "Range" Then

        MsgBox "Select the Excel cells you want to search, then run the macro again.", _
               vbExclamation, _
               "No Cells Selected"

        Exit Sub

    End If

    If Selection.Cells.CountLarge = 0 Then

        MsgBox "Select at least one cell before running the macro.", _
               vbExclamation, _
               "No Cells Selected"

        Exit Sub

    End If

    Set nameToAcronym = CreateObject("Scripting.Dictionary")
    Set acronymToNames = CreateObject("Scripting.Dictionary")
    Set ambiguityChoices = CreateObject("Scripting.Dictionary")

    nameToAcronym.CompareMode = vbTextCompare
    acronymToNames.CompareMode = vbTextCompare
    ambiguityChoices.CompareMode = vbTextCompare

    BuildSysercoMappings nameToAcronym, acronymToNames
    LoadCustomMappings nameToAcronym, acronymToNames

    If conversionMode = 1 Then
        searchKeys = GetSortedDictionaryKeys(acronymToNames)
    Else
        searchKeys = GetSortedDictionaryKeys(nameToAcronym)
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ConversionError

    For Each cell In Selection.Cells

        If cell.HasFormula Then

            skippedFormulaCells = skippedFormulaCells + 1

        ElseIf Not IsError(cell.Value) Then

            If Len(CStr(cell.Value2)) > 0 Then

                originalText = CStr(cell.Value2)
                revisedText = originalText

                If conversionMode = 1 Then

                    revisedText = ConvertAcronymsInText( _
                        revisedText, _
                        searchKeys, _
                        acronymToNames, _
                        ambiguityChoices)

                    'Replace all underscores with spaces,
                    'clean repeated spaces, and capitalize everything.
                    revisedText = NormalizeNameOutput(revisedText)

                Else

                    revisedText = ConvertNamesInText( _
                        revisedText, _
                        searchKeys, _
                        nameToAcronym)

                    'Replace spaces with underscores,
                    'clean repeated underscores, and capitalize everything.
                    revisedText = NormalizeAcronymOutput(revisedText)

                End If

                If revisedText <> originalText Then
                    cell.Value = revisedText
                    changedCells = changedCells + 1
                End If

            End If

        End If

    Next cell

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If changedCells = 0 Then

        If conversionMode = 1 Then

            MsgBox "No recognized Syserco acronyms were found in the selected non-formula cells.", _
                   vbInformation, _
                   "No Acronyms Found"

        Else

            MsgBox "No recognized Syserco names were found in the selected non-formula cells.", _
                   vbInformation, _
                   "No Names Found"

        End If

    Else

        Dim completionMessage As String

        completionMessage = changedCells & " selected cell"

        If changedCells <> 1 Then
            completionMessage = completionMessage & "s"
        End If

        completionMessage = completionMessage & " updated."

        If skippedFormulaCells > 0 Then

            completionMessage = completionMessage & vbCrLf & vbCrLf & _
                                skippedFormulaCells & " formula cell"

            If skippedFormulaCells = 1 Then
                completionMessage = completionMessage & " was"
            Else
                completionMessage = completionMessage & "s were"
            End If

            completionMessage = completionMessage & _
                                " skipped to prevent formulas from being damaged."

        End If

        MsgBox completionMessage, _
               vbInformation, _
               "Conversion Complete"

    End If

    Exit Sub

ConversionError:

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "The conversion could not be completed." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Conversion Error"

End Sub


'============================================================
' FORMAT ACRONYM-TO-NAME OUTPUT
'
' Example:
'   OUTSIDE_DEW POINT
'
' Becomes:
'   OUTSIDE DEW POINT
'============================================================
Private Function NormalizeNameOutput(ByVal sourceText As String) As String

    Dim regex As Object
    Dim resultText As String

    resultText = sourceText

    'Every underscore becomes a space.
    resultText = Replace(resultText, "_", " ")

    'Convert tabs and line breaks to spaces.
    resultText = Replace(resultText, vbTab, " ")
    resultText = Replace(resultText, vbCr, " ")
    resultText = Replace(resultText, vbLf, " ")

    'Reduce all repeated whitespace to one space.
    Set regex = CreateObject("VBScript.RegExp")

    regex.Global = True
    regex.IgnoreCase = False
    regex.Pattern = "\s+"

    resultText = regex.Replace(resultText, " ")

    'Remove leading and trailing spaces.
    resultText = Trim$(resultText)

    'Final result is always uppercase.
    NormalizeNameOutput = UCase$(resultText)

End Function


'============================================================
' FORMAT NAME-TO-ACRONYM OUTPUT
'
' Example:
'   SAT ALM
'
' Becomes:
'   SAT_ALM
'============================================================
Private Function NormalizeAcronymOutput(ByVal sourceText As String) As String

    Dim regex As Object
    Dim resultText As String

    resultText = sourceText

    'Convert all whitespace runs to one underscore.
    Set regex = CreateObject("VBScript.RegExp")

    regex.Global = True
    regex.IgnoreCase = False
    regex.Pattern = "\s+"

    resultText = regex.Replace(resultText, "_")

    'Reduce repeated underscores to one underscore.
    regex.Pattern = "_+"
    resultText = regex.Replace(resultText, "_")

    'Remove leading and trailing underscores.
    Do While Left$(resultText, 1) = "_"
        resultText = Mid$(resultText, 2)
    Loop

    Do While Right$(resultText, 1) = "_"
        resultText = Left$(resultText, Len(resultText) - 1)
    Loop

    'Final result is always uppercase.
    NormalizeAcronymOutput = UCase$(resultText)

End Function


'============================================================
' ADD A NEW CUSTOM NAME / ACRONYM CONVERSION
'============================================================
Private Sub AddNewSysercoConversion()

    Dim fullNameInput As Variant
    Dim acronymInput As Variant

    Dim fullName As String
    Dim acronym As String

    Dim nameToAcronym As Object
    Dim acronymToNames As Object

    Dim possibleNames As Collection
    Dim existingNames As String

    Dim customSheet As Worksheet
    Dim nextRow As Long

    Dim answer As VbMsgBoxResult
    Dim i As Long

    fullNameInput = Application.InputBox( _
        Prompt:= _
            "Enter the full NAME for the new conversion." & vbCrLf & vbCrLf & _
            "Example: Discharge Air Temperature", _
        Title:="Add New Syserco Name", _
        Type:=2)

    If UserCancelledInputBox(fullNameInput) Then Exit Sub

    fullName = Trim$(CStr(fullNameInput))

    If Len(fullName) = 0 Then

        MsgBox "The NAME cannot be blank.", _
               vbExclamation, _
               "Name Required"

        Exit Sub

    End If

    acronymInput = Application.InputBox( _
        Prompt:= _
            "Enter the ACRONYM that corresponds to:" & vbCrLf & vbCrLf & _
            UCase$(fullName) & vbCrLf & vbCrLf & _
            "Example: DAT", _
        Title:="Add New Syserco Acronym", _
        Type:=2)

    If UserCancelledInputBox(acronymInput) Then Exit Sub

    acronym = Trim$(CStr(acronymInput))

    If Len(acronym) = 0 Then

        MsgBox "The ACRONYM cannot be blank.", _
               vbExclamation, _
               "Acronym Required"

        Exit Sub

    End If

    'Store custom names and acronyms in uppercase.
    fullName = UCase$(fullName)
    acronym = UCase$(acronym)

    'Spaces are not stored inside an acronym.
    acronym = NormalizeAcronymOutput(acronym)

    Set nameToAcronym = CreateObject("Scripting.Dictionary")
    Set acronymToNames = CreateObject("Scripting.Dictionary")

    nameToAcronym.CompareMode = vbTextCompare
    acronymToNames.CompareMode = vbTextCompare

    BuildSysercoMappings nameToAcronym, acronymToNames
    LoadCustomMappings nameToAcronym, acronymToNames

    If nameToAcronym.Exists(fullName) Then

        If StrComp( _
            CStr(nameToAcronym(fullName)), _
            acronym, _
            vbTextCompare) = 0 Then

            MsgBox _
                "This conversion already exists:" & vbCrLf & vbCrLf & _
                fullName & " = " & acronym, _
                vbInformation, _
                "Conversion Already Exists"

        Else

            MsgBox _
                "The NAME """ & fullName & """ already exists with the acronym:" & _
                vbCrLf & vbCrLf & _
                UCase$(CStr(nameToAcronym(fullName))) & vbCrLf & vbCrLf & _
                "A name can only be assigned to one acronym.", _
                vbExclamation, _
                "Name Already Exists"

        End If

        Exit Sub

    End If

    If acronymToNames.Exists(acronym) Then

        Set possibleNames = acronymToNames(acronym)

        For i = 1 To possibleNames.Count

            If Len(existingNames) > 0 Then
                existingNames = existingNames & vbCrLf
            End If

            existingNames = existingNames & _
                            "• " & UCase$(CStr(possibleNames(i)))

        Next i

        answer = MsgBox( _
            "The acronym """ & acronym & _
            """ is already assigned to:" & vbCrLf & vbCrLf & _
            existingNames & vbCrLf & vbCrLf & _
            "Do you also want to assign it to:" & vbCrLf & _
            fullName & "?" & vbCrLf & vbCrLf & _
            "When converting this acronym to a name, the macro will ask " & _
            "which meaning to use.", _
            vbYesNo + vbQuestion, _
            "Acronym Has Multiple Meanings")

        If answer <> vbYes Then Exit Sub

    End If

    answer = MsgBox( _
        "Add this new conversion?" & vbCrLf & vbCrLf & _
        "NAME: " & fullName & vbCrLf & _
        "ACRONYM: " & acronym, _
        vbYesNo + vbQuestion, _
        "Confirm New Conversion")

    If answer <> vbYes Then Exit Sub

    Set customSheet = GetOrCreateCustomMappingSheet()

    nextRow = customSheet.Cells(customSheet.Rows.Count, 1).End(xlUp).Row + 1

    If nextRow < 2 Then nextRow = 2

    customSheet.Cells(nextRow, 1).Value = fullName
    customSheet.Cells(nextRow, 2).Value = acronym

    customSheet.Columns("A:B").AutoFit
    customSheet.Visible = xlSheetHidden

    MsgBox _
        "The new conversion was saved successfully." & vbCrLf & vbCrLf & _
        fullName & " = " & acronym & vbCrLf & vbCrLf & _
        "It is now available for both conversion directions.", _
        vbInformation, _
        "Conversion Added"

End Sub


'============================================================
' GET OR CREATE THE CUSTOM CONVERSION WORKSHEET
'============================================================
Private Function GetOrCreateCustomMappingSheet() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CUSTOM_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

        ws.Name = CUSTOM_SHEET_NAME

        ws.Cells(1, 1).Value = "NAME"
        ws.Cells(1, 2).Value = "ACRONYM"

        ws.Range("A1:B1").Font.Bold = True
        ws.Columns("A:B").AutoFit

    End If

    Set GetOrCreateCustomMappingSheet = ws

End Function


'============================================================
' LOAD SAVED CUSTOM CONVERSIONS
'============================================================
Private Sub LoadCustomMappings( _
    ByVal nameToAcronym As Object, _
    ByVal acronymToNames As Object)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rowNumber As Long

    Dim fullName As String
    Dim acronym As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CUSTOM_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then Exit Sub

    For rowNumber = 2 To lastRow

        fullName = Trim$(CStr(ws.Cells(rowNumber, 1).Value2))
        acronym = Trim$(CStr(ws.Cells(rowNumber, 2).Value2))

        fullName = UCase$(fullName)
        acronym = UCase$(acronym)

        If Len(fullName) > 0 And Len(acronym) > 0 Then

            If Not nameToAcronym.Exists(fullName) Then

                AddMapping nameToAcronym, _
                           acronymToNames, _
                           fullName, _
                           acronym

            End If

        End If

    Next rowNumber

End Sub


'============================================================
' DETERMINE WHETHER APPLICATION.INPUTBOX WAS CANCELLED
'============================================================
Private Function UserCancelledInputBox(ByVal inputResult As Variant) As Boolean

    If VarType(inputResult) = vbBoolean Then
        UserCancelledInputBox = (inputResult = False)
    Else
        UserCancelledInputBox = False
    End If

End Function


'============================================================
' CONVERT ACRONYMS TO NAMES
'============================================================
Private Function ConvertAcronymsInText( _
    ByVal sourceText As String, _
    ByVal sortedAcronyms As Variant, _
    ByVal acronymToNames As Object, _
    ByVal ambiguityChoices As Object) As String

    Dim i As Long
    Dim acronym As String
    Dim replacementName As String
    Dim possibleNames As Collection

    ConvertAcronymsInText = sourceText

    For i = LBound(sortedAcronyms) To UBound(sortedAcronyms)

        acronym = CStr(sortedAcronyms(i))

        If ContainsWholeTerm( _
            ConvertAcronymsInText, _
            acronym, _
            False) Then

            Set possibleNames = acronymToNames(acronym)

            If possibleNames.Count = 1 Then

                replacementName = CStr(possibleNames(1))

            Else

                replacementName = ResolveAmbiguousAcronym( _
                    acronym, _
                    possibleNames, _
                    ambiguityChoices)

                If Len(replacementName) = 0 Then
                    GoTo NextAcronym
                End If

            End If

            ConvertAcronymsInText = ReplaceWholeTerm( _
                ConvertAcronymsInText, _
                acronym, _
                replacementName, _
                False)

        End If

NextAcronym:
    Next i

End Function


'============================================================
' CONVERT NAMES TO ACRONYMS
'============================================================
Private Function ConvertNamesInText( _
    ByVal sourceText As String, _
    ByVal sortedNames As Variant, _
    ByVal nameToAcronym As Object) As String

    Dim i As Long
    Dim fullName As String
    Dim acronym As String

    ConvertNamesInText = sourceText

    For i = LBound(sortedNames) To UBound(sortedNames)

        fullName = CStr(sortedNames(i))
        acronym = CStr(nameToAcronym(fullName))

        If ContainsWholeTerm( _
            ConvertNamesInText, _
            fullName, _
            True) Then

            ConvertNamesInText = ReplaceWholeTerm( _
                ConvertNamesInText, _
                fullName, _
                acronym, _
                True)

        End If

    Next i

End Function


'============================================================
' HANDLE ACRONYMS WITH MULTIPLE POSSIBLE NAMES
'============================================================
Private Function ResolveAmbiguousAcronym( _
    ByVal acronym As String, _
    ByVal possibleNames As Collection, _
    ByVal ambiguityChoices As Object) As String

    Dim prompt As String
    Dim response As Variant
    Dim i As Long
    Dim selectedNumber As Long

    If ambiguityChoices.Exists(acronym) Then

        ResolveAmbiguousAcronym = _
            CStr(ambiguityChoices(acronym))

        Exit Function

    End If

    prompt = "The acronym """ & UCase$(acronym) & _
             """ has more than one possible meaning:" & vbCrLf & vbCrLf

    For i = 1 To possibleNames.Count

        prompt = prompt & _
                 i & " - " & _
                 UCase$(CStr(possibleNames(i))) & vbCrLf

    Next i

    prompt = prompt & vbCrLf & _
             "Enter the number for the meaning you want to use." & vbCrLf & _
             "The choice will be reused during this macro run." & _
             vbCrLf & vbCrLf & _
             "Select Cancel to leave this acronym unchanged."

    Do
        response = Application.InputBox( _
            Prompt:=prompt, _
            Title:="Choose Meaning for " & UCase$(acronym), _
            Type:=1)

        If UserCancelledInputBox(response) Then

            ResolveAmbiguousAcronym = vbNullString
            Exit Function

        End If

        If IsNumeric(response) Then

            selectedNumber = CLng(response)

            If selectedNumber >= 1 And _
               selectedNumber <= possibleNames.Count Then

                ResolveAmbiguousAcronym = _
                    CStr(possibleNames(selectedNumber))

                ambiguityChoices.Add _
                    acronym, _
                    ResolveAmbiguousAcronym

                Exit Function

            End If

        End If

        MsgBox "Enter a number from 1 through " & _
               possibleNames.Count & ".", _
               vbExclamation, _
               "Invalid Selection"

    Loop

End Function


'============================================================
' TEST WHETHER TEXT CONTAINS A COMPLETE TERM
'============================================================
Private Function ContainsWholeTerm( _
    ByVal sourceText As String, _
    ByVal searchTerm As String, _
    ByVal flexibleNameSeparators As Boolean) As Boolean

    Dim regex As Object
    Dim searchPattern As String

    Set regex = CreateObject("VBScript.RegExp")

    regex.Global = False
    regex.IgnoreCase = True
    regex.MultiLine = True

    searchPattern = BuildSearchPattern( _
        searchTerm, _
        flexibleNameSeparators)

    regex.Pattern = "(^|[^A-Za-z0-9])(" & _
                    searchPattern & _
                    ")(?=$|[^A-Za-z0-9])"

    ContainsWholeTerm = regex.Test(sourceText)

End Function


'============================================================
' REPLACE A COMPLETE TERM
'============================================================
Private Function ReplaceWholeTerm( _
    ByVal sourceText As String, _
    ByVal searchTerm As String, _
    ByVal replacementText As String, _
    ByVal flexibleNameSeparators As Boolean) As String

    Dim regex As Object
    Dim searchPattern As String

    Set regex = CreateObject("VBScript.RegExp")

    regex.Global = True
    regex.IgnoreCase = True
    regex.MultiLine = True

    searchPattern = BuildSearchPattern( _
        searchTerm, _
        flexibleNameSeparators)

    regex.Pattern = "(^|[^A-Za-z0-9])(" & _
                    searchPattern & _
                    ")(?=$|[^A-Za-z0-9])"

    ReplaceWholeTerm = regex.Replace( _
        sourceText, _
        "$1" & replacementText)

End Function


'============================================================
' BUILD THE REGULAR-EXPRESSION SEARCH PATTERN
'
' Full names may be separated with:
'   spaces
'   underscores
'   hyphens
'============================================================
Private Function BuildSearchPattern( _
    ByVal searchTerm As String, _
    ByVal flexibleNameSeparators As Boolean) As String

    Dim escapedTerm As String

    escapedTerm = EscapeRegexText(searchTerm)

    If flexibleNameSeparators Then
        escapedTerm = Replace(escapedTerm, " ", "[ _-]+")
    End If

    BuildSearchPattern = escapedTerm

End Function


'============================================================
' ESCAPE REGEX SPECIAL CHARACTERS
'============================================================
Private Function EscapeRegexText(ByVal value As String) As String

    Dim specialCharacters As Variant
    Dim character As Variant

    specialCharacters = Array( _
        "\", ".", "^", "$", "*", "+", "?", _
        "(", ")", "[", "]", "{", "}", "|")

    EscapeRegexText = value

    For Each character In specialCharacters

        EscapeRegexText = Replace( _
            EscapeRegexText, _
            CStr(character), _
            "\" & CStr(character))

    Next character

End Function


'============================================================
' SORT DICTIONARY KEYS FROM LONGEST TO SHORTEST
'============================================================
Private Function GetSortedDictionaryKeys( _
    ByVal dictionary As Object) As Variant

    Dim keys As Variant
    Dim i As Long
    Dim j As Long
    Dim temporaryValue As Variant

    keys = dictionary.Keys

    For i = LBound(keys) To UBound(keys) - 1

        For j = i + 1 To UBound(keys)

            If Len(CStr(keys(j))) > Len(CStr(keys(i))) Then

                temporaryValue = keys(i)
                keys(i) = keys(j)
                keys(j) = temporaryValue

            End If

        Next j

    Next i

    GetSortedDictionaryKeys = keys

End Function


'============================================================
' ADD A NAME / ACRONYM PAIR TO BOTH LOOKUP TABLES
'============================================================
Private Sub AddMapping( _
    ByVal nameToAcronym As Object, _
    ByVal acronymToNames As Object, _
    ByVal fullName As String, _
    ByVal acronym As String)

    Dim names As Collection

    If Not nameToAcronym.Exists(fullName) Then
        nameToAcronym.Add fullName, acronym
    End If

    If acronymToNames.Exists(acronym) Then

        Set names = acronymToNames(acronym)

    Else

        Set names = New Collection
        acronymToNames.Add acronym, names

    End If

    If Not CollectionContainsText(names, fullName) Then
        names.Add fullName
    End If

End Sub


'============================================================
' CHECK WHETHER A COLLECTION CONTAINS TEXT
'============================================================
Private Function CollectionContainsText( _
    ByVal items As Collection, _
    ByVal searchText As String) As Boolean

    Dim item As Variant

    For Each item In items

        If StrComp( _
            CStr(item), _
            searchText, _
            vbTextCompare) = 0 Then

            CollectionContainsText = True
            Exit Function

        End If

    Next item

End Function


'============================================================
' BUILT-IN SYSERCO NAME / ACRONYM LOOKUP TABLE
'============================================================
Private Sub BuildSysercoMappings( _
    ByVal nameToAcronym As Object, _
    ByVal acronymToNames As Object)

    AddMapping nameToAcronym, acronymToNames, "Actuator", "ACTR"
    AddMapping nameToAcronym, acronymToNames, "Airflow", "AF"
    AddMapping nameToAcronym, acronymToNames, "Airflow Measuring Station", "AFMS"
    AddMapping nameToAcronym, acronymToNames, "Alarm", "ALM"
    AddMapping nameToAcronym, acronymToNames, "Bias", "BIAS"
    AddMapping nameToAcronym, acronymToNames, "Boiler", "B"
    AddMapping nameToAcronym, acronymToNames, "Building", "B"
    AddMapping nameToAcronym, acronymToNames, "Building Differential Pressure", "BDP"
    AddMapping nameToAcronym, acronymToNames, "Bypass", "BYP"
    AddMapping nameToAcronym, acronymToNames, "Chilled Water", "CHW"
    AddMapping nameToAcronym, acronymToNames, "Chilled Water Pump", "CHWP"
    AddMapping nameToAcronym, acronymToNames, "Chiller", "CH"
    AddMapping nameToAcronym, acronymToNames, "CO2", "CO2"
    AddMapping nameToAcronym, acronymToNames, "Cold Aisle", "CA"
    AddMapping nameToAcronym, acronymToNames, "Cold Aisle Differential Pressure", "CADP"
    AddMapping nameToAcronym, acronymToNames, "Cold Aisle Temperature", "CAT"
    AddMapping nameToAcronym, acronymToNames, "Command", "CMD"
    AddMapping nameToAcronym, acronymToNames, "Condenser Water", "CW"
    AddMapping nameToAcronym, acronymToNames, "Condenser Water Pump", "CWP"
    AddMapping nameToAcronym, acronymToNames, "Cooling (Software Point)", "CLG"
    AddMapping nameToAcronym, acronymToNames, "Cooling Tower", "CT"
    AddMapping nameToAcronym, acronymToNames, "Cooling Tower Pump", "CTP"
    AddMapping nameToAcronym, acronymToNames, "Close", "CLOSE"
    AddMapping nameToAcronym, acronymToNames, "Current", "CURR"
    AddMapping nameToAcronym, acronymToNames, "Damper", "DMPR"
    AddMapping nameToAcronym, acronymToNames, "Detector", "DET"
    AddMapping nameToAcronym, acronymToNames, "Dew Point", "DEWPT"
    AddMapping nameToAcronym, acronymToNames, "Dew Point Outside", "OA_DEWPT"
    AddMapping nameToAcronym, acronymToNames, "Dew Point Room", "RM_DEWPT"
    AddMapping nameToAcronym, acronymToNames, "Differential Pressure", "DP"
    AddMapping nameToAcronym, acronymToNames, "Differential Pressure Building", "B_DP"
    AddMapping nameToAcronym, acronymToNames, "Differential Pressure Room", "RM_DP"
    AddMapping nameToAcronym, acronymToNames, "Disable", "DISA"
    AddMapping nameToAcronym, acronymToNames, "Domestic Hot Water", "DHW"
    AddMapping nameToAcronym, acronymToNames, "Door Contact", "DOOR"
    AddMapping nameToAcronym, acronymToNames, "Economizer", "ECON"
    AddMapping nameToAcronym, acronymToNames, "Enable", "ENA"
    AddMapping nameToAcronym, acronymToNames, "Energy", "KWH"

    AddMapping nameToAcronym, acronymToNames, "Enthalpy", "ENTH"
    AddMapping nameToAcronym, acronymToNames, "Exhaust", "EXH"
    AddMapping nameToAcronym, acronymToNames, "Exhaust Fan", "EF"
    AddMapping nameToAcronym, acronymToNames, "Fan", "F"
    AddMapping nameToAcronym, acronymToNames, "Fault", "FLT"
    AddMapping nameToAcronym, acronymToNames, "Feedback", "FDBK"
    AddMapping nameToAcronym, acronymToNames, "Filter", "FLTR"
    AddMapping nameToAcronym, acronymToNames, "Final", "FINAL"
    AddMapping nameToAcronym, acronymToNames, "Flow", "FLOW"
    AddMapping nameToAcronym, acronymToNames, "Gallons", "GAL"
    AddMapping nameToAcronym, acronymToNames, "Gallons Per Minute", "GPM"
    AddMapping nameToAcronym, acronymToNames, "Heating", "HTG"
    AddMapping nameToAcronym, acronymToNames, "Heating Hot Water", "HHW"
    AddMapping nameToAcronym, acronymToNames, "Hot Water Pump", "HWP"
    AddMapping nameToAcronym, acronymToNames, "High", "HI"
    AddMapping nameToAcronym, acronymToNames, "Hot Aisle", "HA"
    AddMapping nameToAcronym, acronymToNames, "Hot Aisle Differential Pressure", "HADP"
    AddMapping nameToAcronym, acronymToNames, "Hot Aisle Temperature", "HAT"
    AddMapping nameToAcronym, acronymToNames, "Humidity", "H"
    AddMapping nameToAcronym, acronymToNames, "Industrial", "I"
    AddMapping nameToAcronym, acronymToNames, "Industrial Water", "IW"
    AddMapping nameToAcronym, acronymToNames, "Industrial Water Pump", "IWP"
    AddMapping nameToAcronym, acronymToNames, "Isolation", "ISO"
    AddMapping nameToAcronym, acronymToNames, "Leak", "LEAK"
    AddMapping nameToAcronym, acronymToNames, "Limit", "LIM"
    AddMapping nameToAcronym, acronymToNames, "Low", "LO"
    AddMapping nameToAcronym, acronymToNames, "Maximum", "MAX"
    AddMapping nameToAcronym, acronymToNames, "Measuring", "M"
    AddMapping nameToAcronym, acronymToNames, "Metric", "METRIC"
    AddMapping nameToAcronym, acronymToNames, "Minimum", "MIN"
    AddMapping nameToAcronym, acronymToNames, "Mixed Air Temperature", "MAT"
    AddMapping nameToAcronym, acronymToNames, "Mode", "MODE"
    AddMapping nameToAcronym, acronymToNames, "Occupancy Sensor", "OCC_SENSOR"
    AddMapping nameToAcronym, acronymToNames, "Occupied", "OCC"
    AddMapping nameToAcronym, acronymToNames, "Offset", "OFFSET"
    AddMapping nameToAcronym, acronymToNames, "Optimum", "OPT"
    AddMapping nameToAcronym, acronymToNames, "Outside Air", "OSA"
    AddMapping nameToAcronym, acronymToNames, "Outside Air Relative Humidity", "OA_RH"
    AddMapping nameToAcronym, acronymToNames, "Outside Air Temperature", "OAT"

    AddMapping nameToAcronym, acronymToNames, "Outside Dew Point", "OA_DEWPT"
    AddMapping nameToAcronym, acronymToNames, "Outside Wet Bulb Temp", "OA_WB"
    AddMapping nameToAcronym, acronymToNames, "Open", "OPEN"
    AddMapping nameToAcronym, acronymToNames, "Override", "OVRD"
    AddMapping nameToAcronym, acronymToNames, "Position", "POS"
    AddMapping nameToAcronym, acronymToNames, "Power", "KW"
    AddMapping nameToAcronym, acronymToNames, "Pre", "PRE"
    AddMapping nameToAcronym, acronymToNames, "Pressure Switch", "PS"
    AddMapping nameToAcronym, acronymToNames, "Primary", "P"
    AddMapping nameToAcronym, acronymToNames, "Primary Chilled Water Pump", "PCHWP"
    AddMapping nameToAcronym, acronymToNames, "Primary Chilled Water Return Temp", "PCHWRT"
    AddMapping nameToAcronym, acronymToNames, "Primary Chilled Water Supply Temp", "PCHWST"
    AddMapping nameToAcronym, acronymToNames, "Primary Heating Water Pump", "PHWP"
    AddMapping nameToAcronym, acronymToNames, "Primary Heating Water Return Temp", "PHWRT"
    AddMapping nameToAcronym, acronymToNames, "Primary Heating Water Supply Temp", "PHWST"
    AddMapping nameToAcronym, acronymToNames, "Process", "P"
    AddMapping nameToAcronym, acronymToNames, "Process Water", "PW"
    AddMapping nameToAcronym, acronymToNames, "Process Water Pump", "PWP"
    AddMapping nameToAcronym, acronymToNames, "Pump", "P"
    AddMapping nameToAcronym, acronymToNames, "Recirculation", "RECIR"
    AddMapping nameToAcronym, acronymToNames, "Relative Humidity", "RH"
    AddMapping nameToAcronym, acronymToNames, "Relative Humidity Outside", "OA_RH"
    AddMapping nameToAcronym, acronymToNames, "Relay", "RLY"
    AddMapping nameToAcronym, acronymToNames, "Relief Fan", "RF"
    AddMapping nameToAcronym, acronymToNames, "Request", "REQ"
    AddMapping nameToAcronym, acronymToNames, "Reset", "RESET"
    AddMapping nameToAcronym, acronymToNames, "Return", "R"
    AddMapping nameToAcronym, acronymToNames, "Return Air Temperature", "RAT"
    AddMapping nameToAcronym, acronymToNames, "Return Fan", "RF"
    AddMapping nameToAcronym, acronymToNames, "Reverse", "REV"
    AddMapping nameToAcronym, acronymToNames, "Room", "ZONE"
    AddMapping nameToAcronym, acronymToNames, "Room Dew Point", "RM_DEWPT"
    AddMapping nameToAcronym, acronymToNames, "Room Differential Pressure", "RM_DP"
    AddMapping nameToAcronym, acronymToNames, "Room Humidity", "RM_RH"
    AddMapping nameToAcronym, acronymToNames, "Room Temperature", "RMT"
    AddMapping nameToAcronym, acronymToNames, "Runtime", "RUNTIME"
    AddMapping nameToAcronym, acronymToNames, "Schedule", "SCHD"
    AddMapping nameToAcronym, acronymToNames, "Secondary", "S"
    AddMapping nameToAcronym, acronymToNames, "Secondary Chilled Water Pump", "SCHWP"

    AddMapping nameToAcronym, acronymToNames, "Secondary Chilled Water Return Temp", "SCHWRT"
    AddMapping nameToAcronym, acronymToNames, "Secondary Chilled Water Supply Temp", "SCHWST"
    AddMapping nameToAcronym, acronymToNames, "Secondary Heating Water Pump", "SHWP"
    AddMapping nameToAcronym, acronymToNames, "Secondary Heating Water Return Temp", "SHWRT"
    AddMapping nameToAcronym, acronymToNames, "Secondary Heating Water Supply Temp", "SHWST"
    AddMapping nameToAcronym, acronymToNames, "Secondary Hot Water Pump", "SHWP"
    AddMapping nameToAcronym, acronymToNames, "Set Point", "SP"
    AddMapping nameToAcronym, acronymToNames, "Smoke", "SMOKE"
    AddMapping nameToAcronym, acronymToNames, "Space", "ZONE"
    AddMapping nameToAcronym, acronymToNames, "Speed", "SPD"
    AddMapping nameToAcronym, acronymToNames, "Standard", "STD"
    AddMapping nameToAcronym, acronymToNames, "Start", "STRT"
    AddMapping nameToAcronym, acronymToNames, "Start/Stop Command", "SS"
    AddMapping nameToAcronym, acronymToNames, "Station", "S"
    AddMapping nameToAcronym, acronymToNames, "Status", "STS"
    AddMapping nameToAcronym, acronymToNames, "Supply", "S"
    AddMapping nameToAcronym, acronymToNames, "Supply Air Temperature", "SAT"
    AddMapping nameToAcronym, acronymToNames, "Supply Fan", "SF"
    AddMapping nameToAcronym, acronymToNames, "Switch", "SW"
    AddMapping nameToAcronym, acronymToNames, "Temperature", "T"
    AddMapping nameToAcronym, acronymToNames, "Unoccupied", "UNOCC"
    AddMapping nameToAcronym, acronymToNames, "Valve", "VLV"
    AddMapping nameToAcronym, acronymToNames, "Vibration", "VIBRATION"
    AddMapping nameToAcronym, acronymToNames, "Voltage", "VOLTS"
    AddMapping nameToAcronym, acronymToNames, "Water", "W"
    AddMapping nameToAcronym, acronymToNames, "Water Flow", "GPM"
    AddMapping nameToAcronym, acronymToNames, "Water Used", "GAL"
    AddMapping nameToAcronym, acronymToNames, "Wet Bulb Outside Temp", "OA_WB"
    AddMapping nameToAcronym, acronymToNames, "Window Contact", "WINDOW"
    AddMapping nameToAcronym, acronymToNames, "Zone", "ZONE"

End Sub

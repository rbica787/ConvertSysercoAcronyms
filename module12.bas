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
            prompt:= _
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
    Dim acronymToName As Object

    Dim searchKeys As Variant
    Dim cell As Range

    Dim originalText As String
    Dim revisedText As String
    Dim foundConversion As Boolean

    Dim changedCells As Long
    Dim matchedCells As Long
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
    Set acronymToName = CreateObject("Scripting.Dictionary")

    nameToAcronym.compareMode = vbTextCompare
    acronymToName.compareMode = vbTextCompare

    BuildSysercoMappings nameToAcronym, acronymToName
    LoadCustomMappings nameToAcronym, acronymToName

    If conversionMode = 1 Then
        searchKeys = GetSortedDictionaryKeys(acronymToName)
    Else
        searchKeys = GetSortedDictionaryKeys(nameToAcronym)
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo ConversionError

    For Each cell In Selection.Cells

        If cell.HasFormula Then

            skippedFormulaCells = skippedFormulaCells + 1

        ElseIf Not IsError(cell.value) Then

            If Len(CStr(cell.Value2)) > 0 Then

                originalText = CStr(cell.Value2)
                revisedText = originalText
                foundConversion = False

                If conversionMode = 1 Then

                    revisedText = ConvertAcronymsInText( _
                        sourceText:=revisedText, _
                        sortedAcronyms:=searchKeys, _
                        acronymToName:=acronymToName, _
                        foundConversion:=foundConversion)

                    revisedText = NormalizeNameOutput(revisedText)

                Else

                    revisedText = ConvertNamesInText( _
                        sourceText:=revisedText, _
                        sortedNames:=searchKeys, _
                        nameToAcronym:=nameToAcronym, _
                        foundConversion:=foundConversion)

                    If foundConversion Then
                        revisedText = NormalizeAcronymOutput(revisedText)
                    End If

                End If

                If foundConversion Then

                    matchedCells = matchedCells + 1

                    If StrComp( _
                        revisedText, _
                        originalText, _
                        vbBinaryCompare) <> 0 Then

                        cell.value = revisedText
                        changedCells = changedCells + 1

                    End If

                End If

            End If

        End If

    Next cell

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    If matchedCells = 0 Then

        If conversionMode = 1 Then

            MsgBox _
                "No recognized Syserco acronyms were found in the selected non-formula cells.", _
                vbInformation, _
                "No Acronyms Found"

        Else

            MsgBox _
                "No recognized Syserco names were found in the selected non-formula cells.", _
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

        If matchedCells > changedCells Then

            completionMessage = completionMessage & vbCrLf & _
                                (matchedCells - changedCells) & _
                                " additional selected cell"

            If matchedCells - changedCells <> 1 Then
                completionMessage = completionMessage & "s were"
            Else
                completionMessage = completionMessage & " was"
            End If

            completionMessage = completionMessage & _
                                " already in the requested format."

        End If

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

    MsgBox _
        "The conversion could not be completed." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, _
        "Conversion Error"

End Sub


'============================================================
' CONVERT ACRONYMS TO NAMES
'
' Each acronym has one canonical output name.
'
' Examples:
'   EF       -> EXHAUST FAN
'   RF       -> RETURN FAN
'   OA_WB    -> OUTSIDE AIR WET BULB TEMPERATURE
'   WTR_FLOW -> WATER FLOW
'============================================================
Private Function ConvertAcronymsInText( _
    ByVal sourceText As String, _
    ByVal sortedAcronyms As Variant, _
    ByVal acronymToName As Object, _
    ByRef foundConversion As Boolean) As String

    Dim i As Long
    Dim acronym As String
    Dim replacementName As String
    Dim resultText As String

    resultText = sourceText

    For i = LBound(sortedAcronyms) To UBound(sortedAcronyms)

        acronym = CStr(sortedAcronyms(i))

        If ContainsWholeTerm( _
            sourceText:=resultText, _
            searchTerm:=acronym, _
            flexibleNameSeparators:=False) Then

            replacementName = CStr(acronymToName(acronym))

            resultText = ReplaceWholeTerm( _
                sourceText:=resultText, _
                searchTerm:=acronym, _
                replacementText:=replacementName, _
                flexibleNameSeparators:=False)

            foundConversion = True

        End If

    Next i

    ConvertAcronymsInText = resultText

End Function


'============================================================
' CONVERT NAMES TO ACRONYMS
'
' Alternate names may convert to a canonical acronym.
'
' Example:
'   RELIEF FAN -> EF
'
' However:
'   EF -> EXHAUST FAN
'============================================================
Private Function ConvertNamesInText( _
    ByVal sourceText As String, _
    ByVal sortedNames As Variant, _
    ByVal nameToAcronym As Object, _
    ByRef foundConversion As Boolean) As String

    Dim i As Long
    Dim fullName As String
    Dim acronym As String
    Dim resultText As String

    resultText = sourceText

    For i = LBound(sortedNames) To UBound(sortedNames)

        fullName = CStr(sortedNames(i))
        acronym = CStr(nameToAcronym(fullName))

        If ContainsWholeTerm( _
            sourceText:=resultText, _
            searchTerm:=fullName, _
            flexibleNameSeparators:=True) Then

            resultText = ReplaceWholeTerm( _
                sourceText:=resultText, _
                searchTerm:=fullName, _
                replacementText:=acronym, _
                flexibleNameSeparators:=True)

            foundConversion = True

        End If

    Next i

    ConvertNamesInText = resultText

End Function


'============================================================
' FORMAT ACRONYM-TO-NAME OUTPUT
'
' Rules:
'   - EVERY underscore becomes a space.
'   - Multiple spaces collapse to one.
'   - Final output is uppercase.
'
' Examples:
'   OA_DEWPT
'      -> OUTSIDE DEW POINT
'
'   ROOM_DEW_POINT
'      -> ROOM DEW POINT
'
'   EF_STATUS
'      -> EXHAUST FAN STATUS
'============================================================
Private Function NormalizeNameOutput( _
    ByVal sourceText As String) As String

    Dim regex As Object

    'Replace every underscore with a space.
    sourceText = Replace(sourceText, "_", " ")

    'Convert tabs and line breaks to spaces.
    sourceText = Replace(sourceText, vbTab, " ")
    sourceText = Replace(sourceText, vbCr, " ")
    sourceText = Replace(sourceText, vbLf, " ")

    'Collapse repeated whitespace into a single space.
    Set regex = CreateObject("VBScript.RegExp")

    regex.Global = True
    regex.Pattern = "\s+"

    sourceText = regex.Replace(sourceText, " ")

    'Trim leading/trailing spaces.
    sourceText = Trim$(sourceText)

    'Always return uppercase.
    NormalizeNameOutput = UCase$(sourceText)

End Function


'============================================================
' FORMAT NAME-TO-ACRONYM OUTPUT
'
' Rules:
'   - Spaces between terms become underscores.
'   - Repeated underscores are reduced to one.
'   - Final output is uppercase.
'
' Example:
'   Supply Air Temperature Alarm
'
' Becomes:
'   SAT_ALM
'============================================================
Private Function NormalizeAcronymOutput( _
    ByVal sourceText As String) As String

    Dim regex As Object
    Dim resultText As String

    resultText = sourceText

    resultText = Replace(resultText, vbTab, " ")
    resultText = Replace(resultText, vbCr, " ")
    resultText = Replace(resultText, vbLf, " ")

    Set regex = CreateObject("VBScript.RegExp")

    regex.Global = True
    regex.IgnoreCase = False

    regex.Pattern = "\s+"
    resultText = regex.Replace(resultText, "_")

    regex.Pattern = "_+"
    resultText = regex.Replace(resultText, "_")

    Do While Len(resultText) > 0 And Left$(resultText, 1) = "_"
        resultText = Mid$(resultText, 2)
    Loop

    Do While Len(resultText) > 0 And Right$(resultText, 1) = "_"
        resultText = Left$(resultText, Len(resultText) - 1)
    Loop

    NormalizeAcronymOutput = UCase$(resultText)

End Function


'============================================================
' ADD A NEW CUSTOM NAME / ACRONYM CONVERSION
'
' Custom mappings must remain unique:
'
'   - A name cannot use two different acronyms.
'   - An acronym cannot convert back to two different names.
'============================================================
Private Sub AddNewSysercoConversion()

    Dim fullNameInput As Variant
    Dim acronymInput As Variant

    Dim fullName As String
    Dim acronym As String

    Dim nameToAcronym As Object
    Dim acronymToName As Object

    Dim customSheet As Worksheet
    Dim nextRow As Long
    Dim answer As VbMsgBoxResult

    fullNameInput = Application.InputBox( _
        prompt:= _
            "Enter the full NAME for the new conversion." & vbCrLf & vbCrLf & _
            "Example: Discharge Air Temperature", _
        Title:="Add New Syserco Name", _
        Type:=2)

    If UserCancelledInputBox(fullNameInput) Then Exit Sub

    fullName = NormalizeStoredName(CStr(fullNameInput))

    If Len(fullName) = 0 Then

        MsgBox "The NAME cannot be blank.", _
               vbExclamation, _
               "Name Required"

        Exit Sub

    End If

    acronymInput = Application.InputBox( _
        prompt:= _
            "Enter the unique ACRONYM that corresponds to:" & _
            vbCrLf & vbCrLf & _
            fullName & vbCrLf & vbCrLf & _
            "Example: DAT", _
        Title:="Add New Syserco Acronym", _
        Type:=2)

    If UserCancelledInputBox(acronymInput) Then Exit Sub

    acronym = NormalizeStoredAcronym(CStr(acronymInput))

    If Len(acronym) = 0 Then

        MsgBox "The ACRONYM cannot be blank.", _
               vbExclamation, _
               "Acronym Required"

        Exit Sub

    End If

    Set nameToAcronym = CreateObject("Scripting.Dictionary")
    Set acronymToName = CreateObject("Scripting.Dictionary")

    nameToAcronym.compareMode = vbTextCompare
    acronymToName.compareMode = vbTextCompare

    BuildSysercoMappings nameToAcronym, acronymToName
    LoadCustomMappings nameToAcronym, acronymToName

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
                "The name """ & fullName & _
                """ is already assigned to:" & vbCrLf & vbCrLf & _
                CStr(nameToAcronym(fullName)) & vbCrLf & vbCrLf & _
                "Each name can only have one acronym.", _
                vbExclamation, _
                "Name Already Assigned"

        End If

        Exit Sub

    End If

    If acronymToName.Exists(acronym) Then

        MsgBox _
            "The acronym """ & acronym & _
            """ is already assigned to:" & vbCrLf & vbCrLf & _
            CStr(acronymToName(acronym)) & vbCrLf & vbCrLf & _
            "Enter a different unique acronym.", _
            vbExclamation, _
            "Acronym Already Assigned"

        Exit Sub

    End If

    answer = MsgBox( _
        "Add this new conversion?" & vbCrLf & vbCrLf & _
        "NAME: " & fullName & vbCrLf & _
        "ACRONYM: " & acronym, _
        vbYesNo + vbQuestion, _
        "Confirm New Conversion")

    If answer <> vbYes Then Exit Sub

    Set customSheet = GetOrCreateCustomMappingSheet()

    nextRow = customSheet.Cells( _
        customSheet.Rows.Count, _
        1).End(xlUp).Row + 1

    If nextRow < 2 Then nextRow = 2

    customSheet.Cells(nextRow, 1).value = fullName
    customSheet.Cells(nextRow, 2).value = acronym

    customSheet.Columns("A:B").AutoFit
    customSheet.Visible = xlSheetHidden

    MsgBox _
        "The new conversion was saved successfully." & vbCrLf & vbCrLf & _
        fullName & " = " & acronym & vbCrLf & vbCrLf & _
        "It is now available in both conversion directions.", _
        vbInformation, _
        "Conversion Added"

End Sub


'============================================================
' GET OR CREATE CUSTOM CONVERSION WORKSHEET
'============================================================
Private Function GetOrCreateCustomMappingSheet() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CUSTOM_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets( _
                ThisWorkbook.Worksheets.Count))

        ws.Name = CUSTOM_SHEET_NAME

        ws.Cells(1, 1).value = "NAME"
        ws.Cells(1, 2).value = "ACRONYM"

        ws.Range("A1:B1").Font.Bold = True
        ws.Columns("A:B").AutoFit

    End If

    Set GetOrCreateCustomMappingSheet = ws

End Function


'============================================================
' LOAD CUSTOM CONVERSIONS
'
' Conflicting saved rows are ignored so they cannot override
' or corrupt the built-in unique mappings.
'============================================================
Private Sub LoadCustomMappings( _
    ByVal nameToAcronym As Object, _
    ByVal acronymToName As Object)

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

        fullName = NormalizeStoredName( _
            CStr(ws.Cells(rowNumber, 1).Value2))

        acronym = NormalizeStoredAcronym( _
            CStr(ws.Cells(rowNumber, 2).Value2))

        If Len(fullName) > 0 And Len(acronym) > 0 Then

            If Not nameToAcronym.Exists(fullName) And _
               Not acronymToName.Exists(acronym) Then

                AddCanonicalMapping _
                    nameToAcronym:=nameToAcronym, _
                    acronymToName:=acronymToName, _
                    fullName:=fullName, _
                    acronym:=acronym

            End If

        End If

    Next rowNumber

End Sub


'============================================================
' NORMALIZE A STORED NAME
'============================================================
Private Function NormalizeStoredName( _
    ByVal sourceText As String) As String

    Dim regex As Object
    Dim resultText As String

    resultText = sourceText

    resultText = Replace(resultText, "_", " ")
    resultText = Replace(resultText, vbTab, " ")
    resultText = Replace(resultText, vbCr, " ")
    resultText = Replace(resultText, vbLf, " ")

    Set regex = CreateObject("VBScript.RegExp")

    regex.Global = True
    regex.Pattern = "\s+"

    resultText = regex.Replace(resultText, " ")

    NormalizeStoredName = UCase$(Trim$(resultText))

End Function


'============================================================
' NORMALIZE A STORED ACRONYM
'============================================================
Private Function NormalizeStoredAcronym( _
    ByVal sourceText As String) As String

    Dim regex As Object
    Dim resultText As String

    resultText = Trim$(sourceText)

    resultText = Replace(resultText, vbTab, " ")
    resultText = Replace(resultText, vbCr, " ")
    resultText = Replace(resultText, vbLf, " ")

    Set regex = CreateObject("VBScript.RegExp")

    regex.Global = True

    regex.Pattern = "\s+"
    resultText = regex.Replace(resultText, "_")

    regex.Pattern = "_+"
    resultText = regex.Replace(resultText, "_")

    Do While Len(resultText) > 0 And Left$(resultText, 1) = "_"
        resultText = Mid$(resultText, 2)
    Loop

    Do While Len(resultText) > 0 And Right$(resultText, 1) = "_"
        resultText = Left$(resultText, Len(resultText) - 1)
    Loop

    NormalizeStoredAcronym = UCase$(resultText)

End Function


'============================================================
' DETERMINE WHETHER APPLICATION.INPUTBOX WAS CANCELLED
'============================================================
Private Function UserCancelledInputBox( _
    ByVal inputResult As Variant) As Boolean

    If VarType(inputResult) = vbBoolean Then
        UserCancelledInputBox = (inputResult = False)
    Else
        UserCancelledInputBox = False
    End If

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
    regex.Multiline = True

    searchPattern = BuildSearchPattern( _
        searchTerm:=searchTerm, _
        flexibleNameSeparators:=flexibleNameSeparators)

    regex.Pattern = _
        "(^|[^A-Za-z0-9])(" & _
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
    regex.Multiline = True

    searchPattern = BuildSearchPattern( _
        searchTerm:=searchTerm, _
        flexibleNameSeparators:=flexibleNameSeparators)

    regex.Pattern = _
        "(^|[^A-Za-z0-9])(" & _
        searchPattern & _
        ")(?=$|[^A-Za-z0-9])"

    ReplaceWholeTerm = regex.Replace( _
        sourceText, _
        "$1" & replacementText)

End Function


'============================================================
' BUILD REGULAR-EXPRESSION SEARCH PATTERN
'
' Names may be written with:
'   - spaces
'   - underscores
'   - hyphens
'
' Example matches:
'   SUPPLY AIR TEMPERATURE
'   SUPPLY_AIR_TEMPERATURE
'   SUPPLY-AIR-TEMPERATURE
'============================================================
Private Function BuildSearchPattern( _
    ByVal searchTerm As String, _
    ByVal flexibleNameSeparators As Boolean) As String

    Dim escapedTerm As String

    escapedTerm = EscapeRegexText(searchTerm)

    If flexibleNameSeparators Then
        escapedTerm = Replace( _
            escapedTerm, _
            " ", _
            "[ _-]+")
    End If

    BuildSearchPattern = escapedTerm

End Function


'============================================================
' ESCAPE REGULAR-EXPRESSION SPECIAL CHARACTERS
'============================================================
Private Function EscapeRegexText( _
    ByVal value As String) As String

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
' SORT KEYS FROM LONGEST TO SHORTEST
'
' This prevents a shorter name such as:
'   WATER
'
' from being processed before:
'   PROCESS WATER PUMP
'============================================================
Private Function GetSortedDictionaryKeys( _
    ByVal dictionary As Object) As Variant

    Dim keys As Variant
    Dim i As Long
    Dim j As Long
    Dim temporaryValue As Variant

    keys = dictionary.keys

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
' ADD A CANONICAL TWO-WAY MAPPING
'
' Example:
'   EXHAUST FAN -> EF
'   EF -> EXHAUST FAN
'============================================================
Private Sub AddCanonicalMapping( _
    ByVal nameToAcronym As Object, _
    ByVal acronymToName As Object, _
    ByVal fullName As String, _
    ByVal acronym As String)

    fullName = NormalizeStoredName(fullName)
    acronym = NormalizeStoredAcronym(acronym)

    If Len(fullName) = 0 Or Len(acronym) = 0 Then Exit Sub

    If Not nameToAcronym.Exists(fullName) Then
        nameToAcronym.Add fullName, acronym
    End If

    If Not acronymToName.Exists(acronym) Then
        acronymToName.Add acronym, fullName
    End If

End Sub


'============================================================
' ADD A NAME-TO-ACRONYM ALIAS
'
' The alias converts to the acronym, but the acronym retains
' its existing canonical reverse name.
'
' Example:
'   RELIEF FAN -> EF
'
' EF still converts to:
'   EXHAUST FAN
'============================================================
Private Sub AddNameAlias( _
    ByVal nameToAcronym As Object, _
    ByVal fullName As String, _
    ByVal acronym As String)

    fullName = NormalizeStoredName(fullName)
    acronym = NormalizeStoredAcronym(acronym)

    If Len(fullName) = 0 Or Len(acronym) = 0 Then Exit Sub

    If Not nameToAcronym.Exists(fullName) Then
        nameToAcronym.Add fullName, acronym
    End If

End Sub


'============================================================
' BUILT-IN SYSERCO NAME / ACRONYM LOOKUP TABLE
'
' Characteristics incorporated:
'
'   BOILER = BLR
'   BUILDING = BLDG
'   RETURN FAN = RF
'   EXHAUST FAN = EF
'   RELIEF FAN -> EF alias
'   PRIMARY = PRMRY
'   PUMP = PMP
'   WATER = WTR
'   WATER FLOW = WTR_FLOW
'
' No standalone conversion for:
'
'   PROCESS
'   SECONDARY
'   STATION
'   SUPPLY
'   ROOM
'   SPACE
'   ZONE
'   WATER USED
'
' Canonical wording:
'
'   OUTSIDE DEW POINT
'   ROOM DEW POINT
'   ROOM DIFFERENTIAL PRESSURE
'   OUTSIDE AIR RELATIVE HUMIDITY
'   OUTSIDE AIR WET BULB TEMPERATURE
'============================================================
Private Sub BuildSysercoMappings( _
    ByVal nameToAcronym As Object, _
    ByVal acronymToName As Object)

    AddCanonicalMapping nameToAcronym, acronymToName, "Actuator", "ACTR"
    AddCanonicalMapping nameToAcronym, acronymToName, "Airflow", "AF"
    AddCanonicalMapping nameToAcronym, acronymToName, "Airflow Measuring Station", "AFMS"
    AddCanonicalMapping nameToAcronym, acronymToName, "Alarm", "ALM"
    AddCanonicalMapping nameToAcronym, acronymToName, "Bias", "BIAS"
    AddCanonicalMapping nameToAcronym, acronymToName, "Boiler", "BLR"
    AddCanonicalMapping nameToAcronym, acronymToName, "Building", "BLDG"
    AddCanonicalMapping nameToAcronym, acronymToName, "Building Differential Pressure", "BDP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Bypass", "BYP"

    AddCanonicalMapping nameToAcronym, acronymToName, "Chilled Water", "CHW"
    AddCanonicalMapping nameToAcronym, acronymToName, "Chilled Water Pump", "CHWP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Chiller", "CH"
    AddCanonicalMapping nameToAcronym, acronymToName, "CO2", "CO2"
    AddCanonicalMapping nameToAcronym, acronymToName, "Cold Aisle", "CA"
    AddCanonicalMapping nameToAcronym, acronymToName, "Cold Aisle Differential Pressure", "CADP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Cold Aisle Temperature", "CAT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Command", "CMD"
    AddCanonicalMapping nameToAcronym, acronymToName, "Condenser Water", "CW"
    AddCanonicalMapping nameToAcronym, acronymToName, "Condenser Water Pump", "CWP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Cooling", "CLG"
    AddCanonicalMapping nameToAcronym, acronymToName, "Cooling Tower", "CT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Cooling Tower Pump", "CTP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Close", "CLOSE"
    AddCanonicalMapping nameToAcronym, acronymToName, "Current", "CURR"

    AddCanonicalMapping nameToAcronym, acronymToName, "Damper", "DMPR"
    AddCanonicalMapping nameToAcronym, acronymToName, "Detector", "DET"
    AddCanonicalMapping nameToAcronym, acronymToName, "Dew Point", "DEWPT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Outside Dew Point", "OA_DEWPT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Room Dew Point", "RM_DEWPT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Differential Pressure", "DP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Differential Pressure Building", "B_DP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Room Differential Pressure", "RM_DP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Disable", "DISA"
    AddCanonicalMapping nameToAcronym, acronymToName, "Domestic Hot Water", "DHW"
    AddCanonicalMapping nameToAcronym, acronymToName, "Door Contact", "DOOR"

    AddCanonicalMapping nameToAcronym, acronymToName, "Economizer", "ECON"
    AddCanonicalMapping nameToAcronym, acronymToName, "Enable", "ENA"
    AddCanonicalMapping nameToAcronym, acronymToName, "Energy", "KWH"
    AddCanonicalMapping nameToAcronym, acronymToName, "Enthalpy", "ENTH"
    AddCanonicalMapping nameToAcronym, acronymToName, "Exhaust", "EXH"
    AddCanonicalMapping nameToAcronym, acronymToName, "Exhaust Fan", "EF"
    AddNameAlias nameToAcronym, "Relief Fan", "EF"

    AddCanonicalMapping nameToAcronym, acronymToName, "Fault", "FLT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Feedback", "FDBK"
    AddCanonicalMapping nameToAcronym, acronymToName, "Filter", "FLTR"
    AddCanonicalMapping nameToAcronym, acronymToName, "Final", "FINAL"
    AddCanonicalMapping nameToAcronym, acronymToName, "Flow", "FLOW"

    AddCanonicalMapping nameToAcronym, acronymToName, "Gallons", "GAL"
    AddCanonicalMapping nameToAcronym, acronymToName, "Gallons Per Minute", "GPM"
    AddCanonicalMapping nameToAcronym, acronymToName, "Heating", "HTG"
    AddCanonicalMapping nameToAcronym, acronymToName, "Heating Hot Water", "HHW"
    AddCanonicalMapping nameToAcronym, acronymToName, "Hot Water Pump", "HWP"
    AddCanonicalMapping nameToAcronym, acronymToName, "High", "HI"
    AddCanonicalMapping nameToAcronym, acronymToName, "Hot Aisle", "HA"
    AddCanonicalMapping nameToAcronym, acronymToName, "Hot Aisle Differential Pressure", "HADP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Hot Aisle Temperature", "HAT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Humidity", "HUM"

    AddCanonicalMapping nameToAcronym, acronymToName, "Industrial", "I"
    AddCanonicalMapping nameToAcronym, acronymToName, "Industrial Water", "IW"
    AddCanonicalMapping nameToAcronym, acronymToName, "Industrial Water Pump", "IWP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Isolation", "ISO"
    AddCanonicalMapping nameToAcronym, acronymToName, "Leak", "LEAK"
    AddCanonicalMapping nameToAcronym, acronymToName, "Limit", "LIM"
    AddCanonicalMapping nameToAcronym, acronymToName, "Low", "LO"
    AddCanonicalMapping nameToAcronym, acronymToName, "Maximum", "MAX"
    AddCanonicalMapping nameToAcronym, acronymToName, "Measuring", "M"
    AddCanonicalMapping nameToAcronym, acronymToName, "Metric", "METRIC"
    AddCanonicalMapping nameToAcronym, acronymToName, "Minimum", "MIN"

    AddCanonicalMapping nameToAcronym, acronymToName, "Mixed Air Temperature", "MAT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Mode", "MODE"
    AddCanonicalMapping nameToAcronym, acronymToName, "Occupancy Sensor", "OCC_SENSOR"
    AddCanonicalMapping nameToAcronym, acronymToName, "Occupied", "OCC"
    AddCanonicalMapping nameToAcronym, acronymToName, "Offset", "OFFSET"
    AddCanonicalMapping nameToAcronym, acronymToName, "Optimum", "OPT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Outside Air", "OA"
    AddCanonicalMapping nameToAcronym, acronymToName, "Outside Air Relative Humidity", "OA_RH"
    AddCanonicalMapping nameToAcronym, acronymToName, "Outside Air Temperature", "OAT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Outside Air Wet Bulb Temperature", "OA_WB"
    AddCanonicalMapping nameToAcronym, acronymToName, "Open", "OPEN"
    AddCanonicalMapping nameToAcronym, acronymToName, "Override", "OVRD"

    AddCanonicalMapping nameToAcronym, acronymToName, "Position", "POS"
    AddCanonicalMapping nameToAcronym, acronymToName, "Power", "KW"
    AddCanonicalMapping nameToAcronym, acronymToName, "Pre", "PRE"
    AddCanonicalMapping nameToAcronym, acronymToName, "Pressure Switch", "PS"
    AddCanonicalMapping nameToAcronym, acronymToName, "Primary", "PRMRY"
    AddCanonicalMapping nameToAcronym, acronymToName, "Primary Chilled Water Pump", "PCHWP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Primary Chilled Water Return Temp", "PCHWRT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Primary Chilled Water Supply Temp", "PCHWST"
    AddCanonicalMapping nameToAcronym, acronymToName, "Primary Heating Water Pump", "PHWP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Primary Heating Water Return Temp", "PHWRT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Primary Heating Water Supply Temp", "PHWST"

    AddCanonicalMapping nameToAcronym, acronymToName, "Process Water", "PW"
    AddCanonicalMapping nameToAcronym, acronymToName, "Process Water Pump", "PWP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Pump", "PMP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Recirculation", "RECIR"
    AddCanonicalMapping nameToAcronym, acronymToName, "Relative Humidity", "RH"
    AddCanonicalMapping nameToAcronym, acronymToName, "Relay", "RLY"
    AddCanonicalMapping nameToAcronym, acronymToName, "Request", "REQ"
    AddCanonicalMapping nameToAcronym, acronymToName, "Reset", "RESET"
    AddCanonicalMapping nameToAcronym, acronymToName, "Return", "R"
    AddCanonicalMapping nameToAcronym, acronymToName, "Return Air Temperature", "RAT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Return Fan", "RF"
    AddCanonicalMapping nameToAcronym, acronymToName, "Reverse", "REV"

    AddCanonicalMapping nameToAcronym, acronymToName, "Room Humidity", "RM_RH"
    AddCanonicalMapping nameToAcronym, acronymToName, "Room Temperature", "RMT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Runtime", "RUNTIME"
    AddCanonicalMapping nameToAcronym, acronymToName, "Schedule", "SCHD"
    AddCanonicalMapping nameToAcronym, acronymToName, "Secondary Chilled Water Pump", "SCHWP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Secondary Chilled Water Return Temp", "SCHWRT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Secondary Chilled Water Supply Temp", "SCHWST"
    AddCanonicalMapping nameToAcronym, acronymToName, "Secondary Hot Water Pump", "SHWP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Secondary Heating Water Return Temp", "SHWRT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Secondary Heating Water Supply Temp", "SHWST"

    AddCanonicalMapping nameToAcronym, acronymToName, "Set Point", "SP"
    AddCanonicalMapping nameToAcronym, acronymToName, "Smoke", "SMOKE"
    AddCanonicalMapping nameToAcronym, acronymToName, "Speed", "SPD"
    AddCanonicalMapping nameToAcronym, acronymToName, "Standard", "STD"
    AddCanonicalMapping nameToAcronym, acronymToName, "Start", "STRT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Start/Stop Command", "SS"
    AddCanonicalMapping nameToAcronym, acronymToName, "Status", "STS"
    AddCanonicalMapping nameToAcronym, acronymToName, "Supply Air Temperature", "SAT"
    AddCanonicalMapping nameToAcronym, acronymToName, "Supply Fan", "SF"
    AddCanonicalMapping nameToAcronym, acronymToName, "Switch", "SW"

    AddCanonicalMapping nameToAcronym, acronymToName, "Temperature", "T"
    AddCanonicalMapping nameToAcronym, acronymToName, "Unoccupied", "UNOCC"
    AddCanonicalMapping nameToAcronym, acronymToName, "Valve", "VLV"
    AddCanonicalMapping nameToAcronym, acronymToName, "Vibration", "VIBRATION"
    AddCanonicalMapping nameToAcronym, acronymToName, "Voltage", "VOLTS"
    AddCanonicalMapping nameToAcronym, acronymToName, "Water", "WTR"
    AddCanonicalMapping nameToAcronym, acronymToName, "Water Flow", "WTR_FLOW"
    AddCanonicalMapping nameToAcronym, acronymToName, "Window Contact", "WINDOW"

End Sub




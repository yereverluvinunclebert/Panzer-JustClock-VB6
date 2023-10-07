Attribute VB_Name = "modDaylightSavings"
Option Explicit

Public Sub obtainDaylightSavings()
    Dim getDLSrules() As String
    Dim numberOfMonth As String
    Dim numberOfDay As String
    Dim getDaysIn As Integer
    
    ' read the rule list from file
    getDLSrules = getDSTRules(App.path & "\Resources\txt\DLSRules.txt")
    
    ' get the number of the month given a month name
    numberOfMonth = getNumberOfMonth("Feb")
        
    ' get the number of the day given a day name
    numberOfDay = getNumberOfDay("Sat")
    
    ' get the number of days in a given month
    getDaysIn = getDaysInMonth(numberOfMonth, 1961)
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : getDSTRules
' Author    : beededea
' Date      : 06/10/2023
' Purpose   : ["US", "Apr", "Sun>=1", "120", "60", "Oct", "lastSun", "60"]
'---------------------------------------------------------------------------------------
'
Public Function getDSTRules(path) As String()
    '
    Dim ruleList() As String
    Dim rules() As String
    Dim iFile As Integer: iFile = 0
    Dim i As Variant
    Dim lFileLen As Long
    Dim sBuffer As String
    Dim useLoop As Integer: useLoop = 0
    Dim arraySize As Integer
    
    On Error GoTo getDSTRules_Error

    If Dir$(path) = vbNullString Then Exit Function
    
    On Error GoTo ErrorHandler:
    
    iFile = FreeFile
    Open path For Binary Access Read As #iFile
    lFileLen = LOF(iFile)
    If lFileLen Then
        'Create output buffer
        sBuffer = String(lFileLen, " ")
        'Read contents of file
        Get iFile, 1, sBuffer
        'Split the file contents
        ruleList = Split(sBuffer, vbCrLf)
    End If

    ' set the rules array size to match the number of rules found
    arraySize = UBound(ruleList)
    ReDim rules(arraySize)

    ' convert the variants in ruleList to strings in rules
    For Each i In ruleList
        rules(useLoop) = CStr(i)
        useLoop = useLoop + 1
    Next i
    
ErrorHandler:
    If iFile > 0 Then Close #iFile
    
    getDSTRules = rules

    On Error GoTo 0
    Exit Function

getDSTRules_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getDSTRules of Module modDaylightSavings"
End Function


'---------------------------------------------------------------------------------------
' Procedure : getNumberOfMonth
' Author    : beededea
' Date      : 07/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function getNumberOfMonth(ByVal month As String) As Integer
    Dim monthsString As String: monthsString = vbNullString
    Dim monthArray() As String
    Dim months(11) As String
    Dim i As Variant
    Dim useLoop As Integer: useLoop = 0
    
    On Error GoTo getNumberOfMonth_Error

    monthsString = "Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11"
    monthArray = Split(monthsString, ",")
    
    For Each i In monthArray
        months(useLoop) = CStr(i)
        If InStr(months(useLoop), month) > 0 Then
            getNumberOfMonth = Val(LTrim$(Mid$(months(useLoop), 6, Len(months(useLoop))))) ' return
            Exit Function
        End If
        useLoop = useLoop + 1
    Next i

    MsgBox ("getNumberOfMonth: " & month & " is not a valid month name")
    getNumberOfMonth = 99 ' return invalid

    On Error GoTo 0
    Exit Function

getNumberOfMonth_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getNumberOfMonth of Module modDaylightSavings"

End Function

'---------------------------------------------------------------------------------------
' Procedure : getNumberOfDay
' Author    : beededea
' Date      : 07/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function getNumberOfDay(ByVal day As String) As Integer
    Dim daysString As String: daysString = vbNullString
    Dim dayArray() As String
    Dim days(6) As String
    Dim i As Variant
9    Dim useLoop As Integer: useLoop = 0
    
    On Error GoTo getNumberOfDay_Error

    daysString = "Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6"
    dayArray = Split(daysString, ",")
    
    For Each i In dayArray
        days(useLoop) = CStr(i)
        If InStr(days(useLoop), day) > 0 Then
            getNumberOfDay = Val(LTrim$(Mid$(days(useLoop), 6, Len(days(useLoop))))) ' return
            Exit Function
        End If
        useLoop = useLoop + 1
    Next i

    MsgBox ("getNumberOfDay: " & day & " is not a valid day name")
    getNumberOfDay = 99 ' return invalid

    On Error GoTo 0
    Exit Function

getNumberOfDay_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getNumberOfDay of Module modDaylightSavings"

End Function



'---------------------------------------------------------------------------------------
' Procedure : getDaysInMonth
' Author    : beededea
' Date      : 07/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function getDaysInMonth(ByVal month As Integer, ByVal year As Integer) As Integer
    Dim monthDaysString As String: monthDaysString = vbNullString
    Dim monthDaysArray() As String
    Dim useLoop As Integer: useLoop = 0
    
    On Error GoTo getmonthsIn_Error
    
    If month > 11 Then
        MsgBox ("getDaysInMonth: " & month & " is not a valid month number")
        getDaysInMonth = 99 ' return invalid
        Exit Function
    End If

    monthDaysString = "31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31"
    monthDaysArray = Split(monthDaysString, ",")
    
    If month <> 1 Then ' all except Feb
        getDaysInMonth = Val(monthDaysArray(month)) ' return
        Exit Function
    End If
    
    If year Mod 4 <> 0 Then
        getDaysInMonth = 28 ' return
        Exit Function
    End If
    
    If year Mod 400 <> 0 Then
        getDaysInMonth = 29 ' return
        Exit Function
    End If
    
    If year Mod 100 <> 0 Then
        getDaysInMonth = 28 ' return
        Exit Function
    End If

    getDaysInMonth = 29 ' return

    On Error GoTo 0
    Exit Function

getmonthsIn_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getmonthsIn of Module modmonthlightSavings"

End Function

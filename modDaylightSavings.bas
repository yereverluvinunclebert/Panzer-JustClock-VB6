Attribute VB_Name = "modDaylightSavings"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : obtainDaylightSavings
' Author    : beededea
' Date      : 07/10/2023
' Purpose   : some test code
'---------------------------------------------------------------------------------------
'
Public Sub obtainDaylightSavings()

    Dim getDLSrules() As String
    Dim numberOfMonth As String: numberOfMonth = vbNullString
    Dim numberOfDay As String: numberOfDay = vbNullString
    Dim getDaysIn As Integer: getDaysIn = 0
    Dim dateOfFirst As Integer: dateOfFirst = 0
    
    On Error GoTo obtainDaylightSavings_Error

    ' read the rule list from file
    getDLSrules = getDSTRules(App.path & "\Resources\txt\DLSRules.txt")
    
    ' get the number of the month given a month name
    numberOfMonth = getNumberOfMonth("Feb")
        
    ' get the number of the day given a day name
    numberOfDay = getNumberOfDay("Sat")
    
    ' get the number of days in a given month
    getDaysIn = getDaysInMonth(numberOfMonth, 1961)
    
    ' get Date (1..31) Of First dayName (Sun..Sat) after date (1..31) of monthName (Jan..Dec) of year (2004..)
    dateOfFirst = getDateOfFirst("Sun", 15, "Feb", 1961)

    On Error GoTo 0
    Exit Sub

obtainDaylightSavings_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainDaylightSavings of Module modDaylightSavings"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : getDSTRules
' Author    : beededea
' Date      : 06/10/2023
' Purpose   : read the rule list from file
' ["US", "Apr", "Sun>=1", "120", "60", "Oct", "lastSun", "60"]
'---------------------------------------------------------------------------------------
'
Public Function getDSTRules(ByVal path As String) As String()
    
    Dim ruleList() As String
    Dim rules() As String
    Dim iFile As Integer: iFile = 0
    Dim i As Variant
    Dim lFileLen As Long
    Dim sBuffer As String
    Dim useLoop As Integer: useLoop = 0
    Dim arraySize As Integer
    
    On Error GoTo getDSTRules_Error

    If Dir$(path) = vbNullString Then
        Exit Function
    End If
    
    On Error GoTo ErrorHandler:
    
    iFile = FreeFile
    Open path For Binary Access Read As #iFile
    lFileLen = LOF(iFile)
    If lFileLen Then
        'Create output buffer
        sBuffer = String(lFileLen, " ")
        'Read contents of file
        Get iFile, 1, sBuffer
        'Split the file contents into an array
        ruleList = Split(sBuffer, vbCrLf)
    End If

    ' set the output rules array size to match the number of rules found
    arraySize = UBound(ruleList)
    ReDim rules(arraySize)

    ' convert the variants in ruleList to strings in output rules
    For Each i In ruleList
        rules(useLoop) = CStr(i)
        useLoop = useLoop + 1
    Next i
    
ErrorHandler:
    If iFile > 0 Then Close #iFile
    
    getDSTRules = rules ' return

    On Error GoTo 0
    Exit Function

getDSTRules_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getDSTRules of Module modDaylightSavings"
End Function


'---------------------------------------------------------------------------------------
' Procedure : getNumberOfMonth
' Author    : beededea
' Date      : 07/10/2023
' Purpose   : get the number of the month given a month name
'---------------------------------------------------------------------------------------
'
Public Function getNumberOfMonth(ByVal thisMonth As String) As Integer
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
        If InStr(months(useLoop), thisMonth) > 0 Then
            getNumberOfMonth = Val(LTrim$(Mid$(months(useLoop), 6, Len(months(useLoop))))) ' return
            Exit Function
        End If
        useLoop = useLoop + 1
    Next i

    MsgBox ("getNumberOfMonth: " & thisMonth & " is not a valid month name")
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
' Purpose   : get the number of the day given a day name
'---------------------------------------------------------------------------------------
'
Public Function getNumberOfDay(ByVal thisDay As String) As Integer
    Dim daysString As String: daysString = vbNullString
    Dim dayArray() As String
    Dim days(6) As String
    Dim i As Variant
    Dim useLoop As Integer: useLoop = 0
    
    On Error GoTo getNumberOfDay_Error

    daysString = "Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6"
    dayArray = Split(daysString, ",")
    
    For Each i In dayArray
        days(useLoop) = CStr(i)
        If InStr(days(useLoop), thisDay) > 0 Then
            getNumberOfDay = Val(LTrim$(Mid$(days(useLoop), 6, Len(days(useLoop))))) ' return
            Exit Function
        End If
        useLoop = useLoop + 1
    Next i

    MsgBox ("getNumberOfDay: " & thisDay & " is not a valid day name")
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
' Purpose   : get the number of days in a given month
'---------------------------------------------------------------------------------------
'
Public Function getDaysInMonth(ByVal thisMonth As Integer, ByVal thisYear As Integer) As Integer
    Dim monthDaysString As String: monthDaysString = vbNullString
    Dim monthDaysArray() As String
    Dim useLoop As Integer: useLoop = 0
    
    On Error GoTo getmonthsIn_Error
    
    If thisMonth < 0 And thisMonth > 11 Then
        MsgBox ("getDaysInMonth: " & thisMonth & " is not a valid month number")
        getDaysInMonth = 99 ' return invalid
        Exit Function
    End If

    monthDaysString = "31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31"
    monthDaysArray = Split(monthDaysString, ",")
    
    If thisMonth <> 1 Then ' all except Feb
        getDaysInMonth = Val(monthDaysArray(thisMonth)) ' return
        Exit Function
    End If
    
    If thisYear Mod 4 <> 0 Then
        getDaysInMonth = 28 ' return
        Exit Function
    End If
    
    If thisYear Mod 400 <> 0 Then
        getDaysInMonth = 29 ' return
        Exit Function
    End If
    
    If thisYear Mod 100 <> 0 Then
        getDaysInMonth = 28 ' return
        Exit Function
    End If

    getDaysInMonth = 29 ' return

    On Error GoTo 0
    Exit Function

getmonthsIn_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getmonthsIn of Module modmonthlightSavings"

End Function
    

'---------------------------------------------------------------------------------------
' Procedure : getDateOfFirst
' Author    : beededea
' Date      : 07/10/2023
' Purpose   :  get Date (1..31) Of First dayName (Sun..Sat) after date (1..31) of monthName (Jan..Dec) of year (2004..)
'              dayName:     Sun, Mon, Tue, Wed, Thu, Fr, Sat
'              monthName:   Jan, Feb, etc.
'---------------------------------------------------------------------------------------
'
Public Function getDateOfFirst(ByVal dayName As String, ByVal thisDayNumber As Integer, ByVal monthName As String, ByVal thisYear As Integer) As Integer
'
    Dim tDay As Integer
    Dim tMonth As Integer
    Dim last As Integer
    Dim d As Date
    Dim lastDay As Date

    On Error GoTo getDateOfFirst_Error

    tDay = getNumberOfDay(dayName)
    tMonth = getNumberOfMonth(monthName)
    last = thisDayNumber + 6
    d = CDate(last & "/" & tMonth & "/" & thisYear)
    
    lastDay = DateSerial(thisYear, tMonth, last)
    If IsDate(lastDay) Then
        getDateOfFirst = last - (lastDay - tDay + 7) Mod 7 'return
    Else
        getDateOfFirst = 99 ' return invalid
    End If

    On Error GoTo 0
    Exit Function

getDateOfFirst_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getDateOfFirst of Module modDaylightSavings"
End Function

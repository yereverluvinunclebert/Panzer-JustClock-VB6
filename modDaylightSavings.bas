Attribute VB_Name = "modDaylightSavings"
'---------------------------------------------------------------------------------------
' Module    : modDaylightSavings
' Author    : beededea
' Date      : 07/10/2023
' Purpose   : coverting some .js routines to VB6, converting manually, will look for some
'             native vb6 methods of doing the same and use those to test the results.
'---------------------------------------------------------------------------------------

Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : obtainDaylightSavings
' Author    : beededea
' Date      : 07/10/2023
' Purpose   : some test code
'---------------------------------------------------------------------------------------
'
Public Sub obtainDaylightSavings()

    Dim DLSrules() As String
    Dim numberOfMonth As String: numberOfMonth = vbNullString
    Dim numberOfDay As String: numberOfDay = vbNullString
    Dim getDaysIn As Integer: getDaysIn = 0
    Dim dateOfFirst As Integer: dateOfFirst = 0
    Dim tzDelta1 As Long: tzDelta1 = 0
    
    On Error GoTo obtainDaylightSavings_Error

    ' Exit Sub
    

    ' read the rule list from file
    DLSrules = getDLSrules(App.path & "\Resources\txt\DLSRules.txt")
    
    Call updateDLS
'
'    ' get the number of the month given a month name
'    numberOfMonth = getNumberOfMonth("Feb")
'
'    ' get the number of the day given a day name
'    numberOfDay = getNumberOfDay("Sat")
'
'    ' get the number of days in a given month
'    getDaysIn = getDaysInMonth(numberOfMonth, 1961)
'
'    ' get Date (1..31) Of First dayName (Sun..Sat) after date (1..31) of monthName (Jan..Dec) of year (2004..)
'    dateOfFirst = getDateOfFirst("Sun", 1, "Sep", 1961)
'
'    tzDelta1 = theDLSdelta(DLSrules(), "EU", 0)

    On Error GoTo 0
    Exit Sub

obtainDaylightSavings_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainDaylightSavings of Module modDaylightSavings"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : updateDLS
' Author    : beededea
' Date      : 10/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function updateDLS()

    Dim dlsRule As Variant
    Dim remoteGMTOffset1 As Long
    Dim remoteGMTOffset2 As Long
    
    On Error GoTo updateDLS_Error


    remoteGMTOffset1 = getRemoteOffset("GMT - 00:00 Greenwich Mean Time: London")
    remoteGMTOffset2 = getRemoteOffset("GMT - 00:00 Greenwich Mean Time: London")
'
'
'    dlsRule = preferences.mainDLSPref.value.split(/\s/)[0];
'
'    Print ("remoteGMTOffset1 " + remoteGMTOffset1);
'    Print ("dlsRule >> " + dlsRule + "<< ");
'
'    tzDelta1 = theDLSdelta(dlsRule, remoteGMTOffset1);
'
'    dlsRule = preferences.secyDLSPref.value.split(/\s/)[0];
'    tzDelta2 = theDLSdelta(dlsRule, remoteGMTOffset2);

    On Error GoTo 0
    Exit Function

updateDLS_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure updateDLS of Module modDaylightSavings"
End Function
'---------------------------------------------------------------------------------------
' Procedure : getRemoteOffset
' Author    : beededea
' Date      : 10/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
 Private Function getRemoteOffset(ByVal entry As String) As Long

    Dim found As Boolean
    Dim lookFor As Variant
    Dim thisValue As Long
    Dim foundGMT As Boolean
    Dim foundNeg As Boolean
    Dim foundString As Boolean
    Dim foundHrs As Boolean
    Dim foundMins As Boolean: foundMins = False
    Dim subString As String
    Dim hoursOffset As Integer
    Dim minsOffset As Integer
    
    On Error GoTo getRemoteOffset_Error
    
    ' check for GMT 1-3
    subString = Left$(entry, 3)
    foundGMT = InStr(subString, "GMT")
    
    ' check for  +/- at pos. 5
    subString = Mid$(entry, 5, 1)
    If InStr(subString, "-") = 1 Then
        foundNeg = True
    Else
        foundNeg = False
    End If
    
    ' check for a string at 13 - end
    subString = Mid$(entry, 13, Len(entry))
    If subString <> vbNullString Then foundString = True
    
    ' check for a valid time at pos. 7-11
    subString = Mid$(entry, 7, 5)
    If IsNumeric(Mid$(subString, 1, 2)) Then
        hoursOffset = Val(Mid$(subString, 1, 2))
        foundHrs = True
    End If
    If IsNumeric(Mid$(subString, 4, 2)) Then
        minsOffset = Val(Mid$(subString, 4, 2))
        foundMins = True
    End If
    
    ' check all tests have passed
    If foundGMT = True And foundNeg = True And foundString = True And _
        foundHrs = True And _
        foundMins = True Then
        found = True
    Else
        found = False
        getRemoteOffset = thisValue
        Exit Function
    End If
        
    If (found = True) Then
        thisValue = minsOffset + (60 * hoursOffset)
        If foundNeg = True Then
            getRemoteOffset = thisValue - thisValue * 2
            Exit Function
        Else
            getRemoteOffset = thisValue
            Exit Function
        End If
    End If
    
    getRemoteOffset = Null 'return null;
 
    On Error GoTo 0
    Exit Function

getRemoteOffset_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getRemoteOffset of Module modDaylightSavings"
 End Function




'---------------------------------------------------------------------------------------
' Procedure : getDLSrules
' Author    : beededea
' Date      : 06/10/2023
' Purpose   : read the rule list from file
' ["US", "Apr", "Sun>=1", "120", "60", "Oct", "lastSun", "60"]
'---------------------------------------------------------------------------------------
'
Public Function getDLSrules(ByVal path As String) As String()
    
    Dim ruleList() As String
    Dim rules() As String
    Dim iFile As Integer: iFile = 0
    Dim I As Variant
    Dim lFileLen As Long
    Dim sBuffer As String
    Dim useloop As Integer: useloop = 0
    Dim arraySize As Integer
    
    On Error GoTo getDLSrules_Error

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
    For Each I In ruleList
        ' Note: to replicate the .js we should .split the rule by comma and read the contents into
        ' a 2-dimensional rules array but we run into VB6 Redim problems on 2 dimensional arrays
        ' instead we will parse the rules string when we need it.
        rules(useloop) = CStr(I)
        useloop = useloop + 1
    Next I
    
ErrorHandler:
    If iFile > 0 Then Close #iFile
    
    getDLSrules = rules ' return

    On Error GoTo 0
    Exit Function

getDLSrules_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getDLSrules of Module modDaylightSavings"
End Function


'---------------------------------------------------------------------------------------
' Procedure : getNumberOfMonth
' Author    : beededea
' Date      : 07/10/2023
' Purpose   : get the number of the month given a month name
'---------------------------------------------------------------------------------------
'
Public Function getNumberOfMonth(ByVal thisMonth As String) As Integer
'    Dim monthsString As String: monthsString = vbNullString
'    Dim monthArray() As String
'    Dim months(11) As String
'    Dim i As Variant
'    Dim useLoop As Integer: useLoop = 0
    
    On Error GoTo getNumberOfMonth_Error

'    monthsString = "Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11"
'    monthArray = Split(monthsString, ",")
    
    getNumberOfMonth = Month(CDate(thisMonth & "/1/2000"))
    getNumberOfMonth = getNumberOfMonth - 1 ' convert 'normal month starting number of 1 to starting with 0 UTC
    
'    For Each i In monthArray
'        months(useLoop) = CStr(i)
'        If InStr(months(useLoop), thisMonth) > 0 Then
'            getNumberOfMonth = Val(LTrim$(Mid$(months(useLoop), 6, Len(months(useLoop))))) ' return
'            Exit Function
'        End If
'        useLoop = useLoop + 1
'    Next i

    If getNumberOfMonth < 0 Or getNumberOfMonth > 11 Then
        MsgBox ("getNumberOfMonth: " & thisMonth & " is not a valid month name")
        getNumberOfMonth = 99 ' return invalid
    End If
    
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
    Dim I As Variant
    Dim useloop As Integer: useloop = 0
    
    On Error GoTo getNumberOfDay_Error

    daysString = "Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6"
    dayArray = Split(daysString, ",")
    
    For Each I In dayArray
        days(useloop) = CStr(I)
        If InStr(days(useloop), thisDay) > 0 Then
            getNumberOfDay = Val(LTrim$(Mid$(days(useloop), 6, Len(days(useloop))))) ' return
            Exit Function
        End If
        useloop = useloop + 1
    Next I

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
    Dim useloop As Integer: useloop = 0
    
    On Error GoTo getmonthsIn_Error
    
    If thisMonth < 0 And thisMonth > 11 Then
        MsgBox ("getDaysInMonth: " & thisMonth & " is not a valid month number")
        getDaysInMonth = 99 ' return invalid
        Exit Function
    End If

    monthDaysString = "31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31"
    monthDaysArray = Split(monthDaysString, ",")
    
    If thisMonth <> 1 Then ' all except Feb
        getDaysInMonth = Val(LTrim$(monthDaysArray(thisMonth))) ' return
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
    Dim tDay As Integer: tDay = 0
    Dim tMonth As Integer: tMonth = 0
    Dim last As Integer: last = 0
    Dim d As Date
    Dim lastDay As Long

    On Error GoTo getDateOfFirst_Error

    tDay = getNumberOfDay(dayName)
    tMonth = getNumberOfMonth(monthName)
    
    If tDay = 99 Or tMonth = 99 Then
        getDateOfFirst = 99 ' return invalid
        Exit Function
    End If
    
    last = thisDayNumber + 6
    
    ' convert starting with 0 UTC to normal month starting number of 1 for the VB6 CDate function to cope with
    d = CDate(last & "/" & tMonth + 1 & "/" & thisYear)
    
    lastDay = Weekday(d, vbSunday)
        
    getDateOfFirst = last - (lastDay - tDay + 7) Mod 7 'return

    On Error GoTo 0
    Exit Function

getDateOfFirst_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getDateOfFirst of Module modDaylightSavings"
End Function


'---------------------------------------------------------------------------------------
' Procedure : getDateOfLast
' Author    : beededea
' Date      : 07/10/2023
' Purpose   : get Date (1..31) Of Last dayName (Sun..Sat) of monthName (Jan..Dec) of year (2004..)
'             dayName:     Sun, Mon, Tue, Wed, Thu, Fr, Sat
'             monthName:   Jan, Feb, etc.
'---------------------------------------------------------------------------------------
'
Public Function getDateOfLast(ByVal dayName As String, ByVal monthName As String, ByVal thisYear As Integer) As Integer
    Dim tDay As Integer: tDay = 0
    Dim tMonth As Integer: tMonth = 0
    Dim last As Integer: last = 0
    Dim d As Date
    Dim lastDay As Long
    
    On Error GoTo getDateOfLast_Error

    tDay = getNumberOfDay(dayName)
    tMonth = getNumberOfMonth(monthName)
    
    If tDay = 99 Or tMonth = 99 Then
        getDateOfLast = 99 ' return invalid
        Exit Function
    End If
    
    last = getDaysInMonth(tMonth, thisYear)

    ' convert starting with 0 UTC to normal month starting number of 1 for the VB6 CDate cast to cope with
    d = CDate(last & "/" & tMonth + 1 & "/" & thisYear)
    
    'lastDayDate = DateSerial(thisYear, tMonth, last)
    lastDay = Weekday(d, vbSunday)

    getDateOfLast = last - (lastDay - tDay + 7) Mod 7 'return

    On Error GoTo 0
    Exit Function

getDateOfLast_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getDateOfLast of Module modDaylightSavings"

End Function


'---------------------------------------------------------------------------------------
' Procedure : dayOfMonth
' Author    : beededea
' Date      : 09/10/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function dayOfMonth(ByVal monthName As String, ByVal dayRule As String, ByVal thisYear As Integer) As Integer
    Dim dayName As String: dayName = vbNullString
    Dim thisDate As String: thisDate = vbNullString

    On Error GoTo dayOfMonth_Error

    If IsNumeric(dayRule) Then
        dayOfMonth = CInt(dayRule)
        Exit Function
    End If

    ' dayRule of form lastThu or Sun>=15
    If InStr(dayRule, "last") = 1 Then '    // dayRule of form lastThu
        dayName = Mid$(dayRule, 5)
        dayOfMonth = getDateOfLast(dayName, monthName, thisYear)
        Exit Function
    End If
    
'    // dayRule of form Sun>=15
    dayName = Mid$(dayRule, 3)
    thisDate = Val(Mid$(dayRule, 4))
    dayOfMonth = getDateOfFirst(dayName, thisDate, monthName, thisYear)

    On Error GoTo 0
    Exit Function

dayOfMonth_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure dayOfMonth of Module modDaylightSavings"
End Function



'---------------------------------------------------------------------------------------
' Procedure : theDLSdelta
' Author    : beededea
' Date      : 09/10/2023
' Purpose   :
'// parameter 1 all the rules
'// parameter 2 prefs selected rule eg. ["US","Apr","Sun>=1","120","60","Oct","lastSun","60"];
'// parameter 3 remote GMT Offset
'---------------------------------------------------------------------------------------
'
Public Function theDLSdelta(ByRef DLSrules() As String, ByVal rule As String, ByVal cityTimeOffset As Long) As Long
'
    On Error GoTo theDLSdelta_Error
    
'   set up variables
    Dim monthName() As String
'    Dim arrayNumber As Integer: arrayNumber = 0
    Dim startMonth As String: startMonth = vbNullString
    Dim startDay As String: startDay = vbNullString
    Dim startTime As String: startTime = vbNullString
    Dim delta As String: delta = vbNullString
    Dim endMonth  As String: endMonth = vbNullString
    Dim endDay As String:  endDay = vbNullString
    Dim endTime As String: endTime = vbNullString
    
    Dim useUTC As Boolean: useUTC = False
    Dim theDate As Date
    Dim startYear As Integer: startYear = 0
    Dim endYear As Integer: endYear = 0
    Dim currentMonth As String: currentMonth = vbNullString
    Dim newMonthNumber As Integer: newMonthNumber = 0
    Dim startDate As Integer: startDate = 0
    Dim endDate As Integer: endDate = 0
    Dim stdTime As Date
    Dim theGMTOffset As Long: theGMTOffset = 0
    Dim startHour As Integer: startHour = 0
    Dim startMin As Integer: startMin = 0
    Dim theStart As Date
    Dim endHour As Integer: endHour = 0
    Dim endMin As Integer: endMin = 0
    Dim theEnd As Date
    Dim dlsRule As Variant
    
    Dim useloop As Integer: useloop = 0
    Dim arrayElementPresent As Boolean: arrayElementPresent = False
    Dim arrayNumber As Integer: arrayNumber = 0
    Dim ruleString As String: ruleString = vbNullString
    Dim buildDate As String: buildDate = vbNullString
    
    monthName = ArrayString("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    
    Debug.Print ("theDLSdelta(" & rule & ", " & cityTimeOffset & ")")
'
'     check whether DLS is in operation
'
    If rule = "NONE" Then
        theDLSdelta = 0 ' return abnormal
        Exit Function
    End If
    
    arrayElementPresent = False
    
    ' find at least one matching rule in the list
    For useloop = 0 To UBound(DLSrules)
        Dim separator
        separator = (""",""")
        dlsRule = Split(DLSrules(useloop), separator)
        ruleString = Mid$(dlsRule(0), 3, Len(dlsRule(0)))  '
        
        If ruleString = rule Then
            arrayElementPresent = True
            arrayNumber = useloop
            Exit For
        End If
    Next useloop
    
    If arrayElementPresent = False Then
        Debug.Print ("DLSdelta: " & rule & " is not in the list of DLS rules.")
        theDLSdelta = 0 ' return abnormal
        Exit Function
    End If

'    // extract the current rule from the rules array using the arrayNumber
'
    dlsRule = Split(DLSrules(arrayNumber), separator)
'
'    // read the various components of the split rule
'
    startMonth = dlsRule(1)
    startDay = dlsRule(2)
    startTime = dlsRule(3)
    delta = dlsRule(4)
    endMonth = dlsRule(5)
    endDay = dlsRule(6)
    endTime = Left$(dlsRule(7), Len(dlsRule(7)) - 2)

'["AR","Oct","Sun>=15","0","60","Mar","Sun>=15","-60"]
'["US", "Apr", "Sun>=1", "120", "60", "Oct", "lastSun", "60"]

'    negative times for UTC transitions (GMT starts a mid-day)
'
    useUTC = (startTime < 0) And (endTime < 0)
'
    If (useUTC) Then
        startTime = 0 - startTime
        endTime = 0 - endTime
    End If
    
    Debug.Print ("Rule:       " & rule)
    Debug.Print ("startMonth: " & startMonth)
    Debug.Print ("startDay:   " & startDay)
    Debug.Print ("startTime:  " & startTime)
    Debug.Print ("delta:      " & delta)
    Debug.Print ("endMonth:   " & endMonth)
    Debug.Print ("endDay:     " & endDay)
    Debug.Print ("endTime:    " & endTime)
    Debug.Print ("useUTC:     " & useUTC)

    theDate = Now()
    startYear = Year(theDate)
    endYear = startYear
    
    If getNumberOfMonth(startMonth) >= 6 Then          ' Southern Hemisphere
        currentMonth = Month(theDate)
        If currentMonth >= 6 Then
            endYear = endYear + 1
        Else
            startYear = startYear - 1
        End If
    End If

    If startTime < 0 Then
        startTime = 0 - startTime
    End If  ' ignore invalid sign

    startDate = dayOfMonth(startMonth, startDay, startYear)
    If startDate = 0 Then
        theDLSdelta = 0 ' return abnormal
        Exit Function
    End If
    
    endDate = dayOfMonth(endMonth, endDay, endYear)
    If endDate = 0 Then
        theDLSdelta = 0 ' return abnormal
        Exit Function
    End If
    
    If endTime < 0 Then ' transition on previous day in standard time
        endTime = 0 - endTime
        endDate = endDate - 1
        endTime = 1440 - endTime
        If (endDate = 0) Then
            newMonthNumber = getNumberOfMonth(endMonth) - 1
            endMonth = monthName(newMonthNumber)
            endDate = getDaysInMonth(newMonthNumber, endYear)
        End If
    End If
    
    Debug.Print ("startDate:  " & startMonth & " " & startDate & "," & startYear)
    Debug.Print ("startTime:  " & (startTime - startTime Mod 60) / 60 & ":" & startTime Mod 60)
    Debug.Print ("endDate:    " & endMonth & " " & endDate & "," & endYear)
    Debug.Print ("endTime:    " & (endTime - endTime Mod 60) / 60 & ":" & endTime Mod 60)

    theGMTOffset = 60000 * cityTimeOffset    '// was preferences.cityTimeOffset.value
    
    theDate = Now()
    stdTime = Now()

    startHour = Int(startTime / 60)
    startMin = startTime Mod 60
    
    Debug.Print ("----")
    Debug.Print ("startYear=" & startYear)
    Debug.Print ("startMonth=" & (startMonth))
    Debug.Print ("startDate=" & startDate)
    Debug.Print ("startHour=" & startHour)
    Debug.Print ("startMin=" & startMin)
    
    buildDate = Str$(startDate) & "/" & startMonth & "/" & Str$(startYear) & " " & Str$(startHour) & ":" & Str$(startMin)
    theStart = CDate(buildDate)
    
    If useUTC = False Then
        theStart = theStart - theGMTOffset
    End If

    Debug.Print ("theStart= " & theStart)

    endHour = Int(endTime / 60)
    endMin = endTime Mod 60

    Debug.Print ("----")
    Debug.Print ("endYear=" & endYear)
    Debug.Print ("endMonth=" & endMonth)
    Debug.Print ("endDate=" & endDate)
    Debug.Print ("endHour=" & endHour)
    Debug.Print ("endMin=" & endMin)
    
    buildDate = Str$(endDate) & "/" & endMonth & "/" & Str$(endYear) & " " & Str$(endHour) & ":" & Str$(endMin)
    theEnd = CDate(buildDate)

    If useUTC = False Then
        theEnd = theEnd - theGMTOffset
    End If
    
    Debug.Print ("theEnd=   " & theEnd)

    If (stdTime < theStart) Then
        Debug.Print ("DLS starts in " & Int((theStart - stdTime) / 60000) & " minutes.")
    ElseIf (stdTime < theEnd) Then
        Debug.Print ("DLS ends in   " & Int((theEnd - stdTime) / 60000) & " minutes.")
    End If
'
    If (theStart <= stdTime) And (stdTime < theEnd) Then
        Debug.Print ("----DLSdelta=" & delta)
        theDLSdelta = delta ' return
        Exit Function
    Else
        Debug.Print ("----DLSdelta=0")
        theDLSdelta = 0 ' return abnormal
        Exit Function
    End If

    On Error GoTo 0
    Exit Function

theDLSdelta_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure theDLSdelta of Module modDaylightSavings"
End Function






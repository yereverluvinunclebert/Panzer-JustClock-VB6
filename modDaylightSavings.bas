Attribute VB_Name = "modDaylightSavings"
Public Sub obtainDaylightSavings()
    Dim getDLSrules() As String
    Dim numberOfMonth As String
    
    ' read the rule list from file
    getDLSrules = getDSTRules(App.path & "\Resources\txt\DLSRules.txt")
    
    ' get the number of the month given a month name
    numberOfMonth = getNumberOfMonth("Oct")
    
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
    'Open path For Input As #iFile
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
Public Function getNumberOfMonth(month) As String
    Dim monthsString As String
    Dim monthArray() As String
    Dim months(11) As String
    Dim i As Variant
    Dim a As String: a = vbNullString
    Dim useLoop As Integer: useLoop = 0
    
    On Error GoTo getNumberOfMonth_Error

    monthsString = "Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11"
    monthArray = Split(monthsString, ",")
    
    For Each i In monthArray
        months(useLoop) = CStr(i)
        If InStr(months(useLoop), month) Then
            getNumberOfMonth = Right$(months(useLoop), 1)
            Exit Function
        End If
        useLoop = useLoop + 1
    Next i

    MsgBox ("getNumberOfMonth: " & month & " is not a valid month name")
    getNumberOfMonth = a

    On Error GoTo 0
    Exit Function

getNumberOfMonth_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getNumberOfMonth of Module modDaylightSavings"

End Function

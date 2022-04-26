VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'================================================================================================
'                                   Property
'================================================================================================
Private header As scripting.Dictionary
Private colTypes As scripting.Dictionary
Private rowSize As Long, colSize As Long
Private data As Variant
'================================================================================================
'                                   Enum
'================================================================================================
Enum StatusCode
    Error = -1
    Fail = 0
    Pass = 1
End Enum

Enum CalcFunctionType
    Sum
    Average
End Enum
'================================================================================================
'                                   Const
'================================================================================================
Const Type_General = "General"
Const Type_Numeric = "Numeric"
'================================================================================================
'                                   Class Initializer
'================================================================================================
Private Sub Class_Initialize()
    Set header = Nothing
    data = vbNull
End Sub
'================================================================================================
'                                   Getter
'================================================================================================
Function getRowSize() As Long
    getRowSize = rowSize
End Function
Function getColSize() As Long
    getColSize = colSize
End Function
Function getData() As Variant
    getData = data
End Function
Function getColIndex(colName As String) As Long
    getColIndex = 0
    If hasHeader Then If header.Exists(colName) Then getColIndex = header(colName)
End Function
Function hasHeader() As Boolean
    If Not header Is Nothing Then hasHeader = True
End Function
Function hasData() As Boolean
    If Not isArrayEmpty(data) Then hasData = True
End Function
Function hasColumn(columnName As String) As Boolean
    If hasHeader Then hasColumn = header.Exists(columnName)
End Function
Function isGridEmpty() As Boolean
    If Not hasHeader And Not hasData Then isGridEmpty = True
End Function
'================================================================================================
'                                   Setter
'================================================================================================
Sub setHeader(dictHeader As scripting.Dictionary)
    Set header = dictHeader
End Sub
Sub setData(varData As Variant)
    data = varData
End Sub
Public Function validateSize()
    If Not hasData Then rowSize = 0: colSize = 0: Exit Function
    rowSize = UBound(data, 1)
    colSize = UBound(data, 2)
End Function
Public Function initializeColTypes() As scripting.Dictionary
    Set colTypes = New scripting.Dictionary
    Dim i As Long
    For i = 1 To colSize
        Call colTypes.Add(i, Type_General)
    Next
End Function
'================================================================================================
'                                   Load Data
'================================================================================================
Public Function load(dataSource As String, path As String) As Integer

End Function
Public Function loadFromRange(rng As Range, rngHasHeader As Boolean) As datagrid
On Error GoTo errorHandler
    Set loadFromRange = Me
    If rng.Areas.count = 1 Then
        Dim rowNum As Long, colNum As Long
        Set colTypes = Nothing: Set header = Nothing: data = vbNull 'clear old header and data
        rowNum = rng.Rows.count
        colNum = rng.Columns.count
        If rngHasHeader And rowNum > 1 Then
            rowNum = rowNum - 1
            Set header = getDictionaryFromArray(rng.Rows(1).Value2, True)
            data = rng.Rows("2:" & rowNum + 1).value
            rowSize = rowNum
            colSize = UBound(data, 2)
        Else
            data = rng.Rows("1:" & rowNum).Value2
            rowSize = rowNum
            colSize = UBound(data, 2)
        End If
        Call initializeColTypes
        Debug.Print "loadFromRange Status: " & StatusCode.Pass
    Else
        Debug.Print "loadFromRange Status: " & StatusCode.Fail
    End If
    
    Exit Function
errorHandler:
    Debug.Print "loadFromRange Status: " & StatusCode.Error
End Function
Public Function loadFromCSV(strFilePath As String, delimiter As String, hasHeader As Integer) As datagrid

End Function

'================================================================================================
'                                   Output Data
'================================================================================================
Public Function copy() As datagrid
    Dim result As datagrid
    Set result = New datagrid
    Call result.setHeader(header)
    Call result.setData(data)
    Call result.validateSize
    Call result.initializeColTypes
    Set copy = result
End Function
Public Function toConsole()
    Debug.Print
    Debug.Print "The Data Grid is: "
    Debug.Print "====================================================================="
    Debug.Print "Row size: " & rowSize & "   Column size: " & colSize
If Me.hasHeader Then
    Debug.Print "----------------------------HEADER-----------------------------------"
    If header.count > 0 Then Call printArrayToConsole(header.Keys())
End If
If Me.hasData Then
    Debug.Print "----------------------------TYPE-------------------------------------"
    If colTypes.count > 0 Then Call printArrayToConsole(colTypes.Items())
    Debug.Print "-----------------------------DATA------------------------------------"
    Call printArrayToConsole(data)
Else
    Debug.Print "--------------------------NO DATA------------------------------------"
End If
    Debug.Print "====================================================================="
End Function
Public Function toDictionary() As scripting.Dictionary

End Function
Public Function toArray() As Variant

End Function
Public Function toRange(rng As Range, Optional expandRange As Boolean = False)
    Dim rngRowCount As Long: rngRowCount = rng.Rows.count
    Dim rngColCount As Long: rngColCount = rng.Columns.count
    Dim i As Long, j As Long, startRow As Long
    Dim headerKeys As Variant
    If rngRowCount < 1 Or rngColCount < 1 Then Exit Function
    Call excelSetup(False, False, False)
    If Not expandRange Then
        rngRowCount = IIf(rngRowCount < rowSize, rngRowCount, rowSize)
        rngColCount = IIf(rngColCount < colSize, rngColCount, colSize)
    Else
        rngRowCount = rowSize
        rngColCount = colSize
    End If
    If hasHeader Then 'Write header if has header
        headerKeys = header.Keys()
        For j = 1 To rngColCount
            rng.Cells(1, j).value = headerKeys(j - 1)
        Next
        startRow = 2
    Else
        startRow = 1
    End If
    For i = startRow To rngRowCount + startRow - 1
        For j = 1 To rngColCount
            If hasHeader Then
                rng.Cells(i, j).value = data(i - 1, j)
            Else
                rng.Cells(i, j).value = data(i, j)
            End If
        Next
    Next
    Call excelSetup(True, True, True)
End Function
'================================================================================================
'                                   Data Lookup
'================================================================================================
Public Function contains(value As Variant, Optional columnName As String) As Boolean
    Dim colIndex As Long, searchArray As Variant
    If columnName = vbNullString Then
        searchArray = data
    Else
        If hasColumn(columnName) Then
            colIndex = header(columnName)
            searchArray = Application.WorksheetFunction.index(data, 0, colIndex)
        End If
    End If
    contains = isInArray(value, searchArray, True)
End Function
Public Function lookUp(targetcolName As String, lookUpColNameArray As Variant, lookUpValueArray As Variant) As Variant
    If Not hasColumn(targetcolName) Then lookUp = StatusCode.Error: Exit Function
    If getDimension(lookUpColNameArray) = 0 And getDimension(lookUpValueArray) = 0 Then
        If Not varType(lookUpColNameArray) = vbString Or Not isValidDataType(lookUpValueArray) Then lookUp = StatusCode.Error: Exit Function
        lookUpColNameArray = Array(lookUpColNameArray)
        lookUpValueArray = Array(lookUpValueArray)
    ElseIf getDimension(lookUpColNameArray) = 1 And getDimension(lookUpValueArray) = 1 Then
        If UBound(lookUpColNameArray) - LBound(lookUpColNameArray) <> UBound(lookUpValueArray) - LBound(lookUpValueArray) Then lookUp = StatusCode.Error: Exit Function
    Else
        lookUp = StatusCode.Error: Exit Function
    End If
    Dim row As Long, lookUpValue As String, result As String, var As Variant
    For Each var In lookUpColNameArray 'Check all columns exist
        If hasColumn(CStr(var)) = False Then lookUp = StatusCode.Error: Exit Function
    Next
    For Each var In lookUpValueArray 'Build lookup string
        lookUpValue = lookUpValue & "/" & var
    Next
    For row = 1 To rowSize 'Loop through all rows
        result = ""
        For Each var In lookUpColNameArray
            result = result & "/" & data(row, header(CStr(var)))
        Next
        If lookUpValue = result Then
            lookUp = data(row, header(targetcolName))
            Exit Function
        End If
    Next
    lookUp = StatusCode.Fail 'No match
End Function
'================================================================================================
'                                   Data Transformation
'================================================================================================
Public Function filterOut(colName As String, arrayValues, Optional blnExactMatch As Boolean = True) As datagrid
    Set filterOut = Me
    If hasData = False Then Exit Function
    Call filter(colName, arrayValues, True, blnExactMatch)
End Function
Public Function filterIn(colName As String, arrayValues, Optional blnExactMatch As Boolean = True) As datagrid
    Set filterIn = Me
    If hasData = False Then Exit Function
    Call filter(colName, arrayValues, False, blnExactMatch)
End Function
Public Function groupBy(aggregation As CalcFunctionType, colName As String, colNameArray As Variant) As datagrid
    Set groupBy = Me
    If Not hasColumn(colName) Then Debug.Print StatusCode.Error: Exit Function
    If getDimension(colNameArray) = 0 Then
        If isValidDataType(colNameArray) Then colNameArray = Array(colNameArray) Else Debug.Print StatusCode.Error: Exit Function
    End If
    If checkHeader(colNameArray) <> StatusCode.Pass Then Debug.Print StatusCode.Error: Exit Function
    Select Case aggregation
    Case CalcFunctionType.Sum
        Call evaluateColType(header(colName)) 'may be better deleted
        If colTypes(header(colName)) = Type_Numeric Then Call groupBySum(colName, colNameArray): Debug.Print StatusCode.Pass
    Case CalcFunctionType.Average
        Call evaluateColType(header(colName))
        If colTypes(header(colName)) = Type_Numeric Then
        End If
    Case Else
    End Select
End Function
Public Function pivot()

End Function
Public Function unpivot()

End Function
Public Function resize(x As Long, y As Long)
    Me.setData (resizeArray(data, x, y))
End Function
Public Function transpose()

End Function
Public Function sortBy()

End Function
'=============================================================================
'                           Private Helper
'=============================================================================
Private Function checkHeader(arrayColNames As Variant) As Variant
    checkHeader = StatusCode.Fail
    If Me.hasHeader = False Then checkHeader = StatusCode.Error: Exit Function 'check has header
    Dim varColName As Variant, strColName As String
    For Each varColName In arrayColNames 'loop through array
        strColName = CStr(varColName)
        If strColName = "" Then checkHeader = StatusCode.Error: Exit Function
        If Not header.Exists(strColName) Then checkHeader = StatusCode.Fail: Exit Function
    Next
    checkHeader = StatusCode.Pass
End Function
Public Sub evaluateColType(colNum As Long)
    If hasData = False Then Exit Sub
    Dim i As Integer, var As Variant
    For i = 1 To rowSize
        var = data(1, colNum)
        If varType(var) < 2 Or varType(var) > 6 Then Exit Sub
    Next
    colTypes(colNum) = Type_Numeric
End Sub
Private Sub excelSetup(blnAutoCalc As Boolean, blnEvents As Boolean, blnScreen As Boolean, Optional blnAlerts As Boolean, Optional blnStatusBar As Boolean)
    Dim vrtCalc As Variant
    With Application
        If blnAutoCalc = True Then
            vrtCalc = xlCalculationAutomatic
        ElseIf blnAutoCalc = False Then
            vrtCalc = xlCalculationManual
        End If
        .Calculation = vrtCalc
        .ScreenUpdating = blnScreen
        .EnableEvents = blnEvents
        .DisplayAlerts = blnAlerts
        If Not IsMissing(blnStatusBar) Then .StatusBar = blnStatusBar
    End With
End Sub
Private Function filter(colName As String, valArray As Variant, Optional blnFilterOut As Boolean = False, Optional blnExactMatch As Boolean = True) As Variant
    Dim row As Long, newRow As Long, col As Long, i As Long
    Dim newData() As Variant, arrayCriteria As Variant
    Dim dCount As Integer: dCount = getDimension(valArray)
    If hasColumn(colName) = False Then filter = StatusCode.Error: Exit Function
    If dCount > 1 Then filter = StatusCode.Error: Exit Function
    If rowSize = 0 Then filter = StatusCode.Error: Exit Function
    If dCount = 0 And isValidDataType(valArray) Then valArray = Array(valArray) 'if one value
    col = header(colName)
    newRow = 0
    ReDim newData(1 To rowSize, 1 To colSize)
    For row = 1 To rowSize 'Loop through rows to filter
        If Not isInArray(data(row, col), valArray, blnExactMatch) = blnFilterOut Then
            newRow = newRow + 1
            For i = 1 To colSize
                newData(newRow, i) = data(row, i)
            Next
        End If
    Next
    'Resize
    newData = resizeArray(newData, newRow, header.count)
    'Update Property
    Me.setData (newData)
    Me.validateSize
End Function
Private Function getDictionaryFromArray(arr As Variant, Optional blnGetByRow As Boolean = True, Optional dict As scripting.Dictionary) As scripting.Dictionary
    Dim i As Long, arrRowSize As Long, arrColSize As Long
    Dim dCount As Integer
    Dim dictResult As scripting.Dictionary
    Dim varKey As Variant, varVal As Variant
    If dict Is Nothing Then Set dictResult = New scripting.Dictionary Else Set dictResult = dict
    dCount = getDimension(arr)
    If dCount = 0 And isValidDataType(arr) Then
        Call dictResult.Add(arr, 1)
    ElseIf dCount = 1 Then '1-D array: element as key, index (start at 1) as value
        For i = LBound(arr) To UBound(arr)
            varKey = CStr(arr(i))
            If varKey = vbNullString Then varKey = i & ""
            If dictResult.Exists(varKey) Then varKey = varKey & "1"
            Call dictResult.Add(varKey, i)
        Next
    ElseIf dCount = 2 Then '2-D array, read by row
        arrRowSize = UBound(arr, 1) - LBound(arr, 1) + 1
        arrColSize = UBound(arr, 2) - LBound(arr, 2) + 1
        If blnGetByRow Then
            For i = LBound(arr, 2) To UBound(arr, 2) 'loop through column
                varKey = CStr(arr(LBound(arr, 1), i))
                If arrRowSize > 1 Then varVal = arr(LBound(arr, 1) + 1, i) Else varVal = i
                If varKey = vbNullString Then varKey = "Column" & i
                If dictResult.Exists(varKey) Then varKey = varKey & "1"
                Call dictResult.Add(varKey, varVal)
            Next
        Else '2-D array, read by col
            For i = LBound(arr, 1) To UBound(arr, 1) 'loop through rows
                varKey = CStr(arr(i, LBound(arr, 2)))
                If arrColSize > 1 Then varVal = arr(i, LBound(arr, 2) + 1) Else varVal = i
                If varKey = vbNullString Then varKey = "Row" & i
                If dictResult.Exists(varKey) Then varKey = varKey & "1"
                Call dictResult.Add(varKey, varVal)
            Next
        End If
    End If
    Set getDictionaryFromArray = dictResult
End Function
Private Function getDictionaryFromRange(rng As Range, Optional blnGetByRow As Boolean = True, Optional dict As Dictionary) As Variant
    Set getDictionaryFromRange = getDictionaryFromArray(rng.Value2, blnGetByRow, dict)
End Function
Private Function getDimension(var As Variant) As Long
    On Error GoTo Err
    Dim count As Long, temp As Long
    count = 0
    Do While True
        count = count + 1
        temp = UBound(var, count)
    Loop
Err:
    getDimension = count - 1
End Function
Private Function groupBySum(aggregateColName As String, colNameArray As Variant)
    Dim var As Variant
    Dim colName As String, keyString As String
    Dim newCol As Long, newRow As Long, row As Long, col As Long, newRowCount As Long
    Dim keyDict As scripting.Dictionary: Set keyDict = New scripting.Dictionary
    Dim newData As Variant
    ReDim newData(1 To rowSize, 1 To colSize)
    'Build New Header
    Dim newHeader As scripting.Dictionary: Set newHeader = New scripting.Dictionary
    For Each var In colNameArray
        colName = CStr(var)
        If colName <> "" Then newCol = newCol + 1: newHeader.Add colName, newCol
    Next
    newHeader.Add aggregateColName, newCol + 1
    'Build New Data
    newCol = 0: newRow = 0: newRowCount = 0
    For row = 1 To rowSize
        keyString = ""
        col = header(colName): newCol = newHeader(colName)
        For Each var In colNameArray
            colName = CStr(var)
            keyString = keyString & "/" & data(row, header(colName))
        Next
        If Not keyDict.Exists(keyString) Then
            newRow = newRowCount + 1
            newRowCount = newRowCount + 1
            keyDict.Add keyString, newRow 'add the new row into dictionary
            For Each var In colNameArray
                colName = CStr(var)
                newCol = newHeader(colName)
                col = header(colName)
                newData(newRow, newCol) = data(row, col)
            Next
            newData(newRow, newHeader(aggregateColName)) = data(row, header(aggregateColName))
        Else
            newRow = keyDict(keyString) 'find the row number from dictionary
            newData(newRow, newHeader(aggregateColName)) = newData(newRow, newHeader(aggregateColName)) + data(row, header(aggregateColName))
        End If
    Next
    Call Me.setHeader(newHeader)
    Call Me.setData(resizeArray(newData, keyDict.count, newHeader.count))
    Call Me.validateSize
    Call Me.initializeColTypes
    colTypes(header(aggregateColName)) = Type_Numeric
End Function
Private Function isInArray(val As Variant, arr As Variant, Optional blnExactMatch As Boolean = True) As Boolean
    Dim element As Variant, searchStr As String
    Dim valVarType As Variant
    If Not isValidDataType(val) And varType(val) <> vbEmpty Then Exit Function
    searchStr = CStr(val)
    If blnExactMatch Then
        For Each element In arr
            If searchStr = element Then isInArray = True: Exit Function
        Next
    Else
        For Each element In arr
            If searchStr Like "*" & element & "*" Then isInArray = True: Exit Function
        Next
    End If
End Function
Private Function isArrayEmpty(arr As Variant) As Boolean
    Dim var As Variant
    isArrayEmpty = True
    If varType(arr) > vbArray Then
        For Each var In arr
            If isEmpty(var) = False Then isArrayEmpty = False: Exit Function
        Next
    End If
End Function
Private Function isValidDataType(val As Variant) As Boolean
    Dim valVarType As Long
    valVarType = varType(val)
    If valVarType < 2 Or valVarType > 8 Then Exit Function
    isValidDataType = True
End Function
Private Function printArrayToConsole(InputArray As Variant)
'Read 1D or 2D Array and Print to Console
    Dim intNumOfD As Integer
    Dim rowString As String
    Dim lRow As Long, lCol As Long
    intNumOfD = getDimension(InputArray)
    If intNumOfD = 2 Then '2D array
        For lRow = LBound(InputArray, 1) To UBound(InputArray, 1)
            rowString = InputArray(lRow, LBound(InputArray, 2))
            For lCol = LBound(InputArray, 2) + 1 To UBound(InputArray, 2)
                rowString = rowString & "," & InputArray(lRow, lCol)
            Next
            Debug.Print rowString
        Next
    ElseIf intNumOfD = 1 Then '1D array
        rowString = InputArray(LBound(InputArray))
        For lRow = LBound(InputArray) + 1 To UBound(InputArray)
            rowString = rowString & "," & InputArray(lRow)
        Next
        Debug.Print rowString
    End If
End Function
Private Function resizeArray(data As Variant, x As Long, y As Long) As Variant
    Dim dCount As Integer: dCount = getDimension(data)
    If dCount <> 2 Then resizeArray = data: Exit Function
    Dim dataRowSize As Long: dataRowSize = UBound(data, 1)
    Dim dataColSize As Long: dataColSize = UBound(data, 2)
    If x < 1 Or y < 1 Then
        ReDim data(1 To 1, 1 To 1)
        resizeArray = data: Exit Function
    Else
        x = Application.WorksheetFunction.Min(x, dataRowSize)
        y = Application.WorksheetFunction.Min(y, dataColSize)
        'Resize Col
        ReDim Preserve data(1 To dataRowSize, 1 To y)
        'Resize Row
        Dim newData As Variant, i As Long, j As Long
        ReDim newData(1 To x, 1 To y)
        For i = 1 To x
            For j = 1 To y
                newData(i, j) = data(i, j)
            Next
        Next
    End If
    resizeArray = newData
End Function

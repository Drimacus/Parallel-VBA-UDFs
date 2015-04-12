Attribute VB_Name = "AsynchWSFun"
'/*
'' Copyright (c) 2015 Michel Verlinden
'' license: MIT (http://www.opensource.org/licenses/mit-license.php)
'' https://github.com/MichelVerlinden/Parallel-VBA-UDFs
''
'' Module to manage requests for data asynchronously from worksheet
'' Handles Computation events
''
'' author : Michel Verlinden - migul.verlinden@gmail.com
'' 13/03/2014
''
''
'' TODO :   Add error handler procedure
''          Add downloadsheet cleaner
''          Add streaming fuctionality
'*/
Option Explicit
Option Private Module

Private mCalcManager As Dictionary

Private prevCalc As String
Private prevID As Double

Public executed As Boolean
Public calculating As Boolean

Private Const mDebugger = False

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Pre-callback computations (UDF thread)
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' asyncFun:     First function called from the worksheet. If the data for the function is available on data sheet
'               The function returns it otherwise it sends a time call back to check other cells who are calculating
'               In order to compute them together
'
' Parameters:   f           : IAsynchFun
'               Paramarray  : The arguments of the worksheet function
'
' Returns:      String -    This function will always return a string of data
'                           Population of an Array should be done in WS definition
Public Function asyncFun(f As IAsyncWSFun, ParamArray p() As Variant) As String
    On Error GoTo endf
    ' Below line is to prevent computation in function wizard
    If (Not Application.CommandBars("Standard").Controls(1).Enabled) Then Exit Function
    If Debugger.debugging And mDebugger And False Then
        logFunctionCall "asyncFun(f As IAsyncWSFun, ParamArray p() As Variant)", f, p
    End If
    
    ' Extract input into 1d Array
    Dim params() As Variant
    params = Util.formatPArray(p)
    
    ' Check if data has already been fetched
    Dim rg As Range
    If findReq(rg, f, params) Then
        asyncFun = getDt(rg)
    ElseIf Not TypeOf Excel.Application.Caller Is Range Then
        asyncFun = "#Error: Illegal usage"
    Else
        asyncFun = "#requesting data"
        
        ' If data needs to be fetched - is a batch being constructed
        ' or do we need to create a new batch - batch = function + calling cells + data requested
        If mCalcManager Is Nothing Then
            Set mCalcManager = New Dictionary ' main handle
        End If
        Dim batch As Computation, idBatch As Double
        If isNewBatch(f, idBatch) Then
            Set batch = New Computation
            Set batch.fType = f
            batch.addCell Application.Caller, p
            mCalcManager.Add giveID(f), batch ' ID used to know f on data reception
        Else
            Set f = Nothing ' another fType is assigned this Computation
            Set batch = mCalcManager.Item(idBatch)
            batch.addCell Application.Caller, p
        End If
        If Not calculating Then
            calculating = True
            startThread ' the first calculating cell triggers a timed callback
        End If
    End If
    
    Exit Function
endf:
    Set mCalcManager = Nothing
    prevCalc = vbNullString
    prevID = 0
    calculating = False
End Function

' isNewBatch:   Check if the calculation manager needs to add a new item to dictionnary
'
' Parameters:   f  : IAsynchFun
'               id : ID of an existing batch the function can attach to
'
' Returns:      Boolean - True if a new batch calculation needs to be created
Private Function isNewBatch(ByRef f As IAsyncWSFun, ByRef id As Double) As Boolean
    isNewBatch = True
    id = 0
    If StrComp(f.getName, prevCalc) = 0 Then
        isNewBatch = False
        id = prevID
    Else
        Dim d As Variant, c As Computation
        For Each d In mCalcManager.Keys
            Set c = mCalcManager.Item(d)
            If StrComp(c.fType.getName, f.getName) = 0 Then
                If Not c.closed Then
                    isNewBatch = False
                    id = d
                End If
            End If
        Next d
    End If
End Function

' findReq:      Find if the data is available on DataSheet
'
' Parameters:   rg        : Range is assigned if the data is found
'               fType     : Need to know the name in request encoding
'               Paramarray: The arguments given to the worksheet function
'
' Returns:      Boolean -   True if the data is found
Private Function findReq(ByRef rg As Range, ByRef fType As IAsyncWSFun, _
                            ByRef p() As Variant) As Boolean
    On Error GoTo errhandler
    Dim concat As String, c As Variant
    For Each c In p
        concat = concat & Util.arrSep & CStr(c)
    Next
    Set rg = ThisWorkbook.Sheets("Data@Download").Cells.Find(What:=fType.getName & concat & Util.arrSep, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False)
        
    findReq = Not rg Is Nothing
endf:
Exit Function
errhandler:
    findReq = False
    Resume endf
End Function

' getDt:        Extract data from an encoded request on Data Sheet
'
' Parameters:   rg: Range holding the data
'
' Returns:      String -  Answer to ws function
Private Function getDt(rg As Range) As String
    Dim st() As String
    st = Split(rg.Value, Util.arrSep)
    getDt = st(UBound(st))
End Function

' startThread:  Create a Switch Object able to set a time Callback
Private Sub startThread()
    Dim sw As Switch
    Set sw = New Switch
End Sub

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Post-callback computations (timed thread)
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' timedThread:  Entry point of the computation once all calculating cells have been added to
'               the calculation manager : mCalcManager
Public Sub timedThread()
    On Error GoTo errhandler
    If Debugger.debugging And mDebugger And True Then
        logFunctionCall "timedThread()", Nothing
    End If
    If Not executed Then
        ' reset batch making controls
        prevCalc = vbNullString
        prevID = 0
        executed = True
        
        ' Close all Computations
        Dim c As Computation, d As Variant, sent As Boolean
        sent = True
        For Each d In mCalcManager.Keys
            Set c = mCalcManager.Item(d)
            c.closed = True
            If c.fType.validateRequest(c.calcRng) Then
                If c.fType.makeRequest(c.calcRng) Then
                    sent = True
                End If
            End If
            If Not sent Then
                Dim rg As Variant, b As Boolean
                b = False
                For Each rg In c.calcRng.Keys
                    If Not b Then
                        b = True
                        rg.Value = "#Invalid request"
                    Else
                        rg.Clear
                    End If
                Next
                c.killBatch
                mCalcManager.Remove d
            End If
        Next d
        AsynchWSFun.calculating = False
    End If
endf:
    Exit Sub
    
errhandler:
    Set mCalcManager = Nothing
    AsynchWSFun.calculating = False
    Debug.Print "Error in timedThread:" & Err.Number & ", " & Err.Description
    Resume endf
End Sub

' processResp:  Assign the final value to all the requesting cells for a given calculation ID
'
' Parameters:   fType     : IAsyncWSFun
Public Sub processResp(ByRef fType As IAsyncWSFun)
    On Error GoTo errhandler
    If Debugger.debugging And mDebugger And True Then
        logFunctionCall "processResp(ByRef fType As IAsyncWSFun, reqID As Double,response As Variant)", _
                        fType
    End If
    Dim r As Computation, c As Variant, strRes As String
    ' find the calling cells and compute them
    Set r = mCalcManager.Item(fType.id)
    For Each c In r.calcRng.Keys
        Dim p As Variant
        strRes = vbNullString
        If fType.processResponse(strRes, r.calcRng.Item(c)) Then
            addData c, strRes, makeReqStr(fType, r.calcRng.Item(c))
             If c.HasArray Then
                Dim carr As Range
                Set carr = c.CurrentArray
                carr.Dirty
                carr.Calculate
            Else
                c.Formula = c.Formula
            End If
            Call r.calcRng.Remove(c)
        End If
    Next
    If r.calcRng.Count = 0 Then
        r.killBatch
        mCalcManager.Remove fType.id
    End If
endf:
    If mCalcManager.Count = 0 And Debugger.debugging Then
        Debugger.stopDebugging
    End If
    Exit Sub
    
errhandler:
    Set mCalcManager = Nothing
    AsynchWSFun.calculating = False
    Debug.Print "Error in processResp:" & Err.Number & ", " & Err.Description
    Resume endf
End Sub

' killRequest:  Remove a request from the manager
'
' Parameters:   reqID: ID of the calculation to remove
Public Sub killRequest(ByVal reqID As Byte)
    If mCalcManager.Exists(reqID) Then
        Dim d As Dictionary
        Set d = mCalcManager.Item(reqID)
        d.RemoveAll
    End If
End Sub

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Helping functions
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' makeReqStr:   Create a String identifier for a function + arguments
'
' Parameters:   fType : IAsyncWSFun
'               p     : Arguments passed to fType
'
' Returns   :   String: Identifier of the request
Private Function makeReqStr(ByRef fType As IAsyncWSFun, ParamArray p() As Variant) As String
    Dim c As Variant
    For Each c In Util.formatPArray(p(0))
        makeReqStr = makeReqStr & Util.arrSep & CStr(c)
    Next
    makeReqStr = fType.getName & makeReqStr
End Function


Private Sub addData(ByVal rg As Range, str As String, reqStr As String)
    Dim sh As Worksheet
    If Not Util.checkDataSheet(sh) Then
        Util.addDataSheet sh
    End If
    If Len(sh.Range(rg.Cells(0).Address).Value) = 0 Then
        sh.Range(rg.Cells(0).Address).Value = _
            reqStr & Util.arrSep & str
    Else
        sh.Cells(Rows.Count, rg.Cells(0).Column) _
            .End(xlUp).Offset(1, 0).Value = _
            reqStr & Util.arrSep & str
    End If
End Sub

Private Function giveID(ByRef fType As IAsyncWSFun) As Integer
    Dim i As Integer
    While mCalcManager.Exists(i)
        i = i + 1
    Wend
    giveID = i
    fType.id = i
End Function

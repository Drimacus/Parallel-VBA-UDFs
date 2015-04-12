VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "YahooAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*
'' Copyright (c) 2015 Michel Verlinden
'' license: MIT (http://www.opensource.org/licenses/mit-license.php)
'' https://github.com/MichelVerlinden/Parallel-VBA-UDFs
''
'' YahooAPI
'' Send requests for Yahoo finance data using YQL
''
'' Dependencies: MSXML2
''
''
'' author : Michel Verlinden - migul.verlinden@gmail.com
'' 13/03/2014
''
''
'*/

Option Explicit
Implements IAsyncWSFun

Public yqlReq As MSXML2.XMLHTTP60
Public yqlDoc As MSXML2.DOMDocument60

Private mapToGrid As Dictionary
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Implementation of IAsyncWSFun
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Private idReq As Integer

Private Property Let IAsyncWSFun_id(ByVal RHS As Integer)
    idReq = RHS
End Property

Private Property Get IAsyncWSFun_id() As Integer
    IAsyncWSFun_id = idReq
End Property

Public Function IAsyncWSFun_catchResp() As Variant
Attribute IAsyncWSFun_catchResp.VB_UserMemId = 0
    If yqlReq.ReadyState = 4 Then
        Set yqlDoc = New MSXML2.DOMDocument60
        If yqlDoc.LoadXML(yqlReq.responseText) Then
            'Debugger.logThis yqlReq.responseText
            AsynchWSFun.processResp Me
        End If
    End If
End Function

Public Function IAsyncWSFun_getName() As String
    IAsyncWSFun_getName = "YQLQuery"
End Function

Public Function IAsyncWSFun_makeRequest(params As Scripting.IDictionary) As Boolean
    Dim symbols As String, fields As String, p As Variant, url As String
    Dim b As Boolean
    b = makeArrays(fields, symbols, params)
    If b Then
        url = makeYQLRequest(symbols, fields)
        Set yqlReq = New XMLHTTP60
        yqlReq.onreadystatechange = Me
        yqlReq.Open "GET", url, True
        yqlReq.send
    End If
    IAsyncWSFun_makeRequest = True
End Function

Public Function IAsyncWSFun_processResponse(ByRef strRes As String, ParamArray p() As Variant) As Boolean
    On Error GoTo errhandler
    ' Extract input into Array in case of ParamArray compositions
    Dim params() As Variant
    params = Util.formatPArray(p(0))
    ' Find position in table - read xml output for structure
    Dim ij As Variant, i As Integer
    For i = 1 To UBound(params)
        ij = mapToGrid.Item(params(0) & " " & params(i))
        strRes = strRes & yqlDoc.SelectNodes("//results/quote[" & ij(0) + 1 & "]/" & params(i)).Item(0).Text & ";;"
    Next i
    strRes = Left(strRes, Len(strRes) - 2)
    IAsyncWSFun_processResponse = True
    Exit Function
errhandler:
    strRes = "#N/A"
   IAsyncWSFun_processResponse = True ' setting to false only if we know we are calculating later
End Function

Public Function IAsyncWSFun_validateRequest(params As Scripting.IDictionary) As Boolean
    ' TODO add validation
    IAsyncWSFun_validateRequest = True
End Function

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Core processing
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Function makeArrays(ByRef symbols As String, ByRef fields As String, ByRef params As Dictionary) As Boolean
    Dim p As Variant, str As Variant
    Dim dupe As Boolean, st() As String
    Dim countT As Integer, countF As Integer
    For Each p In params.Items
        p = Util.formatPArray(p)
        st = Split(symbols, "%22%2C%20%22")
        countT = 0
        countF = 0
        If Len(symbols) > 0 Then
            While StrComp(st(countT), p(0)) <> 0 And countT < UBound(st)
                countT = countT + 1
            Wend
            If countT = UBound(st) Then
                symbols = symbols & p(0) & "%22%2C%20%22"
            End If
        Else
            symbols = p(0) & "%22%2C%20%22"
        End If
        Dim it As Integer
        For it = 1 To UBound(p)
            st = Split(fields, "%2C%20")
            If Len(fields) > 0 Then
                While StrComp(st(countF), p(it)) <> 0 And countF < UBound(st)
                    countF = countF + 1
                Wend
                If countF = UBound(Split(fields, "%2C%20")) Then
                    fields = fields & p(it) & "%2C%20"
                End If
            Else
                fields = fields & p(it) & "%2C%20"
            End If
            If Not mapToGrid Is Nothing Then
                mapToGrid.Item(p(0) & " " & p(it)) = Array(countT, countF)
            Else
                Set mapToGrid = New Dictionary
                mapToGrid.Item(p(0) & " " & p(it)) = Array(countT, countF)
            End If
        Next it
    Next p
    symbols = Left(symbols, Len(symbols) - 12)
    fields = Left(fields, Len(fields) - 6)
    makeArrays = True
End Function

Private Function makeYQLRequest(ByRef fields As String, ByRef symbols As String) As String
    makeYQLRequest = "https://query.yahooapis.com/v1/public/yql?q=select%20"
    makeYQLRequest = makeYQLRequest & fields
    makeYQLRequest = makeYQLRequest & "%20from%20yahoo.finance.quotes%20where%20symbol%20in%20(%22"
    makeYQLRequest = makeYQLRequest & symbols
    makeYQLRequest = makeYQLRequest & "%22)&diagnostics=true&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys"
    'Debug.Print makeYQLRequest
End Function

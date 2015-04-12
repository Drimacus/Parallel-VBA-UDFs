VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBVlookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*
'' Copyright (c) 2015 Michel Verlinden
'' license: MIT (http://www.opensource.org/licenses/mit-license.php)
'' https://github.com/MichelVerlinden/Parallel-VBA-UDFs
''
'' DBVlookup (database vlookup)
'' Example of SQL SELECT asynchrous function for worksheet purpose
''
'' Dependencies : ActiveX Data Objects
''
'' author : Michel Verlinden
'' 17/03/2014
''
''
'' TODO :   Add connection event handlers
''          Add generic type handler
''          Add test cases
''
'*/

Option Explicit
Implements IAsyncWSFun

Private Const mDebugger = True

' MySQL
'Private Const connString = "Provider='SQLOLEDB';Data Source=<server>;" & _
 "Initial Catalog=<db>;Integrated Security='SSPI';"

' Oracle - Microsoft driver - ODBC
'Private Const connString = "driver={Microsoft ODBC for Oracle};server=<server>;" & _
                                "uid=userid;pwd=userpassword"
' MS SQL Server (OLEDB)
'Private Const connString = "Provider=sqloledb;Data Source=<server>;" & _
                            "Initial Catalog=db;User Id=userid;Password=userpassword"
' etc...
' Excel Workbook
'Private Const connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
                                <path to workbook> & _
                                "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
                                
Private connString As String
Private Const dbTable = "[DBTable$]" ' For example purpose

Private cn As ADODB.Connection
Public WithEvents respH As ADODB.Recordset
Attribute respH.VB_VarHelpID = -1

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

Private Function IAsyncWSFun_catchResp() As Variant
    ' the record set has been initialized in event handler -> just need to callback
    AsynchWSFun.processResp Me
    ' Close connections
    respH.Close
    cn.Close
    Set respH = Nothing
    Set cn = Nothing
End Function

Private Function IAsyncWSFun_getName() As String
    IAsyncWSFun_getName = "Vlookup"
End Function

Private Function IAsyncWSFun_makeRequest(params As Scripting.IDictionary) As Boolean
    On Error GoTo errhandler
    If Debugger.debugging And mDebugger And True Then
        logFunctionCall "IAsyncWSFun_makeRequest(params As Scripting.IDictionary)", params
    End If
    IAsyncWSFun_makeRequest = False
    Set cn = New ADODB.Connection
    Set respH = New ADODB.Recordset
    cn.Open connString
    Dim sqlQuery As String
    sqlQuery = makeSQLSelect(params)
    
    With respH
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .Properties("Initial Fetch Size") = 1
        .Properties("Background Fetch Size") = 9
    End With
    respH.Open sqlQuery, cn, adOpenKeyset, adLockOptimistic, adAsyncFetch
    IAsyncWSFun_makeRequest = True
Exit Function
errhandler:
    IAsyncWSFun_makeRequest = False
End Function

Private Function IAsyncWSFun_processResponse(ByRef strRes As String, ParamArray p() As Variant) As Boolean
    On Error GoTo errhandler
    IAsyncWSFun_processResponse = True
    If respH Is Nothing Then
        strRes = "#Data not fetched"
    Else
        respH.Filter = colStr(p(0)(1)) & "=" & valStr(p(0)(0))
        strRes = respH.fields(p(0)(2))
    End If
Exit Function

errhandler:
    strRes = "#N/A"
End Function

Private Function IAsyncWSFun_validateRequest(params As Scripting.IDictionary) As Boolean
    '***************************************************
    ' TESTING
    ' comment the below if connString initialized as Const
    '***************************************************
    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & ThisWorkbook.FullName & _
                                ";Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    '***************************************************
    '***************************************************
    
    IAsyncWSFun_validateRequest = True
End Function

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Core processing
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Function makeSQLSelect(params As Dictionary) As String
    Dim p As Variant, outputCols As String, filterCols As Dictionary
    Set filterCols = New Dictionary
    For Each p In params.Items
        If InStr(outputCols, colStr(p(2)) & ",") = 0 Then
            outputCols = outputCols & colStr(p(2)) & ","
        End If
        If InStr(outputCols, colStr(p(1)) & ",") = 0 Then
            outputCols = outputCols & colStr(p(1)) & ","
        End If
        If filterCols.Exists(colStr(p(1))) Then
            If InStr(filterCols.Item(colStr(p(1))), valStr(p(0)) & ",") = 0 Then
                filterCols.Item(colStr(p(1))) = filterCols.Item(colStr(p(1))) & valStr(p(0)) & ","
            End If
        Else
            filterCols.Add colStr(p(1)), valStr(p(0)) & ","
        End If
    Next p
    Dim filters As String
    For Each p In filterCols.Keys
        filters = filters & " OR " & p & " IN(" & Left(filterCols.Item(p), Len(filterCols.Item(p)) - 1) & ")"
    Next p
    outputCols = Left(outputCols, Len(outputCols) - 1)
    makeSQLSelect = Replace("SELECT " & outputCols & " FROM " & dbTable & " WHERE" & filters, "WHERE OR", "WHERE")
    Debug.Print makeSQLSelect
End Function

Private Function colStr(ByVal col As String) As String
    colStr = "[" & col & "]"
End Function

Private Function valStr(ByVal val As String) As String
    valStr = val
    'valStr = Chr$(34) & val & Chr$(34)
End Function

Private Sub respH_FetchComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If (adStatus = adStatusOK) Then
        Call IAsyncWSFun_catchResp
    ElseIf (cn.State = ObjectStateEnum.adStateOpen) Then
        cn.Close
        Set cn = Nothing
        Set respH = Nothing
        AsynchWSFun.processResp Me
    End If
End Sub

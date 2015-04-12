VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Switch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*
'' Copyright (c) 2015 Michel Verlinden
'' license: MIT (http://www.opensource.org/licenses/mit-license.php)
'' https://github.com/MichelVerlinden/Parallel-VBA-UDFs
''
'' Switch
'' This class allows to set a timed callback into Excel from a VBA UDF
''
'' author : Michel Verlinden
'' 17/03/2015
''
'' Example for asynchronous xmlHttp well described in
'' http://www.tushar-mehta.com/publish_train/xl_vba_cases/vba_web_pages_services/
'*/

Option Explicit

Public dummyReq As MSXML2.XMLHTTP60

' Default method of this class
Public Sub setTimedCallback()
Attribute setTimedCallback.VB_UserMemId = 0
    If Not AsynchWSFun.executed And dummyReq.ReadyState = 4 Then
        Application.OnTime Now + TimeValue("00:00:01"), _
            "'AsynchWSFun.timedThread'"
    End If
End Sub

' send request upon intialization
Private Sub Class_Initialize()
    AsynchWSFun.executed = False
    Set dummyReq = New MSXML2.XMLHTTP60
    dummyReq.onreadystatechange = Me
    dummyReq.Open "GET", "http://www.DUMMYURL.zzz", True
    dummyReq.send
End Sub

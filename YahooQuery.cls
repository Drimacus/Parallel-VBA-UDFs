VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "YahooQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*
'' Copyright (c) 2015 Michel Verlinden
'' license: MIT (http://www.opensource.org/licenses/mit-license.php)
'' https://github.com/MichelVerlinden/Parallel-VBA-UDFs
''
'' YahooQuery
'' Send requests for Yahoo finance data using by navigating website
''
'' Dependencies: MSXML2
''
'' This module is just an example of how to send requests in multiple steps
'' It is too inefficient for worksheet usage
'' TODO: find better example website to scrape
''
'' author : Michel Verlinden - migul.verlinden@gmail.com
'' 13/03/2014
''
'*/

Option Explicit
Implements IAsyncWSFun

Public xmlReq As MSXML2.XMLHTTP60
Public htmlDoc As MSHTML.HTMLDocument

Private tickers As String
Private idReq As Integer

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Implementation of IAsyncWSFun
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Private Property Let IAsyncWSFun_id(ByVal RHS As Integer)
    idReq = RHS
End Property

Private Property Get IAsyncWSFun_id() As Integer
    IAsyncWSFun_id = idReq
End Property

Public Function IAsyncWSFun_catchResp() As Variant
Attribute IAsyncWSFun_catchResp.VB_UserMemId = 0
    If xmlReq.ReadyState = 4 Then
        Debug.Print "Response Received: " & xmlReq.Status
        Set htmlDoc = New MSHTML.HTMLDocument
        htmlDoc.body.innerHTML = xmlReq.responseText
        If Len(tickers) > 2 Then
            sendOneRequest
            AsynchWSFun.processResp Me
        Else
            AsynchWSFun.processResp Me
            AsynchWSFun.killRequest (idReq)
        End If
    End If
End Function

Public Function IAsyncWSFun_getName() As String
    IAsyncWSFun_getName = "Yahoo"
End Function

Public Function IAsyncWSFun_makeRequest(params As Scripting.IDictionary) As Boolean
    Dim url As String, c As Variant
    For Each c In params.Items
        If InStr(tickers, c(0)) = 0 Then ' TODO make this a better catch of suffix/prefix
            tickers = tickers & c(0) & ";@"
        End If
    Next
    sendOneRequest
    IAsyncWSFun_makeRequest = True
End Function

Public Function IAsyncWSFun_processResponse(ByRef strRes As String, ParamArray p() As Variant) As Boolean
    If InStr(htmlDoc.getElementById("yfi_investing_nav").outerHTML, "symbol=" & p(0)(0) & ";") = 0 Then
        IAsyncWSFun_processResponse = False
    Else
        Dim tbElem As IHTMLElement, info As String
        info = p(0)(1)
        Set tbElem = htmlDoc.getElementById("table1")
        Dim elemsArr() As String
        elemsArr = Split(tbElem.innerText, vbNewLine)
        Dim field() As String, found As Boolean, str As Variant
        found = False
        For Each str In elemsArr
            field = Split(str, ":")
            If StrComp(field(0), info) = 0 And UBound(field) = 1 Then
                found = True
                strRes = field(1)
            End If
        Next str
        If Not found Then
            Set tbElem = htmlDoc.getElementById("table2")
            For Each str In elemsArr
                field = Split(str, ":")
                If StrComp(field(0), info) = 0 And UBound(field) = 1 Then
                    found = True
                    strRes = field(1)
                End If
            Next str
        End If
        If Not found Then
            strRes = "Not found"
        End If
        IAsyncWSFun_processResponse = True
    End If
End Function

Public Function IAsyncWSFun_validateRequest(params As Scripting.IDictionary) As Boolean
    IAsyncWSFun_validateRequest = True
End Function

Private Sub sendOneRequest()
    Dim url As String
    url = "https://uk.finance.yahoo.com/q?s="
    url = url & Left(tickers, InStr(tickers, ";@") - 1)
    url = url & "&ql=0"
    tickers = Right(tickers, Len(tickers) - InStr(tickers, ";@") - 1)
    Set xmlReq = New MSXML2.XMLHTTP60
    xmlReq.onreadystatechange = Me
    xmlReq.Open "GET", url, True
    xmlReq.send
    Debug.Print "Request Sent: " & url
End Sub

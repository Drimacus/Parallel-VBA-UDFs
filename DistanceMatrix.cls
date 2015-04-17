VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DistanceMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*
'' Copyright (c) 2015 Michel Verlinden
'' license: MIT (http://www.opensource.org/licenses/mit-license.php)
'' https://github.com/MichelVerlinden/Parallel-VBA-UDFs
''
'' DistanceMatrix (Google)
'' Calculates the distance between two given addresses by querying
''
'' Dependencies : MSXML2
''
'' author : Michel Verlinden
'' 17/03/2014
''
'' TODO: Split the request into multiple requests of size MAX_REQ
''
'*/

Option Explicit
Implements IAsyncWSFun

' Let me know if you want IP to be permissioned for use with the below key
' This key allows only a limited ammount of requests

Private Const API_KEY = ""
Private Const MAX_REQ = 100

Public xmlReq As MSXML2.XMLHTTP60
Public xmlDoc As MSXML2.DOMDocument60

Private mapToGrid As Dictionary

Private tos As String
Private froms As String

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
 
'Default method of this class
Public Function IAsyncWSFun_catchResp() As Variant
Attribute IAsyncWSFun_catchResp.VB_UserMemId = 0
    If xmlReq.ReadyState = 4 Then
        Set xmlDoc = New MSXML2.DOMDocument60
        If xmlDoc.LoadXML(xmlReq.responseText) Then
            AsynchWSFun.processResp Me
        End If
    End If
End Function

Public Function IAsyncWSFun_getName() As String
    IAsyncWSFun_getName = "ShortestPath"
End Function

Public Function IAsyncWSFun_makeRequest(params As Scripting.IDictionary) As Boolean
    Dim url As String
    url = makeDistRequest(tos, froms)
    Set xmlReq = New XMLHTTP60
    xmlReq.onreadystatechange = Me
    xmlReq.Open "GET", url, True
    xmlReq.send
    IAsyncWSFun_makeRequest = True
End Function

Public Function IAsyncWSFun_processResponse(ByRef strRes As String, ParamArray p() As Variant) As Boolean
    On Error GoTo errhandler
    Dim ij As Variant
    ij = mapToGrid.Item(Join(p(0)))
    strRes = xmlDoc.SelectNodes("/DistanceMatrixResponse/row[" & ij(0) + 1 _
                                    & "]/element[" & ij(1) + 1 _
                                    & "]/distance/text").Item(0).Text
    IAsyncWSFun_processResponse = True
    Exit Function
errhandler:
    strRes = "#N/A"
    IAsyncWSFun_processResponse = True
End Function

Public Function IAsyncWSFun_validateRequest(params As Scripting.IDictionary) As Boolean
    If Len(API_KEY) = 0 Then
        MsgBox "API_KEY not set in class module 'DistanceMatrix'"
        IAsyncWSFun_validateRequest = False
    Else
        If makeArrays(tos, froms, params) Then
            If mapToGrid.Count < MAX_REQ Then ' not sure of the actual limit on this key
                IAsyncWSFun_validateRequest = True
            Else
                MsgBox "Too much data for this API_KEY"
                IAsyncWSFun_validateRequest = False
            End If
        Else
            ' TODO add error handler
            IAsyncWSFun_validateRequest = False
        End If
    End If
End Function

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Core processing
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

Private Function makeArrays(ByRef tos As String, ByRef froms As String, ByRef params As Dictionary) As Boolean
    tos = vbNullString
    froms = vbNullString
    Dim p As Variant, str As Variant
    Dim dupe As Boolean, st() As String
    Dim countT As Integer, countF As Integer
    For Each p In params.Items
        st = Split(tos, "|")
        countT = 0
        countF = 0
        If Len(tos) > 0 Then
            While StrComp(st(countT), urlEncoded(p(0))) <> 0 And countT < UBound(st)
                countT = countT + 1
            Wend
            If countT = UBound(st) Then
                tos = tos & urlEncoded(p(0)) & "|"
            End If
        Else
            tos = urlEncoded(p(0)) & "|"
        End If
        st = Split(froms, "|")
        If Len(froms) > 0 Then
            While StrComp(st(countF), urlEncoded(p(1))) <> 0 And countF < UBound(st)
                countF = countF + 1
            Wend
            If countF = UBound(Split(froms, "|")) Then
                froms = froms & urlEncoded(p(1)) & "|"
            End If
        Else
            froms = froms & urlEncoded(p(1)) & "|"
        End If
        If Not mapToGrid Is Nothing Then
            mapToGrid.Item(Join(p)) = Array(countT, countF)
        Else
            Set mapToGrid = New Dictionary
            mapToGrid.Item(Join(p)) = Array(countT, countF)
        End If
    Next p
    tos = Left(tos, Len(tos))
    froms = Left(froms, Len(froms))
    makeArrays = True
End Function

Private Function makeDistRequest(ByRef tos As String, ByRef froms As String) As String
    makeDistRequest = "https://maps.googleapis.com/maps/api/distancematrix/xml?origins="
    makeDistRequest = makeDistRequest & tos
    makeDistRequest = Left(makeDistRequest, Len(makeDistRequest) - 1) & "&destinations="
    makeDistRequest = makeDistRequest & froms
    makeDistRequest = Left(makeDistRequest, Len(makeDistRequest) - 1) & "&key="
    makeDistRequest = makeDistRequest & API_KEY
End Function

' https://en.wikipedia.org/wiki/Percent-encoding
' Except space needs to be +
Private Function urlEncoded(ByVal url As String) As String
    urlEncoded = False
    Dim chrs() As String, hexd() As String, i As Integer
    chrs = Split("!.#.$.&.'.(.).*.+.,./.:.;.=.?.@.[.]. ", ".")
    hexd = Split("%21;%23;%24;%26;%27;%28;%29;%2A;%2B;%2C;%2F;%3A;%3B;%3D;%3F;%40;%5B;%5D;+", ";")
    For i = 0 To UBound(chrs)
        url = Replace(url, chrs(i), hexd(i))
    Next i
    urlEncoded = url
End Function

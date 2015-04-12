VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Computation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*
'' Copyright (c) 2015 Michel Verlinden
'' license: MIT (http://www.opensource.org/licenses/mit-license.php)
'' https://github.com/MichelVerlinden/Parallel-VBA-UDFs
''
'' Computation
'' This Object is a handle on a set of cells that are
'' to be computed simultenuously - currently a dictionnary maps
'' the calling cells to their arguments
''
'' author : Michel Verlinden
'' 13/03/2015
''
'' TODO :   Change data representation to 2 strings or arrays and test performance
''          R1C1 encodings
''
''
'*/

Option Explicit

Private pf As IAsyncWSFun
Private pCalc As Dictionary
Private pClosed As Boolean

Public Property Set fType(ByRef f As IAsyncWSFun)
    Set pf = f
End Property

Public Property Get fType() As IAsyncWSFun
    Set fType = pf
End Property

Public Property Set calcRng(ByRef d As Dictionary)
    Set pCalc = d
End Property

Public Property Get calcRng() As Dictionary
    Set calcRng = pCalc
End Property

Public Property Let closed(ByVal b As Boolean)
    pClosed = b
End Property

Public Property Get closed() As Boolean
    closed = pClosed
End Property

Public Function addCell(ByRef c As Range, ParamArray p() As Variant) As Boolean
    If pCalc Is Nothing Then
        Set pCalc = New Dictionary
        pCalc.Add c, p(0)
    Else
        pCalc.Add c, p(0)
    End If
    addCell = True
End Function

Public Sub killBatch()
    Set pCalc = Nothing
    Set pf = Nothing
End Sub

Public Sub cancelCalc()
    Dim c As Range, b As Boolean
    b = False
    For Each c In Me.calcRng.Keys
        If Not b Then
            b = True
            c.Value = "#Invalid request"
        Else
            c.Clear
        End If
    Next c
End Sub

Private Sub Class_Initialize()
    Set pCalc = New Dictionary
    Set pf = New IAsyncWSFun
    pClosed = False
End Sub

Private Sub Class_Terminate()
    Set pCalc = Nothing
    Set pf = Nothing
End Sub

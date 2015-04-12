Attribute VB_Name = "Util"
'/*
'' Copyright (c) 2015 Michel Verlinden
'' license: MIT (http://www.opensource.org/licenses/mit-license.php)
'' https://github.com/MichelVerlinden/Parallel-VBA-UDFs
''
'' Util
''
'' author : Michel Verlinden - migul.verlinden@gmail.com
'' 13/03/2014
''
'*/

Option Explicit
Option Private Module

' Generate arrays from strings using below separator
Public Const arrSep = "¬£"

' Iterate over a ParamArray to put all variables in one d
Public Function formatPArray(ParamArray p() As Variant) As Variant
    ' Extract input into Array in case of ParamArray compositions
    Dim params(), it1, it2 As Variant
    ReDim params(0)
    For Each it1 In p(0)
        If IsArray(it1) Then
            For Each it2 In it1
                params(UBound(params)) = it2
                ReDim Preserve params(UBound(params) + 1)
            Next it2
        Else
            params(UBound(params)) = it1
            ReDim Preserve params(UBound(params) + 1)
        End If
    Next it1
    ReDim Preserve params(UBound(params) - 1)
    formatPArray = params
End Function

' Clean the datasheet
'Callback for tableButton onAction
Sub VBA_clearcache(control As IRibbonControl)
    ThisWorkbook.Sheets("Data@Download").Cells.Clear
    AsynchWSFun.calculating = False
    AsynchWSFun.executed = False
End Sub

' If the data sheet is not present we must add with below
Public Function checkDataSheet(ByRef sh As Worksheet) As Boolean
    On Error GoTo errhandler
    Set sh = ThisWorkbook.Sheets("Data@Download")
    checkDataSheet = True
endf:
    Exit Function
errhandler:
   checkDataSheet = False
   Resume endf
End Function

Public Sub addDataSheet(ByRef sh As Worksheet)
    Dim bSU As Boolean
    bSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    With ThisWorkbook.Sheets.Add
        .Name = "Data@Download"
        .Visible = -1
    End With
    Set sh = ThisWorkbook.Sheets("Data@Download")
    Application.ScreenUpdating = bSU
End Sub

Attribute VB_Name = "Debugger"
'/*
'' Copyright (c) 2015 Michel Verlinden
'' license: MIT (http://www.opensource.org/licenses/mit-license.php)
'' https://github.com/MichelVerlinden/Parallel-VBA-UDFs
''
'' Debugger module
'' There are 4 separate threads at least executing per ws function call
'' (1) the ws udf
'' (2) the call back in Switch class
'' (3) the timed callback
'' (4) the callback initiated upon data reception (one for each asynchonous request made in (3))
'' This number of threads increases with the number of asynchronous requests that
'' a computation can make. This module helps to keep track of how the
'' program runs after an end user does "autocomplete" with ws formulas.
''
'' Debugging is controlled via 3 controls
'' below global : debugging
'' module level : mModule boolean
'' procedure    : true/false
''
'' author : Michel Verlinden - migul.verlinden@gmail.com
'' 17/03/2014
''
'*/

Option Explicit
Option Private Module

Public Const debugging = False ' set to true to enable debugging
Private filenum As Integer

Public Sub logFunctionCall(fName As String, ParamArray p() As Variant)
    On Error GoTo errhandler
    If Not openLog(filenum) Then
        Exit Sub
    End If
    Dim paramDescr() As String
    ReDim paramDescr(0)
    paramDescr(0) = Now & vbNewLine & _
                    "Executing: " & fName & vbNewLine & "With Parameters:"
    Dim i As Integer
    For i = LBound(p) To UBound(p)
        describeArgument p(i), paramDescr
    Next i
    ' print to file
    Dim s As Variant
    For Each s In paramDescr
        Print #filenum, s
    Next s
    Print #filenum, "--------------------------------------------------------------------------"
    Exit Sub
errhandler:
    stopDebugging
End Sub

Private Sub describeArgument(ByVal v As Variant, ByRef paramDescr() As String)
    If IsArray(v) Then
        Dim it As Variant
        For Each it In v
            describeArgument it, paramDescr
        Next it
    Else ' TODO: add type descriptions
        ReDim Preserve paramDescr(UBound(paramDescr) + 1)
        If IsObject(v) Then
            If v Is Nothing Then
                paramDescr(UBound(paramDescr)) = "Nothing"
            ElseIf TypeOf v Is IAsyncWSFun Then
                paramDescr(UBound(paramDescr)) = printWSF(v)
            ElseIf TypeOf v Is Dictionary Then
                paramDescr(UBound(paramDescr)) = printDico(v)
            Else
                paramDescr(UBound(paramDescr)) = TypeName(v)
            End If
        Else
            paramDescr(UBound(paramDescr)) = printVariable(v)
        End If
    End If
End Sub

Public Sub logThis(str As String)
    On Error GoTo errhandler
    If Not openLog(filenum) Then
        Exit Sub
    End If
    Print #filenum, "--------------------------------------------------------------------------"
    Print #filenum, str
    Print #filenum, "--------------------------------------------------------------------------"
errhandler:
    stopDebugging
End Sub

Public Sub stopDebugging()
    On Error Resume Next
    Close #filenum
    filenum = 0
End Sub

Private Function openLog(ByRef n As Integer) As Boolean
    On Error GoTo errhandler
    openLog = False
    If n = 0 Then
        n = FreeFile
        Open ThisWorkbook.Path & "\log.txt" For Output As #n
    End If
    openLog = True
    Exit Function
errhandler:
    Resume Next
End Function

Private Function printDico(ByVal d As Dictionary) As String
    On Error GoTo errhandler
    printDico = "-----" & vbNewLine & _
                "Dictionary" & vbNewLine & _
                "Size:" & d.Count & vbNewLine & "-----"
    Exit Function
errhandler:
    printDico = "Print not available"
    Resume Next
End Function

Private Function printWSF(ByVal f As IAsyncWSFun) As String
    On Error GoTo errhandler
    printWSF = "-----" & vbNewLine & _
                "Batch ID: " & f.id & vbNewLine & _
                "Name: " & f.getName & vbNewLine & "-----"
    Exit Function
errhandler:
    printWSF = "Print not available"
    Resume Next
End Function

Private Function printVariable(ByVal v As Variant) As String
    On Error GoTo errhandler
    printVariable = CStr(v)
    Exit Function
errhandler:
    printVariable = "Print not available"
    Resume Next
End Function


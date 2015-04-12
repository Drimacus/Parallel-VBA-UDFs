VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAsyncWSFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/*
'' Copyright (c) 2015 Michel Verlinden
'' license: MIT (http://www.opensource.org/licenses/mit-license.php)
'' https://github.com/MichelVerlinden/Parallel-VBA-UDFs
''
'' IAsyncWSFun
'' Implement this interface to define a worksheet function that can compute simultaneously
'' with other cells
''
'' author : Michel Verlinden
'' 17/03/2015
''
''
''
'*/

Option Explicit

' request id - implementation must have GET and LET
Public id As Integer

' getName : Return a unique identifying name - if the name is not unique
'           the manager module will confuse data with that of another function
Public Function getName() As String
End Function
 
' makeRequest:  This is the function responsible for sending a bundled data request to web-services and databases
'
' Parameter:    A dictionnary consisting of all caller cells/arrays : (Key, Item) = (Range, Function Arguments (As String))
'               Iterate over the dictionary to generate a bundled SQL, xmlhttp, SOAP, KDB/q ...
'               query for all calling cells of the calculation
' Return:       True if request successfully sent
Public Function makeRequest(params As Dictionary) As Boolean
End Function

' validateRequest:  Determine if the request can be made for data in dictionnary
'
' Parameter:        A dictionnary consisting of all caller cells/arrays : (Key, Item) = (Range, Function Arguments (As String))
'
' Return:           True if the request can be made for params
'                   Typically one should refuse unreasonable calculations ~ 1000000+ data points
'                   But it is up to implementation to know what is too much
'                   Alternative is to break it down in multiple requests - See YahooQuery Class
Public Function validateRequest(params As Dictionary) As Boolean
End Function

' processResponse:  In this function one needs to implement how the function assigns value to 1 calling range
'                   based on response received. Implementation is usually easier to store response of makeRequest in module
'                   level variable.
'                   Note:   put the result of f(p()) into strRes
'                           Return false if response is not enough to calculate p() yet -> the cell will be calculated at a later stage
'                           This allows to send multiple requests and compute "response" in multiple steps
' Parameters:       strRes:     result of applying the worksheet function to ParamArray p()
'                   p:          the arguments given to the worksheet functions
'
' Return:           True if strRes has been computed and the calling range should calculate
Public Function processResponse(ByRef strRes As String, ParamArray p() As Variant) As Boolean
End Function

' catchResp:    This function MUST call : AsynchWSFun.processResp Me
'               where response is data created in this function
'               The data should also be loaded in a module variable for later processing
'               Set this method as default to handle asynchronous xmlhttp
'               Call this method from an event handler to manage ADO fetch complete events
Public Function catchResp()
    'AsynchWSFun.processResp Me
End Function

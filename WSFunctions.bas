Attribute VB_Name = "WSFunctions"
'/*
'' WSFunctions
'' Module where worksheet functions should be defined
''
'' In order to define "parralel" function: Implement IAsynchWSFun
'' in a class module and use the object in AsychWSFun.asyncFun(<your object>, <your parameters>)
''
''
'' author : Michel Verlinden
'' 17/03/2014
''
'' TODO :   Add generic argument validator
''          Function registration
''
'*/
Option Explicit
'Option Private Module ' comment this if not registering functions

' drivingDistance
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function drivingDistance(Origin As String, Destination As String)
    Dim f As IAsyncWSFun
    Set f = New DistanceMatrix
    drivingDistance = AsynchWSFun.asyncFun(f, Origin, Destination)
End Function

' testYahoo
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function testYahoo(s As String, field As String)
    Dim f As IAsyncWSFun
    Set f = New YahooQuery
    testYahoo = AsynchWSFun.asyncFun(f, s, field)
End Function

' YDP
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function YDP(symbol As String, ParamArray p() As Variant) As Variant
    ' Validate arguments
    Dim possCols As String, possSymbolChars As String
    possSymbolChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.-1234567890"
    possCols = "Ask,AverageDailyVolume,Bid,AskRealtime,BidRealtime,BookValue,Change&PercentChange,Change,Commission," & _
                "Currency,ChangeRealtime,AfterHoursChangeRealtime,DividendShare,LastTradeDate,TradeDate,EarningsShare," & _
                "ErrorIndicationreturnedforsymbolchangedinvalid,EPSEstimateCurrentYear,EPSEstimateNextYear,EPSEstimateNextQuarter," & _
                "DaysLow,DaysHigh,YearLow,YearHigh,HoldingsGainPercent,AnnualizedGain,HoldingsGain,HoldingsGainPercentRealtime," & _
                "HoldingsGainRealtime,MoreInfo,OrderBookRealtime,MarketCapitalization,MarketCapRealtime,EBITDA,ChangeFromYearLow," & _
                "PercentChangeFromYearLow,LastTradeRealtimeWithTime,ChangePercentRealtime,ChangeFromYearHigh,PercebtChangeFromYearHigh," & _
                "LastTradeWithTime,LastTradePriceOnly,HighLimit,LowLimit,DaysRange,DaysRangeRealtime,FiftydayMovingAverage," & _
                "TwoHundreddayMovingAverage,ChangeFromTwoHundreddayMovingAverage,PercentChangeFromTwoHundreddayMovingAverage," & _
                "ChangeFromFiftydayMovingAverage,PercentChangeFromFiftydayMovingAverage,Name,Notes,Open,PreviousClose,PricePaid," & _
                "ChangeinPercent,PriceSales,PriceBook,ExDividendDate , PERatio, DividendPayDate, PERatioRealtime, PEGRatio, " & _
                "PriceEPSEstimateCurrentYear, PriceEPSEstimateNextYear, Symbol, SharesOwned, ShortRatio, LastTradeTime, TickerTrend," & _
                "OneyrTargetPrice, Volume, HoldingsValue, HoldingsValueRealtime, YearRange, DaysValueChange, DaysValueChangeRealtime," & _
                "StockExchange, DividendYield"
    Dim isValid As Boolean, col As Variant
    isValid = True
    For Each col In p
        If InStr(possCols, col & ",") = 0 Then
            isValid = False
        End If
        If Not isValid Or Len(col) = 0 Then
            YDP = "'" & col & "': not a valid field"
        End If
    Next
    Dim i As Integer
    For i = 1 To Len(symbol)
        If InStr(possSymbolChars, Mid$(symbol, i, 1)) = 0 Then
            isValid = False
        End If
    Next
    If Not isValid Or Len(symbol) = 0 Then
        YDP = "'" & symbol & "': not a valid security"
    End If
   ' Query Data
    If isValid Then
        ' fetch data
        Dim f As New YahooAPI, res As String
        res = AsynchWSFun.asyncFun(f, symbol, p)
        If Len(res) > 0 Then
            ' Populate array
            YDP = Split(res, ";;")
        Else
            YDP = vbNullString
        End If
    End If
End Function

' testSQL
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Public Function testSQL(toMatch As String, matchingCol As String, outputCol As String) As String
    If Len(toMatch) = 0 Or Len(matchingCol) = 0 Or Len(outputCol) = 0 Then
        testSQL = "#Invalid arguments"
    Else
        Dim f As DBVLOOKUP
        Set f = New DBVLOOKUP
        testSQL = AsynchWSFun.asyncFun(f, toMatch, matchingCol, outputCol)
    End If
End Function

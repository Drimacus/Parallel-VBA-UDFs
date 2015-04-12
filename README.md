# Parallel-VBA-UDFs
Compute VBA UDFs simultaneously

# Summary

In this project we get a handle on the calculation events that occur in Excel VBA user defined functions. 

By implementing IAsyncWSFun on defines the way a worksheet function should compute knowing what the other cells are calculating. This is particularly useful when functions need to request data over a network. By creating a batch  calculation it is possible to request data to REST APIs/Databases efficiently and compute all the cells based on the same data request.

The interface is designed so that one only needs to worry about fomulating a batch request and catching the response. The manager module takes care of reassigning and calculating all cells that are calling the function. Separately one needs to write the public function that will be called in the worksheets. This function will use the implemented object and call the manager module.

# Examples

## Google Distance Matrix

This worksheet function allows to compute the distance between 2 addresses:

```VB.net
Public Function drivingDistance(origin as String, destination as String) as String
	Dim f As IAsyncWSFun
    	Set f = New DistanceMatrix
    	drivingDistance = AsynchWSFun.asyncFun(f, origin, destination) 
End Function
```

## Yahoo Finance API

This worksheet function does the same as the Bloomberg BDP function. Given a security, output the information of interest about it (Ask, Bid, Volume, ...)

This function outputs an array depending on the number of output parameters specified

```VB.net
Public Function YDP(security as String, info1 as String, info2 as String, ...) as Variant
	Dim f As IAsyncWSFun, res As String
    	Set f = New YahooAPI
	' AsynchWSFun.asyncFun returns String
    	res = AsynchWSFun.asyncFun(f, origin, destination) 
        If Len(res) > 0 Then
            ' Populate array
            YDP = Split(res, ";;")
        Else
            YDP = vbNullString
        End If
End Function
```

## Database VLOOKUP

This worksheet is equivalent to VLOOKUP except that the table queried is in a database

```VB.net
Public Function testSQL(toMatch As String, matchingCol As String, outputCol As String) As String
	Dim f As IAsyncWSFun
    	Set f = New DBVLOOKUP
    	testSQL = AsynchWSFun.asyncFun(f, toMatch, matchingCol, outputCol)
End Function
```

Complete implementations with arguments checking and function registering are in module: WSFunctions



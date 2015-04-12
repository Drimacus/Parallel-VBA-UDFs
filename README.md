# Parallel-VBA-UDFs
Compute VBA UDFs simultaneously

# Summary

In this project we get a handle on the calculation events that occur in Excel VBA user defined functions. 

By implementing IAsyncWSFun on defines the way a worksheet function should compute knowing what the other cells are calculating. This is particularly useful when functions need to request data over a network. By creating a batch  calculation it is possible to request data to REST APIs/Databases efficiently and compute all the cells based on the same data request.

The interface is designed so that one only needs to worry about fomulating a batch request and catching the response. The manager module takes care of reassigning and calculating all cells that are calling the function. Separately one needs to write the public function that will be called in the worksheets. This function will use the implemented object and call the manager module.

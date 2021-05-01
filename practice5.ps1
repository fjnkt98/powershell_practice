# Object[] Type Properties and Methods

$a = @(10, 20, 30)
$a.Count  # Get length of array

$a[0]       # Get element of the array
$a.Item(0)  # Get element of the array (same as above)

# ArrayList Type Properties and Methods
$array = New-Object System.Collections.ArrayList
$array.AddRange(@(10, 20))

$array.Count

$array.Add(30)
$array.Add(40)

$array.Count

$array.Insert(1, 99)
$array
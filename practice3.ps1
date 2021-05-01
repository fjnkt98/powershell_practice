# Declare integer
$a = 1
$a.GetType()

[int]$hoge = 3
$hoge.GetType()

# Declare float
$b = 3.14
$b.GetType()
[double]$moge = 2.73
$moge.GetType()

# Declare string
$str = "Hello, World!"
$str.GetType()

# Booleans
$fuga = $true
$fuga = $false
$fuga.GetType()

# Declare array
$arr = 1, 2, 3
$arr = @(1, 2, 3)
$arr = @(1)
$arr = @()
[int[]]$arr = @(1, 2, 3, 4, 5)

foreach ($item in $arr) {
  Write-Host ([Math]::Pow($item, 2))
}

#Declare hashtable
$dict = @{"hoge"=111; "moge"=122}
[hashtable]$dict = @{"hoge"=111; "moge"=222}
$dict.GetType()

# Where-Object
$array = @(5, 1, 2, 3, 6, 7, 2, 9, 3, 10)
$striped_array = $array | Where-Object { $_ -gt 2 }
Write-Host $striped_array

# ForEach-Object
$array2 = @(0, 1, 2, 3, 4, 5)
$array2 | ForEach-Object { $_ * 2 }

# Select-Object
$array3 = @(13, 32, 81, 1, 45)
$array3 | Select-Object -First 2

# Sort-Object
$array4 = @(22, 3, 140, 57, 84)
$array4 | Sort-Object
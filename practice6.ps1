# Function

function  myfunc {
  param (
    $arg1
  )
  Write-Host $arg1
  $arg1 | Get-Member
}

myfunc("hogehoge")

# PSCustomObject
$a = [PSCustomObject]@{
  "key1" = 1
  "key2" = 2
}
$a.psobject.Properties.Name
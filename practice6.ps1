# Function

function  myfunc {
  param (
    $arg1
  )
  Write-Host $arg1
  $arg1 | Get-Member
}

myfunc("hogehoge")

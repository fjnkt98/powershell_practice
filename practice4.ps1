$a = 114514

$str1 = "hogehoge"
Write-Host $str1
$str2 = "hoge`n`moge"
Write-Host $str2
$str3 = "a is $a"
Write-Host $str3
$str4 = 'a is $a'
Write-Host $str4

$str1.Length

# Substring
$str5 = "abcdef"
$str5.Substring(1, 3)  # インデックス1から3文字分を得る
$str5.Substring(0, 3)  # インデックス0から3文字分を得る

# Trim
$str6 = " a b "
$str7 = $str6.Trim()
Write-Host "[$str6]"
Write-Host "[$str7]"

# Replace
$str8 = "abc"
$str9 = $str8.Replace("b", "HOGE")
Write-Host $str8
Write-Host $str9
###########################
# Training of Get-Content #
###########################

$lines = Get-Content -Path .\practice6.txt

foreach ($items in $lines) {
  Write-Host $items
}
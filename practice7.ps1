$products = Import-Csv .\practice7.csv -Encoding UTF8
$products | Format-Table

#$products[0].product_code

foreach ($item in $products) {
  $item.product_name
}

$products.Count
$filePath = "C:\Users\Par149\AppData\Local\Microsoft\Outlook\Offline Address Books\9f97379f-d3ed-41a0-a0c0-9f50ecc3f3e8\udetails.oab"
$parsePattern = "Microsoft Private MDB"

$foundData = Select-String -Path $filePath -Pattern $parsePattern -AllMatches
$instanceCount = $foundData.Matches.Count
"$instanceCount"
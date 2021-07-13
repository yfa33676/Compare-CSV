# エンコーディング（SJIS）
$OutputEncoding = [console]::OutputEncoding

# ファイル保存ダイアログ
Add-Type -AssemblyName System.Windows.Forms
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog 
$SaveFileDialog.Filter = "csvファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*"
$SaveFileDialog.InitialDirectory = ".\"

# ファイル開くダイアログ
Add-Type -AssemblyName System.Windows.Forms
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog 
$OpenFileDialog.Filter = "csvファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*"
$OpenFileDialog.InitialDirectory = ".\"

Write-Progress -Activity "変更前CSVの読み込み" -Status 読み込み開始
$OpenFileDialog.Filename = "変更前.csv"
if($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
  $AHeader = (Get-Content -Path $OpenFileDialog.Filename)[0]
  Write-Progress -Activity "変更前CSVの読み込み" -Status 読み込み中
  $A = Get-Content -Path $OpenFileDialog.Filename | % Insert 0 "削除,"
  Write-Progress -Activity "変更前CSVの読み込み" -Status 読み込み終了
  $A[0] = $A[0].Replace("削除", "変更区分")
} else{
  return
}

Write-Progress -Activity "変更後CSVの読み込み" -Status 読み込み開始
$OpenFileDialog.Filename = "変更後.csv"
if($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
  $BHeader = (Get-Content -Path $OpenFileDialog.Filename)[0]
  Write-Progress -Activity "変更後CSVの読み込み" -Status 読み込み中
  $B = Get-Content -Path $OpenFileDialog.Filename | % Insert 0 "追加,"
  Write-Progress -Activity "変更後CSVの読み込み" -Status 読み込み終了
  $B[0] = $B[0].Replace("追加", "非表示")
} else{
  return
}

if($AHeader -eq $BHeader){
  $Header = $AHeader.Split(",") | % Replace "`"" "" | % Trim
} else {
  "ヘッダーが一致しません" | Out-Host
  pause
  return
}


Write-Progress "CSVをソート"
$PrimaryKey  = $Header | Out-GridView -PassThru -Title "主キーを選んでください"
$C = $A + $B | ConvertFrom-CSV | Sort-Object ($PrimaryKey + "変更区分") -CaseSensitive

Write-Progress "CSVを比較" -percentComplete 0
$CompKey = $Header | Out-GridView -PassThru -Title "比較キーを選んでください"
for($i = 0; $i -lt $C.length - 1; $i++){
  Write-Progress "CSVを比較" -percentComplete (100 * ($i)/($C.length))
  if($C[$i].変更区分 -eq "削除" -and $C[$i+1].変更区分 -eq "追加"){
    if(
      ($C[$i]   | Select-Object $CompKey | ConvertTo-Json) -eq
      ($C[$i+1] | Select-Object $CompKey | ConvertTo-Json)
    ){
      $C[$i].変更区分 = "非表示"
      $C[$i+1].変更区分 = ""
    } elseif(
      ($C[$i]   | Select-Object $PrimaryKey | ConvertTo-Json) -eq
      ($C[$i+1] | Select-Object $PrimaryKey | ConvertTo-Json)
    ){
      $C[$i].変更区分 = "変更前"
      $C[$i+1].変更区分 = "変更後"
    }
    $i++
  }
}

$C = $C | ? 変更区分 -ne "非表示"
$C | Out-GridView -title "比較結果"

while($true){
  $Return = Read-Host "結果を出力しますか？(はい：Y、いいえ：N)"
  if ($Return -eq "Y"){
    $SaveFileDialog.Filename = "比較結果.csv"
    if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
      $C | Export-Csv -Encoding Default -NoTypeInformation -Path $SaveFileDialog.Filename
    }
  }
  if ($Return -eq "N"){
    break
  }
}

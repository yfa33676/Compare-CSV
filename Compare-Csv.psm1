function Compare-Csv{
  [CmdletBinding()]
  param(
    [string]$ReferencePath = "", # 比較前CSVパス
    [string]$DifferencePath  = "" # 比較後CSVパス
  )
  begin{
    # エンコーディング（SJIS）
    $OutputEncoding = [console]::OutputEncoding

    # ファイルダイアログ
    Add-Type -AssemblyName System.Windows.Forms

    # 開く
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog 
    $OpenFileDialog.Filter = "csvファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*"
    $OpenFileDialog.InitialDirectory = ".\"

    # 保存
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog 
    $SaveFileDialog.Filter = "csvファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*"
    $SaveFileDialog.InitialDirectory = ".\"

  }
  process{
    if ($ReferencePath -eq ""){
      Write-Progress -Activity "変更前CSVの読み込み" -Status 読み込み開始
      $OpenFileDialog.Filename = "変更前.csv"
      if($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $ReferencePath = $OpenFileDialog.Filename
      } else{
        return
      }
    }
    # 見出し行
    $Header = Get-Content -Path $ReferencePath -Head 1

    # 変更前CSV
    Write-Progress -Activity "変更前CSVの読み込み" -Status 読み込み中
    $ReferenceCSV = Get-Content -Path $ReferencePath | Select-Object -Skip 1 | % Insert 0 "削除,"

    if ($DifferencePath -eq ""){
      Write-Progress -Activity "変更後CSVの読み込み" -Status 読み込み開始
      $OpenFileDialog.Filename = "変更後.csv"
      if($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $DifferencePath = $OpenFileDialog.Filename
      } else{
        return
      }
    }

    Write-Progress -Activity "変更後CSVの読み込み" -Status 読み込み中

    # 見出し行チェック
    if ($Header -ne (Get-Content -Path $DifferencePath -Head 1)){
      "ヘッダーが一致しません"
      Read-Host
      break
    }

    # 変更前CSV
    $DifferenceCSV = Get-Content -Path $DifferencePath | Select-Object -Skip 1 | % Insert 0 "追加,"

    # CSVをソート
    Write-Progress "CSVをソート" -Status 主キーの選択
    $PrimaryKey  = $Header -split "," | Out-GridView -PassThru -Title "主キーを選んでください"
    if ($PrimaryKey.count -eq 0) {
      "少なくともひとつは主キーを選んでください" | Out-Host
      
    } elseif($PrimaryKey.count -eq 1) {
      $SortKey = $PrimaryKey,  "変更区分"
      break
    } else {
      $SortKey = $PrimaryKey + "変更区分"
    }
    $MergeData = $Header.Insert(0,"変更区分,"),$ReferenceCSV,$DifferenceCSV | ConvertFrom-CSV | Sort-Object $SortKey -CaseSensitive

    Write-Progress "CSVをソート" -Status 比較キーの選択
    $CompKey = $Header -split "," | Out-GridView -PassThru -Title "比較キーを選んでください"

    for($i = 0; $i -lt $MergeData.length - 1; $i++){
      if($MergeData[$i].変更区分 -eq "削除" -and $MergeData[$i+1].変更区分 -eq "追加"){
        if(
          ($MergeData[$i]   | Select-Object $CompKey | ConvertTo-JSON) -eq
          ($MergeData[$i+1] | Select-Object $CompKey | ConvertTo-JSON)
        ){
          $MergeData[$i].変更区分 = ""
          $MergeData[$i+1].変更区分 = ""
        } elseif(
          ($MergeData[$i]   | Select-Object $PrimaryKey | ConvertTo-JSON) -eq
          ($MergeData[$i+1] | Select-Object $PrimaryKey | ConvertTo-JSON)
        ){
          $MergeData[$i].変更区分 = "変更前"
          $MergeData[$i+1].変更区分 = "変更後"
        }
        $i++
      }
      $Status = [string]$MergeData.length + "件中" + [string]$i + "件完了"
      Write-Progress "CSVを比較" -Status $Status -PercentComplete (100 * ($i)/($MergeData.length))
    }
    $Status = [string]$MergeData.length + "件中" + [string]$MergeData.length + "件完了"
    Write-Progress "CSVを比較" -Status $Status -PercentComplete (100 * ($MergeData.length)/($MergeData.length))
    $MergeData = $MergeData | Sort-Object $SortKey -CaseSensitive -Unique
    $MergeData | Out-GridView -title ("【変更前】" + ($ReferencePath | Split-Path -Leaf) + " - 【変更後】" + ($DifferencePath | Split-Path -Leaf))

    while($true){
      $Return = Read-Host "結果を出力しますか？(はい：Y、いいえ：N)"
      if ($Return -eq "Y"){
        $SaveFileDialog.Filename = "比較結果.csv"
        if ($SaveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
          $MergeData | Export-Csv -Encoding Default -NoTypeInformation -Path $SaveFileDialog.Filename
        }
        break
      }
      if ($Return -eq "N"){
        break
      }
    }
  }
  end{
  }
}

function Compare-Csv{
  [CmdletBinding()]
  param(
    [string]$ReferencePath = "", # ��r�OCSV�p�X
    [string]$DifferencePath  = "" # ��r��CSV�p�X
  )
  begin{
    # �G���R�[�f�B���O�iSJIS�j
    $OutputEncoding = [console]::OutputEncoding

    # �t�@�C���_�C�A���O
    Add-Type -AssemblyName System.Windows.Forms

    # �J��
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog 
    $OpenFileDialog.Filter = "csv�t�@�C��(*.csv)|*.csv|���ׂẴt�@�C��(*.*)|*.*"
    $OpenFileDialog.InitialDirectory = ".\"

    # �ۑ�
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog 
    $SaveFileDialog.Filter = "csv�t�@�C��(*.csv)|*.csv|���ׂẴt�@�C��(*.*)|*.*"
    $SaveFileDialog.InitialDirectory = ".\"

  }
  process{
    if ($ReferencePath -eq ""){
      Write-Progress -Activity "�ύX�OCSV�̓ǂݍ���" -Status �ǂݍ��݊J�n
      $OpenFileDialog.Filename = "�ύX�O.csv"
      if($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $ReferencePath = $OpenFileDialog.Filename
      } else{
        return
      }
    }
    # ���o���s
    $Header = Get-Content -Path $ReferencePath -Head 1

    # �ύX�OCSV
    Write-Progress -Activity "�ύX�OCSV�̓ǂݍ���" -Status �ǂݍ��ݒ�
    $ReferenceCSV = Get-Content -Path $ReferencePath | Select-Object -Skip 1 | % Insert 0 "�폜,"

    if ($DifferencePath -eq ""){
      Write-Progress -Activity "�ύX��CSV�̓ǂݍ���" -Status �ǂݍ��݊J�n
      $OpenFileDialog.Filename = "�ύX��.csv"
      if($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        $DifferencePath = $OpenFileDialog.Filename
      } else{
        return
      }
    }

    Write-Progress -Activity "�ύX��CSV�̓ǂݍ���" -Status �ǂݍ��ݒ�

    # ���o���s�`�F�b�N
    if ($Header -ne (Get-Content -Path $DifferencePath -Head 1)){
      "�w�b�_�[����v���܂���"
      Read-Host
      break
    }

    # �ύX�OCSV
    $DifferenceCSV = Get-Content -Path $DifferencePath | Select-Object -Skip 1 | % Insert 0 "�ǉ�,"

    # CSV���\�[�g
    Write-Progress "CSV���\�[�g" -Status ��L�[�̑I��
    $PrimaryKey  = $Header -split "," | Out-GridView -PassThru -Title "��L�[��I��ł�������"
    if ($PrimaryKey.count -eq 0) {
      "���Ȃ��Ƃ��ЂƂ͎�L�[��I��ł�������" | Out-Host
      
    } elseif($PrimaryKey.count -eq 1) {
      $SortKey = $PrimaryKey,  "�ύX�敪"
      break
    } else {
      $SortKey = $PrimaryKey + "�ύX�敪"
    }
    $MergeData = $Header.Insert(0,"�ύX�敪,"),$ReferenceCSV,$DifferenceCSV | ConvertFrom-CSV | Sort-Object $SortKey -CaseSensitive

    Write-Progress "CSV���\�[�g" -Status ��r�L�[�̑I��
    $CompKey = $Header -split "," | Out-GridView -PassThru -Title "��r�L�[��I��ł�������"

    for($i = 0; $i -lt $MergeData.length - 1; $i++){
      if($MergeData[$i].�ύX�敪 -eq "�폜" -and $MergeData[$i+1].�ύX�敪 -eq "�ǉ�"){
        if(
          ($MergeData[$i]   | Select-Object $CompKey | ConvertTo-JSON) -eq
          ($MergeData[$i+1] | Select-Object $CompKey | ConvertTo-JSON)
        ){
          $MergeData[$i].�ύX�敪 = ""
          $MergeData[$i+1].�ύX�敪 = ""
        } elseif(
          ($MergeData[$i]   | Select-Object $PrimaryKey | ConvertTo-JSON) -eq
          ($MergeData[$i+1] | Select-Object $PrimaryKey | ConvertTo-JSON)
        ){
          $MergeData[$i].�ύX�敪 = "�ύX�O"
          $MergeData[$i+1].�ύX�敪 = "�ύX��"
        }
        $i++
      }
      $Status = [string]$MergeData.length + "����" + [string]$i + "������"
      Write-Progress "CSV���r" -Status $Status -PercentComplete (100 * ($i)/($MergeData.length))
    }
    $Status = [string]$MergeData.length + "����" + [string]$MergeData.length + "������"
    Write-Progress "CSV���r" -Status $Status -PercentComplete (100 * ($MergeData.length)/($MergeData.length))
    $MergeData = $MergeData | Sort-Object $SortKey -CaseSensitive -Unique
    $MergeData | Out-GridView -title ("�y�ύX�O�z" + ($ReferencePath | Split-Path -Leaf) + " - �y�ύX��z" + ($DifferencePath | Split-Path -Leaf))

    while($true){
      $Return = Read-Host "���ʂ��o�͂��܂����H(�͂��FY�A�������FN)"
      if ($Return -eq "Y"){
        $SaveFileDialog.Filename = "��r����.csv"
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

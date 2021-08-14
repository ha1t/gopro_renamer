#
# GoPro Renamer
#

Param([switch]$d,[switch]$r)
$dryRunEnabled = $d
$recurseEnabled = $r

# include config
. ".\config.ps1"

$appVersion = "v1.0.6"
Add-Type -AssemblyName System.Drawing

if (Test-Path $FFMPEG -ne $true)
{
    Write-Host "ffmpeg not found"
    Exit
}

# ffmpeg で creation_time を更新する
# creation_time=2021-08-14T23:51:59.000000Z
function updateCreationTime($fileName, [System.DateTime]$updateDate)
{
    $updateDate -= New-TimeSpan -Hours 9 # すでに引き算しているはずなのに…
    $update_time = $updateDate.ToString("yyyy-MM-ddTHH:mm:ss.000000Z")
    $newFileName = $fileName + ".mp4"

    if ((Test-Path $newFileName) -eq $true)
    {
        return
    }

    $process = Start-Process -FilePath "`"${FFMPEG}`"" -PassThru  -Wait -ArgumentList "-i `"${fileName}`" -c copy -metadata creation_time=`"${update_time}`" `"${newFileName}`""

    if ($process.ExitCode -ne 0)
    {
        Write-Host "Error: " + $fileName
        Exit
    }
}

# Exifから日時文字列を生成する
function getExifDate($path) {
  try {
    $img = New-Object Drawing.Bitmap($path)
  } catch {
    return ""
  }
  $byteAry = ($img.PropertyItems | Where-Object{$_.Id -eq 36867}).Value
  if (!$byteAry) {
    $img.Dispose()
    $img = $null
    return ""
  }

  # "YYYY:MM:DD HH:MM:SS " -> "YYYY/MM/DD HH:MM:SS"
  $byteAry[4] = 47
  $byteAry[7] = 47
  $ret = [System.Text.Encoding]::UTF8.GetString($byteAry)
  $ret = $ret.substring(0, 19)
  $img.Dispose()
  $img = $null

  return $ret
}

# 詳細プロパティから日時文字列を生成する
function getPropDate($folder, $file) {
  $shellFolder = $shellObject.namespace($folder)
  $shellFile = $shellFolder.parseName($file)
  $selectedPropertyNo = ""
  $selectedPropertyName = ""
  $selectedPropertyValue = ""

  for ($i = 0; $i -lt 300; $i++) { # 208まで探せば十分?
    $propertyName = $shellFolder.getDetailsOf($Null, $i)
    if (($propertyName -eq "撮影日時") `
        -or ($propertyName -eq "メディアの作成日時")) {
      $propertyValue = $shellFolder.getDetailsOf($shellFile, $i)
      if ($propertyValue) {
        $selectedPropertyNo = $i
        $selectedPropertyName = $propertyName
        $selectedPropertyValue = $propertyValue
        break
      }
    }
  }
  if (!$selectedPropertyNo) {
    return ""
  }

  # " YYYY/ MM/ DD   H:MM" -> "YYYY/MM/DD HH:MM:00"
  $ret = $selectedPropertyValue
  $time = "0" + $ret.substring(16) + ":00" # 秒は取得できないので00を設定
  $time = $time.substring($time.length - 8, 8)
  $ret = $ret.substring(1, 5) + $ret.substring(7, 3) + $ret.substring(11, 2) + " " + $time
  $date = Get-Date $ret
  $ts = New-TimeSpan -Hours 9
  return ($date - $ts)
}

# ファイル名から日時文字列を生成する
function getFnameDate($file) {
  $ret = ""

  if ($file -match "^([0-9]{4})([0-9]{2})([0-9]{2})\-([0-9]{2})([0-9]{2})([0-9]{2})") {
    # "YYYYMMDD-HH:MM:SS*" -> "YYYY/MM/DD HH:MM:SS"
    $ret = $Matches[1] + "/" + $Matches[2] + "/" + $Matches[3]
    $ret = $ret + " " + $Matches[4] + ":" + $Matches[5] + ":" + $Matches[6]
  }

  return $ret
}

# ファイルスキップ時の表示
function printSkipped($folder, $file) {
  $rfPath = (Resolve-Path $folder -Relative)
  if ($rfPath.StartsWith("..\")) {
    $rfPath = ".\"
  }
  Write-Host "[$rfPath] $file (" -NoNewline
  Write-Host "skipped" -ForegroundColor Red -NoNewline
  Write-Host ")"
}

# メイン処理
function main {
  # バナーを表示
  $mode = if ($dryRunEnabled) { " (dry run)" } else { "" }
  Write-Host "== Media Dater $appVersion$mode =="

  # シェルオブジェクトを生成
  $shellObject = New-Object -ComObject Shell.Application

  # ファイルリストを取得
  if ($recurseEnabled) {
    $targetFiles = Get-ChildItem -File -Recurse | ForEach-Object { $_.Fullname }
  } else {
    $targetFiles = Get-ChildItem -File | ForEach-Object { $_.Fullname }
  }

  # ファイル毎の処理
  foreach($targetFile in $targetFiles) {
    $dateStr = ""
    $dateSource = ""
    $dateSourceColor = "" 

    # フォルダパス/ファイル名/拡張子を取得
    $folderPath = Split-Path $targetFile
    $fileName = Split-Path $targetFile -Leaf
    $fileExt = (Get-Item $targetFile).Extension.substring(1).ToLower()

    # 日付文字列を取得(YYYY/MM/DD HH:MM:SS)
    if ($fileExt -eq "jpg") {
      # Exifより取得
      $dateStr = getExifDate $targetFile
      $dateSource = "EXIF"
      $dateSourceColor = "Green"
    } elseif (($fileExt -eq "mov") `
              -or ($fileExt -eq "mp4") `
              -or ($fileExt -eq "heic")) {
      # 詳細プロパティより取得
      $date = getPropDate $folderPath $fileName
      $dateStr = $date.ToString("yyyy/MM/dd HH:mm:ss")
      $dateSource = "DETL"
      $dateSourceColor = "Cyan"
    }
    if (!$dateStr -and `
        (($fileExt -eq "jpg") `
         -or ($fileExt -eq "mov") `
         -or ($fileExt -eq "mp4") `
         -or ($fileExt -eq "heic") `
         -or ($fileExt -eq "png"))) {
      # 失敗したらファイル名より取得
      $dateStr = getFnameDate $fileName
      $dateSource = "NAME"
      $dateSourceColor = "Yellow"
    }
    if (!$dateStr) {
      # それでも失敗したらスキップ
      printSkipped $folderPath $fileName
      continue
    }

    # ファイル名を変更(YYYYMMDD-HHMMSS-NNN.EXT)
    $renamed = $false
    $newFileName = ""
    $tempFileBase = "GoPro_" + $dateStr.replace("/", "-").replace(" ", "_").replace(":", "")
    if ($dryRunEnabled) {
        $newPath = $folderPath + "\" + $tempFileBase + "-NNN" + "." + $fileExt
        $newFileName = Split-Path $newPath -Leaf
        $renamed = $true
    } else {
      for ([int]$i = 0; $i -le 999; $i++)
      {
        $newPath = $folderPath + "\" + $tempFileBase + "-" + $i.ToString("000") + "." + $fileExt
        $newFileName = Split-Path $newPath -Leaf
        # 変更不要なら抜ける
        if ($fileName -eq $newFileName) {
          $renamed = $true
          break
        }
        # ファイル重複チェック
        if ((Test-Path $newPath) -eq $false)
        {
          try {
            updateCreationTime $targetFile $date
            $updatedFileName = $targetFile + ".mp4"
            Rename-Item $updatedFileName -newName $newFileName
          } catch {
            break
          }
          $renamed = $true
          break
        }
      }
    }
    if (!$renamed) {
      printSkipped $folderPath $fileName
      continue
    }

    # 作成/更新日時を変更
    if (!$dryRunEnabled) {
      Set-ItemProperty $newPath -Name CreationTime -Value $dateStr
      Set-ItemProperty $newPath -Name LastWriteTime -Value $dateStr
    }

    # 結果表示
    $rfPath = (Resolve-Path $folderPath -Relative)
    if ($rfPath.StartsWith("..\")) {
      $rfPath = ".\"
    }
    Write-Host "[$rfPath] $fileName -> $newFileName ($dateStr " -NoNewline
    Write-Host "$dateSource" -ForegroundColor $dateSourceColor -NoNewline
    Write-Host ")"
  }

  # シェルオブジェクトを解放
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellObject) | out-null
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

# 実行
main
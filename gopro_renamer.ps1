#
# GoPro Renamer
#

Param(
    [switch]$d,
    [switch]$r,
    [string]$folder
)

$PSDefaultParameterValues['*:Encoding'] = 'utf8'


$dryRunEnabled = $d
$recurseEnabled = $r

# include config
. ".\config.ps1"

$appVersion = "v1.0.6"
Add-Type -AssemblyName System.Drawing

if ((Test-Path $FFMPEG) -ne $true)
{
    Write-Host "ffmpeg not found"
    Exit
}


if (!$folder)
{
    Write-Host "undefined target folder"
    Exit
}

if ((Test-Path $folder) -eq $false)
{
    Write-Host "undefined target folder"
    Exit
}

$tpath = $folder

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

    return $newFileName
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
        -or ($propertyName.Trim() -eq "メディアの作成日時")) {
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
    Write-Host "Error: fail getPropDate:" + $file;
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

# GoProのデータかどうか判定
function isGoProFile($targetFile)
{
    $file = Split-Path $targetFile -Leaf

    if ($file.Length -ne 12)
    {
        return $false
    }

    if (@("GH", "GX").Contains($file.Substring(0, 2)) -ne $true)
    {
        return $false
    }

    return $true
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
    $targetFiles = Get-ChildItem -Path "$tpath" -File -Recurse | ForEach-Object { $_.Fullname }
  } else {
    $targetFiles = Get-ChildItem -Path "$tpath" -File | ForEach-Object { $_.Fullname }
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
    $fileBaseName = $fileName -replace ".mp4$", ""
    
    if ((isGoProFile $targetFile) -eq $false)
    {
        printSkipped $folderPath $fileName
        continue
    }

    if ($fileName.Substring(2, 2) -eq "00")
    {
        printSkipped $folderPath $fileName
        continue
    }

    # 詳細プロパティより取得
    $date = getPropDate $folderPath $fileName
    $dateStr = $date.ToString("yyyy/MM/dd HH:mm:ss")
    $dateSource = "DETL"
    $dateSourceColor = "Cyan"

    # ファイル名を変更(YYYYMMDD-HHMMSS-NNN.EXT)
    $newFileName = ""
    $tempFileBase = $dateStr.replace("/", "-").replace(" ", "_").replace(":", "")
    if ($dryRunEnabled) {
        $newPath = $folderPath + "\" + $tempFileBase + "-${fileBaseName}" + "." + $fileExt
        $newFileName = Split-Path $newPath -Leaf
    } else {
        $newPath = $folderPath + "\" + $tempFileBase + "-" + $fileBaseName + "." + $fileExt
        $newFileName = Split-Path $newPath -Leaf

        # 生成後のファイルがある場合はSkip
        if ((Test-Path $newPath))
        {
            printSkipped $folderPath $fileName
            continue
        }

        try {
            $updatedFileName = updateCreationTime $targetFile $date
            Rename-Item $updatedFileName -newName $newFileName
        } catch {
            break
        }
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
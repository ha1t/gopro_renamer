﻿#
# Media Dater
# メタデータ内の撮影日時情報をもとにファイルの名前/作成日時/更新日時を変更する
#
# 動作環境:
#   Windows PowerShell 5.1
#   ・このスクリプトの実行が許可されていること
#   ・このスクリプトファイルがBOM付UTF-8であること
#
# 対応ファイル:
#   ・拡張子がjpg/png/heicの画像ファイル
#   ・拡張子がmov/mp4の動画ファイル
#
# 動作内容:
#   1. 日付情報を取得する
#     ・jpgファイルはExifを参照
#     ・heic/mov/mp4ファイルは詳細プロパティの「撮影日時」「メディアの作成日時」を参照
#       (両方設定されている場合は「撮影日時」を優先する)
#     ・上記処理で日時情報を取得できなかったファイルやpngファイルはファイル名を参照
#       (YYYYMMDD-HHMMSS\*.拡張子 の形式であれば日付と見なす)
#     ・それでも取得できなかった場合はそのファイルはスキップ
#   2. ファイル名を変更する
#     ・YYYYMMDD-HHMMSS-3桁連番.元の拡張子 となる
#   3. ファイルの作成日時/更新日時を変更する
#     ・撮影日時と同じとなる
#
# 使い方:
#   1. media_dater.ps1 を適当な場所へ配置
#   2. 画像/動画ファイルが存在するフォルダ上でPowerShellを開く
#   3. media_dater.ps1 を実行する
#
# 参考:
#   https://neos21.hatenablog.com/entry/2019/11/11/080000
#   https://qiita.com/Kosen-amai/items/52ec7e4e2f15f6a09bc3
#   https://qiita.com/kmr_hryk/items/882b4851e23cec607e70
#
Add-Type -AssemblyName System.Drawing
$appVersion = "v1.0.1"
$dryRun = $false

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

  return $ret
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
function printSkipped($fname) {
  Write-Host "$fname (skipped)"
}

# メイン処理
function main {
  # バナーを表示
  $mode = if ($dryRun) { " (dry run)" } else { "" }
  Write-Host "== Media Dater $appVersion$mode =="

  # シェルオブジェクトを生成
  $shellObject = New-Object -ComObject Shell.Application

  # ファイルリストを取得
  $targetFiles = Get-ChildItem -File | ForEach-Object { $_.Fullname }

  # ファイル毎の処理
  foreach($targetFile in $targetFiles) {
    $dateStr = ""
    $dateSource = ""

    # フォルダパス/ファイル名/拡張子を取得
    $folderPath = Split-Path $targetFile
    $fileName = Split-Path $targetFile -Leaf
    $fileExt = (Get-Item $targetFile).Extension.substring(1).ToLower()

    # 日付文字列を取得(YYYY/MM/DD HH:MM:SS)
    if ($fileExt.endsWith("jpg")) {
      # Exifより取得
      $dateStr = getExifDate $targetFile
      $dateSource = "EXIF"
    } elseif (($fileExt -eq "mov") `
              -or ($fileExt -eq "mp4") `
              -or ($fileExt -eq "heic")) {
      # 詳細プロパティより取得
      $dateStr = getPropDate $folderPath $fileName
      $dateSource = "DETL"
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
    }
    if (!$dateStr) {
      # それでも失敗したらスキップ
      printSkipped $fileName
      continue
    }

    # ファイル名を変更(YYYYMMDD-HHMMSS-NNN.EXT)
    $renamed = $false
    $newFileName = ""
    $tempFileBase = $dateStr.replace("/", "").replace(" ", "-").replace(":", "")
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
        if (!$dryRun) {
          try {
            Rename-Item $targetFile -newName $newFileName
          } catch {
            break
          }
        }
        $renamed = $true
        break
      }
    }
    if (!$renamed) {
      printSkipped $fileName
      continue
    }

    # 作成/更新日時を変更
    if (!$dryRun) {
      Set-ItemProperty $newFileName -Name CreationTime -Value $dateStr
      Set-ItemProperty $newFileName -Name LastWriteTime -Value $dateStr
    }

    Write-Host "$fileName -> $newFileName ($dateStr $dateSource)"
  }

  # シェルオブジェクトを解放
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellObject) | out-null
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

# 実行
main

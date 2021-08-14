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

# ffmpeg �� creation_time ���X�V����
# creation_time=2021-08-14T23:51:59.000000Z
function updateCreationTime($fileName, [System.DateTime]$updateDate)
{
    $updateDate -= New-TimeSpan -Hours 9 # ���łɈ����Z���Ă���͂��Ȃ̂Ɂc
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

# Exif�������������𐶐�����
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

# �ڍ׃v���p�e�B�������������𐶐�����
function getPropDate($folder, $file) {
  $shellFolder = $shellObject.namespace($folder)
  $shellFile = $shellFolder.parseName($file)
  $selectedPropertyNo = ""
  $selectedPropertyName = ""
  $selectedPropertyValue = ""

  for ($i = 0; $i -lt 300; $i++) { # 208�܂ŒT���Ώ\��?
    $propertyName = $shellFolder.getDetailsOf($Null, $i)
    if (($propertyName -eq "�B�e����") `
        -or ($propertyName -eq "���f�B�A�̍쐬����")) {
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
  $time = "0" + $ret.substring(16) + ":00" # �b�͎擾�ł��Ȃ��̂�00��ݒ�
  $time = $time.substring($time.length - 8, 8)
  $ret = $ret.substring(1, 5) + $ret.substring(7, 3) + $ret.substring(11, 2) + " " + $time
  $date = Get-Date $ret
  $ts = New-TimeSpan -Hours 9
  return ($date - $ts)
}

# �t�@�C���X�L�b�v���̕\��
function printSkipped($folder, $file) {
  $rfPath = (Resolve-Path $folder -Relative)
  if ($rfPath.StartsWith("..\")) {
    $rfPath = ".\"
  }
  Write-Host "[$rfPath] $file (" -NoNewline
  Write-Host "skipped" -ForegroundColor Red -NoNewline
  Write-Host ")"
}

# ���C������
function main {
  # �o�i�[��\��
  $mode = if ($dryRunEnabled) { " (dry run)" } else { "" }
  Write-Host "== Media Dater $appVersion$mode =="

  # �V�F���I�u�W�F�N�g�𐶐�
  $shellObject = New-Object -ComObject Shell.Application

  # �t�@�C�����X�g���擾
  if ($recurseEnabled) {
    $targetFiles = Get-ChildItem -File -Recurse | ForEach-Object { $_.Fullname }
  } else {
    $targetFiles = Get-ChildItem -File | ForEach-Object { $_.Fullname }
  }

  # �t�@�C�����̏���
  foreach($targetFile in $targetFiles) {
    $dateStr = ""
    $dateSource = ""
    $dateSourceColor = "" 

    # �t�H���_�p�X/�t�@�C����/�g���q���擾
    $folderPath = Split-Path $targetFile
    $fileName = Split-Path $targetFile -Leaf
    $fileExt = (Get-Item $targetFile).Extension.substring(1).ToLower()

    # ���t��������擾(YYYY/MM/DD HH:MM:SS)
    if ($fileExt -eq "jpg") {
      # Exif���擾
      $dateStr = getExifDate $targetFile
      $dateSource = "EXIF"
      $dateSourceColor = "Green"
    } elseif (($fileExt -eq "mov") `
              -or ($fileExt -eq "mp4") `
              -or ($fileExt -eq "heic")) {
      # �ڍ׃v���p�e�B���擾
      $date = getPropDate $folderPath $fileName
      $dateStr = $date.ToString("yyyy/MM/dd HH:mm:ss")
      $dateSource = "DETL"
      $dateSourceColor = "Cyan"
    }
    if (!$dateStr) {
      # ����ł����s������X�L�b�v
      printSkipped $folderPath $fileName
      continue
    }

    # �t�@�C������ύX(YYYYMMDD-HHMMSS-NNN.EXT)
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
        # �ύX�s�v�Ȃ甲����
        if ($fileName -eq $newFileName) {
          $renamed = $true
          break
        }
        # �t�@�C���d���`�F�b�N
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

    # �쐬/�X�V������ύX
    if (!$dryRunEnabled) {
      Set-ItemProperty $newPath -Name CreationTime -Value $dateStr
      Set-ItemProperty $newPath -Name LastWriteTime -Value $dateStr
    }

    # ���ʕ\��
    $rfPath = (Resolve-Path $folderPath -Relative)
    if ($rfPath.StartsWith("..\")) {
      $rfPath = ".\"
    }
    Write-Host "[$rfPath] $fileName -> $newFileName ($dateStr " -NoNewline
    Write-Host "$dateSource" -ForegroundColor $dateSourceColor -NoNewline
    Write-Host ")"
  }

  # �V�F���I�u�W�F�N�g�����
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($shellObject) | out-null
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

# ���s
main
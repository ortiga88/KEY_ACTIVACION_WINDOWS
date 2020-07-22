#------------------------------------------[Initialisations]---------------------------------------

$scriptName = $myInvocation.MyCommand.Name

#-------------------------------------------[Declarations]-----------------------------------------

$key = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
$digitalProductId = (Get-ItemProperty -Path $key).DigitalProductId
$isWin8 = $digitalProductId[66] / 6 -band 1
$digitalProductId[66] = $digitalProductId[66] -band 247 -or ($isWin8 -band 2) * 4
$productName = (Get-ItemProperty -Path $key).ProductName
$productId = (Get-ItemProperty -Path $key).ProductId
$maps = "BCDFGHJKMPQRTVWXY2346789"

$keyOffset = 52
$i = 24

#--------------------------------------------[Functions]-------------------------------------------

Function RR-getProductKey {
  Try {
    Do {
      $j = 14
      $current = 0
      Do {
        $current = $current * 256
        $current = $digitalProductId[$j + $keyOffset] + $current
        $round = $digitalProductId[$j + $keyOffset] = ($current / 24)
        $round = [math]::floor($round)
        $digitalProductId[$j + $keyOffset] = $round
        $current = $current % 24
        $j--
      } While ($j -ge 0)
      $i--
      $keyOutput = $maps.Substring($current++, 1) + $keyOutput
      $last = $current
      $last--
    } While ($i -ge 0)

    if($isWin8 = 1) {
      $keypart1 = $keyOutput.Substring(1, $last)
      $insert = "N"
      $replace = $keypart1 + $insert
      $keyOutput = $keyOutput.Replace("$keypart1", "$replace")
      if($last = 0) {
        $keyOutput = $insert + $keyOutput
      }
    }

    $productKey = $keyOutput.Substring(1, 5) + "-" + $keyOutput.Substring(6, 5) + "-" + $keyOutput.Substring(11, 5) + "-" + $keyOutput.Substring(16, 5) + "-" + $keyOutput.Substring(21, 5)

    Write-Host " Product Name:" $productName
    Write-Host " Product Id:" $productId
    Write-Host " Product Key:" $productKey
  }
  catch {
    Write-Host $_.Exception.Message
  }
}

#--------------------------------------------[Execution]-------------------------------------------

$startDate = Get-Date -format MM-dd-yyyy_HH:mm:ss
Write-Host "Starting $scriptName at $startDate" -ForegroundColor Green

RR-getProductKey

$stopDate = Get-Date -format MM-dd-yyyy_HH:mm:ss
Write-Host "Stopped $scriptName at $startDate" -ForegroundColor Green


pause

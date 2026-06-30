param(
  [string]$BrowserPath,
  [string]$UserDataDir,
  [int]$Port = 9223,
  [string]$Url = "https://seller.ozon.ru/app/dashboard/main",
  [switch]$RestartExisting
)

function Resolve-BrowserPath {
  param([string]$ExplicitPath)

  if ($ExplicitPath) {
    return $ExplicitPath
  }

  $candidates = @(
    "C:\Program Files\Yandex\YandexBrowser\Application\browser.exe",
    "C:\Program Files\Google\Chrome\Application\chrome.exe",
    "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
  )

  foreach ($candidate in $candidates) {
    if (Test-Path -LiteralPath $candidate) {
      return $candidate
    }
  }

  throw "Browser executable not found. Pass -BrowserPath explicitly."
}

function Resolve-UserDataDir {
  param([string]$ResolvedBrowserPath, [string]$ExplicitDir)

  if ($ExplicitDir) {
    return $ExplicitDir
  }

  switch -Wildcard ($ResolvedBrowserPath) {
    "*YandexBrowser*" { return "$env:LOCALAPPDATA\Yandex\YandexBrowser\User Data" }
    "*Chrome*" { return "$env:LOCALAPPDATA\Google\Chrome\User Data" }
    "*Edge*" { return "$env:LOCALAPPDATA\Microsoft\Edge\User Data" }
    default { throw "Cannot infer user data dir. Pass -UserDataDir explicitly." }
  }
}

$resolvedBrowserPath = Resolve-BrowserPath -ExplicitPath $BrowserPath
$resolvedUserDataDir = Resolve-UserDataDir -ResolvedBrowserPath $resolvedBrowserPath -ExplicitDir $UserDataDir
$processName = [System.IO.Path]::GetFileNameWithoutExtension($resolvedBrowserPath)

if ($RestartExisting) {
  Get-Process -Name $processName -ErrorAction SilentlyContinue | Stop-Process -Force
  Start-Sleep -Seconds 2
}

$arguments = @(
  "--remote-debugging-port=$Port",
  "--user-data-dir=$resolvedUserDataDir",
  "--new-window",
  $Url
)

Start-Process -FilePath $resolvedBrowserPath -ArgumentList $arguments -WindowStyle Hidden
Start-Sleep -Seconds 6

Invoke-WebRequest -UseBasicParsing "http://127.0.0.1:$Port/json/version" | Select-Object -ExpandProperty Content

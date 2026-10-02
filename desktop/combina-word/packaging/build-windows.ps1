$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
Set-Location $Root

if (-not (Test-Path .venv-build)) {
    py -3 -m venv .venv-build
}
& .\.venv-build\Scripts\python.exe -m pip install --upgrade pip
& .\.venv-build\Scripts\python.exe -m pip install -r requirements.txt
& .\.venv-build\Scripts\pyinstaller.exe --clean --noconfirm combina-word.spec
Write-Host "Ejecutable creado en $Root\dist\CombinaWord.exe"

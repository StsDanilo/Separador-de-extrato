$ErrorActionPreference = "Stop"

function Get-BasePython {
  if ($env:PYTHON -and (Test-Path $env:PYTHON)) {
    return @($env:PYTHON)
  }

  $python = Get-Command python -ErrorAction SilentlyContinue
  if ($python) {
    return @($python.Source)
  }

  $py = Get-Command py -ErrorAction SilentlyContinue
  if ($py) {
    return @($py.Source, "-3")
  }

  throw "Python nao encontrado. Instale o Python 3 ou defina a variavel de ambiente PYTHON com o caminho do python.exe."
}

function Invoke-BasePython {
  param([string[]] $Arguments)
  $command = $script:BasePython[0]
  $prefixArgs = @()
  if ($script:BasePython.Count -gt 1) {
    $prefixArgs = $script:BasePython[1..($script:BasePython.Count - 1)]
  }
  & $command @prefixArgs @Arguments
}

if (-not (Test-Path ".\.venv")) {
  $script:BasePython = Get-BasePython
  Invoke-BasePython -Arguments @("-m", "venv", ".venv")
}

.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe run_app.py

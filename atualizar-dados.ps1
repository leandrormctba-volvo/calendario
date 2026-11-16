# Atualiza o arquivo data.js do dashboard a partir da planilha Kanban (aba "Responsabilidades").
# Uso:  .\atualizar-dados.ps1                       (usa a planilha Kanban mais recente da pasta Downloads)
#       .\atualizar-dados.ps1 -Arquivo "C:\caminho\planilha.xlsb"
param(
    [string]$Arquivo,
    [string]$Aba = 'Responsabilidades'
)

$ErrorActionPreference = 'Stop'

if (-not $Arquivo) {
    $achado = Get-ChildItem "$env:USERPROFILE\Downloads" -Filter 'Kanban calend*.xls*' |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $achado) { throw 'Nenhuma planilha "Kanban calendario" encontrada em Downloads. Use -Arquivo.' }
    $Arquivo = $achado.FullName
}
if (-not (Test-Path -LiteralPath $Arquivo)) { throw "Planilha nao encontrada: $Arquivo" }
Write-Host "Lendo: $Arquivo"

$C = [char]0xC7; $A = [char]0xC3
$meses = @('JANEIRO','FEVEREIRO',"MAR${C}O",'ABRIL','MAIO','JUNHO','JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO')
$naoRealizado = "N${A}O REALIZADO"

function Texto($v) { if ($null -eq $v) { '' } else { ([string]$v).Trim() } }

function Participantes($valor, $obs) {
    if ($valor -is [double]) { return [int][math]::Round($valor) }
    $t = Texto $valor
    if ($t -match '^\s*(\d+)') { return [int]$Matches[1] }
    if ($obs -match '(\d+)\s*participantes') { return [int]$Matches[1] }
    return $null
}

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
try {
    $wb = $xl.Workbooks.Open($Arquivo, 0, $true)
    $ws = $wb.Worksheets.Item($Aba)
    $ultima = $ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1
    # Colunas: B=Tarefa C=Projeto D=Responsavel E=Participantes F=Data I=Status J=% K=Observacao
    $vals = $ws.Range("B7:K$ultima").Value2
    $datas = $ws.Range("F7:F$ultima")

    $mesAtual = $null
    $tarefas = New-Object System.Collections.ArrayList
    for ($r = 1; $r -le $vals.GetLength(0); $r++) {
        $nome = Texto $vals[$r, 1]
        $cat  = Texto $vals[$r, 2]
        if (-not $nome) { continue }
        $idx = [array]::IndexOf($meses, $nome.ToUpper())
        if ($idx -ge 0 -and -not $cat) { $mesAtual = $idx + 1; continue }
        if (-not $mesAtual) { continue }

        $obs    = Texto $vals[$r, 10]
        $status = (Texto $vals[$r, 8]).ToUpper()
        $pctRaw = $vals[$r, 9]
        $pct    = if ($pctRaw -is [double]) { $pctRaw } else { $null }
        $part   = Participantes $vals[$r, 4] $obs
        $data   = Texto $datas.Cells.Item($r, 1).Text

        $inferido = $false
        if ($status -eq 'FEITO') { $sit = 'Realizada' }
        elseif ($status -eq 'CANCELADO' -or $status -eq $naoRealizado) { $sit = 'Cancelada' }
        elseif ($status -eq 'TRANSFERIDO') { $sit = 'Transferida' }
        elseif ($status -eq 'FAZENDO') { $sit = 'Em andamento' }
        elseif (($null -ne $pct -and $pct -ge 1) -or ($part -gt 0)) { $sit = 'Realizada'; $inferido = $true }
        else { $sit = 'Pendente' }

        [void]$tarefas.Add([ordered]@{
            mes           = $mesAtual
            evento        = ($nome -replace '\s+', ' ')
            categoria     = $cat
            responsavel   = (Texto $vals[$r, 3])
            participantes = $part
            data          = $data
            status        = $status
            situacao      = $sit
            inferido      = $inferido
        })
    }
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
}

$saida = [ordered]@{
    atualizadoEm = (Get-Date -Format 'dd/MM/yyyy HH:mm')
    ano          = 2026
    tarefas      = $tarefas
}
$json = $saida | ConvertTo-Json -Depth 5
$destino = Join-Path $PSScriptRoot 'data.js'
[IO.File]::WriteAllText($destino, "window.DADOS = $json;`n", (New-Object Text.UTF8Encoding $false))
Write-Host "OK: $($tarefas.Count) tarefas exportadas para $destino"

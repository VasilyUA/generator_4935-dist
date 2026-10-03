. (Join-Path $PSScriptRoot 'common.ps1')

$Projects = @(
    @{ Key = '1'; Dir = 'report_generator';                 Label = 'Рапорти (здав/прийняв посаду, відпустка тощо) та масове переміщення' }
    @{ Key = '2'; Dir = 'report_generator_old_new_state';   Label = 'Масова зміна штату (старий список -> новий список)' }
    @{ Key = '3'; Dir = 'generator_timesheet';               Label = 'Табель обліку особового складу' }
    @{ Key = '4'; Dir = 'generator_br_and_report_for_money'; Label = 'БР + рапорти на грошову винагороду' }
    @{ Key = '5'; Dir = 'final_combat_report';               Label = 'Щоденне бойове донесення (ПБД)' }
)

function Ensure-ProjectResources {
    param($Project)
    $projectDir = Join-Path $RepoRoot $Project.Dir
    $resourcesDir = Join-Path $projectDir 'resources'
    $sampleScript = Join-Path $projectDir 'resources_example\generate_sample.py'

    if (-not (Test-Path $resourcesDir) -and (Test-Path $sampleScript)) {
        Write-Info "У проєкту немає теки resources - створюю зразок (вигадані дані)..."
        Invoke-StepWithGuidance -PathForGuidance $resourcesDir -Action {
            Push-Location $projectDir
            try {
                & $VenvPython 'resources_example\generate_sample.py'
            } finally {
                Pop-Location
            }
        }
        Write-Info "Зразок створено з ВИГАДАНИМИ даними. Замініть їх реальними у resources\ (див. README.md)."
        Start-Sleep -Seconds 2
    }
}

function Show-Menu {
    Write-Host ""
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host " Генератори звітів - оберіть програму" -ForegroundColor Cyan
    Write-Host "================================================" -ForegroundColor Cyan
    foreach ($p in $Projects) {
        Write-Host "  $($p.Key)) $($p.Label)"
    }
    Write-Host "  0) Вихід"
    Write-Host ""
    return (Read-Host "Ваш вибір").Trim()
}

function Invoke-Project {
    param($Project)
    Ensure-ProjectResources -Project $Project

    $projectDir = Join-Path $RepoRoot $Project.Dir
    Write-Section "Запуск: $($Project.Label)"
    Push-Location $projectDir
    try {
        & $VenvPython 'index.py'
        $exitCode = $LASTEXITCODE
    } finally {
        Pop-Location
    }

    Write-Host ""
    if ($exitCode -ne 0) {
        Write-Err "Програма завершилась з помилкою (код $exitCode). Перегляньте повідомлення вище."
    } else {
        Write-Ok "Готово."
    }
    Write-Host "Натисніть будь-яку клавішу, щоб повернутись у меню..."
    Wait-KeyPress
}

function Main {
    Write-Section "Генератори звітів - запуск"

    Update-IfPossible
    Ensure-PythonEnvironment

    while ($true) {
        $choice = Show-Menu
        if ($choice -eq '0') { break }
        $selected = $Projects | Where-Object { $_.Key -eq $choice }
        if (-not $selected) {
            Write-Err "Невірний вибір, спробуйте ще раз."
            continue
        }
        Invoke-Project -Project $selected
    }

    Write-Host ""
    Write-Host "До побачення!"
}

try {
    Main
} catch {
    Write-Host ""
    Write-Err "СТАЛАСЯ ПОМИЛКА:"
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Info "Якщо звертаєтесь за допомогою - надішліть, будь ласка, увесь текст вище й ці дані:"
    Write-Info "  Користувач Windows: $env:USERNAME | Адміністратор: $(Test-IsAdmin) | PowerShell: $($PSVersionTable.PSVersion) | Папка проєкту: $RepoRoot"
} finally {
    Write-Host ""
    Write-Host "Натисніть будь-яку клавішу, щоб закрити вікно..."
    Wait-KeyPress
}

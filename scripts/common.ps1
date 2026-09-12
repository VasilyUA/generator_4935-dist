$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding(1251)

# Без цього на старіших системах (PowerShell 5.1 зі старими налаштуваннями
# .NET) Invoke-WebRequest падає з незрозумілою помилкою SSL/TLS ще ДО того,
# як узагалі дійде до мережі - користувач бачить це як "немає інтернету",
# хоча інтернет насправді є.
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch {
}

# $PSScriptRoot тут - завжди тека scripts\ (де лежить сам common.ps1),
# незалежно від того, звідки саме його підключили через
# ". (Join-Path $PSScriptRoot 'common.ps1')" - PowerShell виставляє
# $PSScriptRoot окремо для КОЖНОГО файлу за його ВЛАСНИМ розташуванням, а не
# успадковує від того, хто підключив.
$RepoRoot = Split-Path -Parent $PSScriptRoot

$VenvDir = Join-Path $RepoRoot 'venv'
$VenvPython = Join-Path $VenvDir 'Scripts\python.exe'
$RequirementsTxt = Join-Path $RepoRoot 'requirements.txt'
$MarkerFile = Join-Path $VenvDir '.deps_installed.marker'
$PythonVersion = '3.12.10'
$PythonInstallerUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-amd64.exe"
$KnownUserInstallPython = Join-Path $env:LocalAppData 'Programs\Python\Python312\python.exe'

# Портативний Git (без встановлювача, без прав адміністратора) - на "чистій"
# машині (нічого не встановлено) git відсутній так само часто, як і Python, і
# без нього автооновлення взагалі не має чим працювати. ВАЖЛИВО: тека
# призначення має лишатись КОРОТКОЮ (напр. саме ця, під LocalAppData) -
# перевірено на практиці: розпаковування в задовгий шлях (глибоко вкладена
# тека) провалюється мовчки через ліміт Windows MAX_PATH, оскільки всередині
# PortableGit і так багато вкладених тек (mingw64\... тощо).
$GitPortableUrl = 'https://github.com/git-for-windows/git/releases/download/v2.55.0.windows.5/PortableGit-2.55.0.5-64-bit.7z.exe'
$KnownPortableGitDir = Join-Path $env:LocalAppData 'Programs\PortableGit'
$KnownPortableGitExe = Join-Path $KnownPortableGitDir 'cmd\git.exe'

# --- Автооновлення: публічне дзеркало приватного репозиторію ---------------
# Приватний репозиторій (звичайна розробка, реальний git push розробника) НЕ
# годиться як джерело оновлень для роздистрибутованих копій кінцевих
# користувачів - анонімний git fetch до ПРИВАТНОГО репозиторію завжди
# провалюється ("could not read Username", підтверджено на практиці). Тому
# окремий ПУБЛІЧНИЙ репозиторій - точна копія лише ПОТОЧНОГО стану файлів,
# БЕЗ історії git приватного репозиторію (щоб не розкрити нічого з 500+
# старих комітів приватного репо) - синхронізується окремим GitHub Actions
# workflow у приватному репозиторії при кожному push у main.
$PublicMirrorUrl = 'https://github.com/VasilyUA/generator_4935-dist.git'
$UpdateRemoteName = 'update-source'
$UpdateBranch = 'main'
$PrivateRepoOwnerSlashName = 'VasilyUA/generator_4935'

function Write-Section {
    param([string]$Text)
    Write-Host ""
    Write-Host $Text -ForegroundColor Cyan
}

function Write-Info {
    param([string]$Text)
    Write-Host $Text -ForegroundColor Gray
}

function Write-Ok {
    param([string]$Text)
    Write-Host $Text -ForegroundColor Green
}

function Write-Err {
    param([string]$Text)
    Write-Host $Text -ForegroundColor Red
}

function Write-Warn {
    param([string]$Text)
    Write-Host $Text -ForegroundColor Yellow
}

function Wait-KeyPress {
    # ReadKey кидає виняток, якщо консоль перенаправлена (наприклад, запуск
    # не напряму подвійним кліком) - у такому разі просто не чекаємо,
    # аби скрипт не впав саме на кроці "натисніть клавішу".
    try {
        [void][System.Console]::ReadKey($true)
    } catch {
    }
}

# ---------------------------------------------------------------------------
# Діагностика прав доступу та зрозумілі підказки для користувача, який не
# розбирається в програмуванні - "Access is denied" саме по собі нічого не
# пояснює, тому кожна ризикована операція (запис на диск, встановлення
# програм) супроводжується конкретними кроками, що спробувати.
# ---------------------------------------------------------------------------

function Test-IsAdmin {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function Test-IsAccessDeniedMessage {
    param([string]$Message)
    if (-not $Message) { return $false }
    # Реальний .NET-виняток пишеться як "Access to the path 'X' is denied." -
    # тобто між "access" і "denied" стоять ще слова, тому регекс перевіряє їх
    # ОКРЕМО (обидва слова десь у тексті), а не як одну сусідню фразу -
    # інакше саме найпоширеніший реальний формат помилки не розпізнавався б.
    return (
        $Message -match '(?i)unauthorizedaccess' -or
        $Message -match '(?i)winerror\s*5\b' -or
        $Message -match 'відмовлено в доступі' -or
        $Message -match 'відмовлено у доступі' -or
        (($Message -match '(?i)access') -and ($Message -match '(?i)denied')) -or
        (($Message -match '(?i)permission') -and ($Message -match '(?i)denied'))
    )
}

function Get-AccessDeniedGuidance {
    param([string]$Path, [string]$LauncherFileName = 'START.bat')
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("")
    $lines.Add("⚠ Схоже на помилку доступу (Access is denied).")
    if ($Path) {
        $lines.Add("Місце, де сталась помилка: $Path")
    }
    $lines.Add("")
    $lines.Add("Спробуйте по черзі (після кожного кроку запускайте $LauncherFileName ще раз):")
    $lines.Add("  1. Тимчасово вимкніть антивірус (чи додайте папку 'generator_4935' у винятки антивіруса) - антивірус часто на секунду блокує щойно завантажений файл, поки перевіряє його.")
    $lines.Add("  2. Перевірте, що папка проєкту НЕ лежить у захищеній системній теці (напр. 'C:\Program Files', 'C:\Windows', корінь диска 'C:\'). Перемістіть усю папку 'generator_4935' на Робочий стіл або в 'Документи'.")
    $lines.Add("  3. Закрийте Word/Excel і будь-які вікна провідника, відкриті саме в цій папці.")
    $lines.Add("  4. Переконайтесь, що OneDrive/хмарне сховище не 'заморозило' файл (значок хмаринки біля файлу) - зачекайте, поки файл повністю синхронізується.")
    $lines.Add("  5. Якщо нічого з вище не допомогло - натисніть правою кнопкою миші на $LauncherFileName і оберіть 'Запуск від імені адміністратора'.")
    if (-not (Test-IsAdmin)) {
        $lines.Add("")
        $lines.Add("(Зараз скрипт запущено БЕЗ прав адміністратора - зазвичай це нормально й достатньо, права адміністратора потрібні лише як останній варіант.)")
    }
    return ($lines -join "`n")
}

function Test-PathWritable {
    param([string]$Path)
    try {
        if (-not (Test-Path $Path)) {
            New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
        }
        $probe = Join-Path $Path (".write_test_{0}.tmp" -f ([guid]::NewGuid().ToString('N')))
        [IO.File]::WriteAllText($probe, "test")
        Remove-Item -Path $probe -Force -ErrorAction SilentlyContinue
        return $true
    } catch {
        return $false
    }
}

# Обгортка для кроку, що може впасти через права доступу - додає зрозумілу
# підказку ПРЯМО до повідомлення про помилку, а не лишає користувача сам на
# сам із сирим текстом винятку PowerShell/Python.
function Invoke-StepWithGuidance {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Action,
        [string]$PathForGuidance = $null
    )
    try {
        & $Action
    } catch {
        $message = $_.Exception.Message
        if (Test-IsAccessDeniedMessage $message) {
            throw "$message`n$(Get-AccessDeniedGuidance -Path $PathForGuidance)"
        }
        throw
    }
}

# Повторює операцію кілька разів із паузою - антивірус, що на мить блокує
# щойно записаний файл, чи короткий збій мережі - типові ТИМЧАСОВІ причини,
# коли повторна спроба сама по собі вирішує проблему без жодних дій людини.
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Action,
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 5,
        [string]$Description = "операцію",
        [string]$PathForGuidance = $null
    )
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            & $Action
            return
        } catch {
            $isLast = ($attempt -eq $MaxAttempts)
            if (-not $isLast) {
                Write-Warn "Спроба $attempt з $MaxAttempts ($Description) не вдалась: $($_.Exception.Message)"
                Write-Info "Пробую ще раз через $DelaySeconds сек..."
                Start-Sleep -Seconds $DelaySeconds
            } else {
                $message = $_.Exception.Message
                if (Test-IsAccessDeniedMessage $message) {
                    throw "$message`n$(Get-AccessDeniedGuidance -Path $PathForGuidance)"
                }
                throw
            }
        }
    }
}

function Test-PythonUsable {
    param([string]$PythonExe)
    if (-not $PythonExe) { return $false }
    if (-not (Test-Path $PythonExe -PathType Leaf)) { return $false }
    try {
        $out = & $PythonExe --version 2>&1
        $text = ($out -join ' ')
        return ($LASTEXITCODE -eq 0 -and $text -match '^Python 3')
    } catch {
        return $false
    }
}

function Find-Python {
    # На чистій системі python.exe/python3.exe в PATH за замовчуванням - це
    # заглушка (App Execution Alias), яка відкриває Microsoft Store замість
    # запуску, тому наявність у PATH недостатня - треба реально перевірити
    # вивід `--version`.
    if (Test-PythonUsable $KnownUserInstallPython) {
        return $KnownUserInstallPython
    }

    foreach ($name in @('python', 'python3')) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($cmd -and (Test-PythonUsable $cmd.Source)) {
            return $cmd.Source
        }
    }

    $pyCmd = Get-Command py -ErrorAction SilentlyContinue
    if ($pyCmd) {
        try {
            $verOut = & py -3.12 --version 2>&1
            if ($LASTEXITCODE -eq 0 -and (($verOut -join ' ') -match '^Python 3')) {
                $resolved = (& py -3.12 -c "import sys; print(sys.executable)" 2>&1 | Select-Object -Last 1).Trim()
                if ($resolved -and (Test-PythonUsable $resolved)) {
                    return $resolved
                }
            }
        } catch {
        }
    }

    return $null
}

function Install-Python {
    Write-Section "Встановлення Python"
    Write-Info "Python 3.12 не знайдено на цьому комп'ютері."
    Write-Info "Зараз буде завантажено офіційний встановлювач Python $PythonVersion з python.org:"
    Write-Info "  $PythonInstallerUrl"
    Write-Info "Встановлення відбудеться ТІЛЬКИ для вашого користувача, без прав адміністратора."
    Start-Sleep -Seconds 3

    if (-not (Test-PathWritable $env:TEMP)) {
        throw "Немає доступу на запис у тимчасову теку ($env:TEMP).$(Get-AccessDeniedGuidance -Path $env:TEMP)"
    }

    $installerPath = Join-Path $env:TEMP "python-$PythonVersion-amd64.exe"

    Write-Info "Завантаження встановлювача Python..."
    Invoke-WithRetry -Description "завантаження Python" -PathForGuidance $installerPath -Action {
        Invoke-WebRequest -Uri $PythonInstallerUrl -OutFile $installerPath -UseBasicParsing
    }

    Write-Info "Встановлення Python (це може зайняти кілька хвилин)..."
    $proc = Invoke-StepWithGuidance -PathForGuidance $installerPath -Action {
        Start-Process -FilePath $installerPath -ArgumentList @('/quiet', 'InstallAllUsers=0', 'PrependPath=1', 'Include_launcher=0') -Wait -PassThru
    }
    if ($proc.ExitCode -ne 0) {
        $extra = ""
        if ($proc.ExitCode -eq 1602) {
            $extra = " (код 1602 зазвичай означає, що встановлення було скасовано - у вікні встановлювача Python треба натиснути 'Так'/'Install', а не 'Скасувати'.)"
        }
        throw "Встановлення Python завершилось з кодом помилки $($proc.ExitCode).$extra"
    }

    Remove-Item -Path $installerPath -Force -ErrorAction SilentlyContinue

    # Встановлювач python.org з InstallAllUsers=0 завжди кладе Python за цим
    # фіксованим шляхом - простіше й надійніше перевірити його напряму, ніж
    # намагатися оновити PATH поточного процесу (PATH оновлюється лише для
    # НОВИХ процесів, а не для вже запущеного PowerShell).
    if (-not (Test-PythonUsable $KnownUserInstallPython)) {
        throw "Python нібито встановлено, але не вдалося знайти його за очікуваним шляхом ($KnownUserInstallPython)."
    }

    Write-Ok "Python успішно встановлено."
    return $KnownUserInstallPython
}

function Ensure-Venv {
    param([string]$PythonExe)
    if (Test-Path $VenvPython -PathType Leaf) {
        return
    }
    Write-Info "Створюю віртуальне середовище (venv)..."
    Invoke-StepWithGuidance -PathForGuidance $VenvDir -Action {
        & $PythonExe -m venv $VenvDir
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path $VenvPython -PathType Leaf)) {
            throw "Не вдалося створити віртуальне середовище (venv)."
        }
    }
    Write-Ok "Віртуальне середовище створено."
}

function Install-Dependencies {
    Write-Info "Встановлюю Python-бібліотеки (перший запуск або оновлені requirements.txt)."
    Write-Info "Це може зайняти кілька хвилин..."
    Invoke-WithRetry -Description "встановлення бібліотек (pip install)" -PathForGuidance $VenvDir -Action {
        & $VenvPython -m pip install --upgrade pip --quiet
        & $VenvPython -m pip install -r $RequirementsTxt
        if ($LASTEXITCODE -ne 0) {
            throw "Не вдалося встановити залежності (pip install)."
        }
    }
}

function Ensure-Dependencies {
    param([string]$PythonExe)

    $currentHash = (Get-FileHash -Path $RequirementsTxt -Algorithm SHA256).Hash
    $markerHash = $null
    if (Test-Path $MarkerFile) {
        $markerHash = (Get-Content -Path $MarkerFile -Raw).Trim()
    }

    if ($markerHash -eq $currentHash) {
        return
    }

    try {
        Install-Dependencies
    } catch {
        # Найчастіша причина ПОВТОРНОЇ (уже після retry) невдачі встановлення
        # бібліотек - наполовину пошкоджене venv від попереднього перерваного
        # запуску (напр. вимкнули комп'ютер посеред встановлення). venv тут -
        # це лише робоча "будівельна" тека проєкту (не дані користувача), тож
        # безпечно перестворити її з нуля один раз замість того, щоб просто
        # показати помилку й здатись.
        if (Test-IsAccessDeniedMessage $_.Exception.Message) {
            throw
        }
        Write-Warn "Не вдалося встановити бібліотеки навіть після кількох спроб."
        Write-Info "Схоже, віртуальне середовище пошкоджене - перестворюю його з нуля й пробую ще раз (це не зачіпає жодних ваших файлів у resources/output)..."
        try {
            Remove-Item -Path $VenvDir -Recurse -Force -ErrorAction Stop
        } catch {
            throw "Не вдалося видалити пошкоджене віртуальне середовище ($VenvDir), щоб перестворити його. Видаліть цю теку вручну і запустіть START.bat ще раз.`n$($_.Exception.Message)"
        }
        Ensure-Venv -PythonExe $PythonExe
        Install-Dependencies
    }

    Set-Content -Path $MarkerFile -Value $currentHash -NoNewline -Encoding utf8
    Write-Ok "Залежності встановлено."
}

# Крок запуску (START.bat, через launcher.ps1) - Python (встановити за
# потреби) + venv + бібліотеки, у ЦЬОМУ порядку.
function Ensure-PythonEnvironment {
    if (-not (Test-PathWritable $RepoRoot)) {
        throw "Немає доступу на запис у теку проєкту ($RepoRoot).$(Get-AccessDeniedGuidance -Path $RepoRoot)"
    }

    $python = Find-Python
    if (-not $python) {
        $python = Install-Python
    }

    Ensure-Venv -PythonExe $python
    Ensure-Dependencies -PythonExe $python
}

# ---------------------------------------------------------------------------
# Автооновлення коду проєкту з GitHub (перед запуском, launcher.ps1) - раніше
# був окремий файл ОНОВИТИ_ПРОГРАМУ_ДО_НОВОЇ_ВЕРСІЇ.bat, яким користувач мав
# ЗГАДАТИ запустити ОКРЕМО, ДО звичайного START.bat - на практиці про нього
# просто забували й місяцями працювали зі старою версією. Об'єднано в ОДИН
# файл (START.bat): подвійний клік на нього тепер САМ перевіряє й (за
# потреби) підтягує оновлення, а вже ПОТІМ запускає меню - нічого зайвого
# натискати не треба.
# ---------------------------------------------------------------------------

function Test-GitAvailable {
    return [bool](Get-Command git -ErrorAction SilentlyContinue)
}

function Test-GitUsable {
    param([string]$GitExe)
    if (-not $GitExe) { return $false }
    if (-not (Test-Path $GitExe -PathType Leaf)) { return $false }
    try {
        $out = & $GitExe --version 2>&1
        $text = ($out -join ' ')
        return ($LASTEXITCODE -eq 0 -and $text -match '(?i)^git version')
    } catch {
        return $false
    }
}

function Find-Git {
    if (Test-GitUsable $KnownPortableGitExe) {
        return $KnownPortableGitExe
    }
    $cmd = Get-Command git -ErrorAction SilentlyContinue
    if ($cmd -and (Test-GitUsable $cmd.Source)) {
        return $cmd.Source
    }
    return $null
}

function Install-Git {
    Write-Section "Встановлення Git"
    Write-Info "Git не знайдено на цьому комп'ютері - без нього неможливо перевіряти та встановлювати оновлення."
    Write-Info "Зараз буде завантажено портативну версію Git (без встановлювача, без прав адміністратора):"
    Write-Info "  $GitPortableUrl"
    Start-Sleep -Seconds 3

    if (-not (Test-PathWritable $env:TEMP)) {
        throw "Немає доступу на запис у тимчасову теку ($env:TEMP).$(Get-AccessDeniedGuidance -Path $env:TEMP)"
    }

    $sfxPath = Join-Path $env:TEMP 'PortableGit.7z.exe'

    Write-Info "Завантаження Git..."
    Invoke-WithRetry -Description "завантаження Git" -PathForGuidance $sfxPath -Action {
        Invoke-WebRequest -Uri $GitPortableUrl -OutFile $sfxPath -UseBasicParsing
    }

    Write-Info "Розпаковування Git..."
    if (Test-Path $KnownPortableGitDir) {
        Remove-Item -Path $KnownPortableGitDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    # SFX-розпаковувач НЕ створює вкладені теки призначення сам - без цього
    # кроку розпаковування мовчки провалюється (перевірено на практиці).
    New-Item -ItemType Directory -Path $KnownPortableGitDir -Force | Out-Null

    $proc = Invoke-StepWithGuidance -PathForGuidance $KnownPortableGitDir -Action {
        Start-Process -FilePath $sfxPath -ArgumentList @('-y', "-o$KnownPortableGitDir") -Wait -PassThru -WindowStyle Hidden
    }
    if ($proc.ExitCode -ne 0) {
        throw "Розпаковування Git завершилось з кодом помилки $($proc.ExitCode)."
    }

    Remove-Item -Path $sfxPath -Force -ErrorAction SilentlyContinue

    if (-not (Test-GitUsable $KnownPortableGitExe)) {
        throw "Git нібито встановлено, але не вдалося знайти його за очікуваним шляхом ($KnownPortableGitExe)."
    }

    Write-Ok "Git успішно встановлено."
    return $KnownPortableGitExe
}

# Знаходить (чи, за потреби, встановлює) git і додає його теку в PATH ЦЬОГО
# процесу - усі виклики git у цьому файлі йдуть через голе "& git" (у
# Invoke-Git), тож досить оновити PATH ОДИН раз тут, а не переписувати кожен
# виклик на повний шлях. Оновлення PATH поточного (вже запущеного) процесу
# спрацьовує одразу - на відміну від PATH, який міг би виставити зовнішній
# встановлювач (той підхопився б лише в НОВОМУ процесі).
function Ensure-GitAvailable {
    $gitExe = Find-Git
    if (-not $gitExe) {
        $gitExe = Install-Git
    }
    $gitDir = Split-Path -Parent $gitExe
    if (($env:PATH -split ';') -notcontains $gitDir) {
        $env:PATH = "$gitDir;$env:PATH"
    }
}

function Test-NoInternetMessage {
    param([string]$Message)
    if (-not $Message) { return $false }
    return (
        $Message -match '(?i)could not resolve host' -or
        $Message -match '(?i)failed to connect' -or
        $Message -match '(?i)unable to access' -or
        $Message -match '(?i)connection (timed out|refused)' -or
        $Message -match '(?i)network is unreachable'
    )
}

# git часто пише НЕШКІДЛИВІ попередження (напр. про LF/CRLF) у stderr - "&
# git ... 2>&1" під $ErrorActionPreference = 'Stop' (вище в цьому файлі)
# перетворює КОЖЕН такий рядок stderr на PowerShell-виняток і ОБРИВАЄ
# виконання ще до перевірки $LASTEXITCODE, НАВІТЬ КОЛИ сам git завершився з
# кодом 0 - підтверджено тестуванням (git commit на щойно доданих файлах зі
# змішаними символами кінця рядка стабільно провокує це). Тому кожен виклик
# git - через цю обгортку: тимчасово послаблює ErrorActionPreference лише на
# час самого виклику, а успіх/невдачу визначає ВИКЛЮЧНО $LASTEXITCODE (як і
# мало б бути для звичайної консольної команди), а не наявність stderr-виводу.
function Invoke-Git {
    param([Parameter(Mandatory = $true)][string[]]$GitArgs)
    $prevPref = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    # Людина, для якої зроблено START.bat, не програміст - вона НІКОЛИ не
    # повинна побачити ЖОДНОГО вікна/запиту від git, хай там що:
    # - GIT_TERMINAL_PROMPT=0 - забороняє git чекати логін/пароль у консолі
    #   (текстовий запит користувач однаково не зможе заповнити);
    # - GCM_INTERACTIVE=Never - забороняє Git Credential Manager відкривати
    #   ГРАФІЧНЕ вікно логіна/браузер OAuth, якщо збереженого облікового
    #   запису ще немає - GIT_TERMINAL_PROMPT саме по собі це не зупинило б,
    #   бо стосується лише текстового запиту в консолі.
    #   ВАЖЛИВО: тут навмисно НЕ вимикається сам credential.helper
    #   ("-c credential.helper=") - якщо на машині користувача вже є
    #   ЗБЕРЕЖЕНИЙ (кешований) обліковий запис для github.com, саме він
    #   дозволяє fetch/merge пройти повністю мовчки, без жодного вікна.
    #   Вимкнення helper'а робить це неможливим і ламає оновлення навіть
    #   тоді, коли автентифікація насправді вже налаштована й працює
    #   (підтверджено на практиці: "fatal: could not read Username for
    #   'https://github.com': terminal prompts disabled").
    # - "-c safe.directory=$RepoRoot" - обходить окрему й доволі поширену
    #   "пастку" git на Windows: якщо тека проєкту належить ІНШОМУ
    #   користувачу Windows (напр. після копіювання профілю, чи через
    #   OneDrive/мережевий диск) - git з міркувань безпеки відмовляється
    #   працювати з повідомленням "dubious ownership", яке вимагає від
    #   людини самої ввести команду в консоль git, щоб "підтвердити" довіру
    #   до теки - те, чого вона просто не зможе зробити сама.
    $prevTerminalPrompt = $env:GIT_TERMINAL_PROMPT
    $prevGcmInteractive = $env:GCM_INTERACTIVE
    $env:GIT_TERMINAL_PROMPT = '0'
    $env:GCM_INTERACTIVE = 'Never'
    try {
        $output = @(& git -c "safe.directory=$RepoRoot" @GitArgs 2>&1 | ForEach-Object { $_.ToString() })
        $exitCode = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $prevPref
        $env:GIT_TERMINAL_PROMPT = $prevTerminalPrompt
        $env:GCM_INTERACTIVE = $prevGcmInteractive
    }
    return [PSCustomObject]@{ Output = $output; Text = ($output -join "`n"); ExitCode = $exitCode }
}

# Автооновлення потребує лише git на комп'ютері - НЕ обов'язково теки .git
# у проєкті: копія, розпакована з ZIP-архіву (найпростіший спосіб для
# нетехнічного користувача - завантажив ZIP із GitHub і розпакував, БЕЗ
# .git взагалі, незалежно від того, з якого репозиторію завантажив), теж
# підлягає оновленню - Update-Repository нижче сама перетворює її на
# git-репозиторій, що стежить за публічним дзеркалом (Ensure-GitRepository).
# Відсутність git на комп'ютері - НЕ помилка користувача, тому
# Update-IfPossible нижче мовчки пропускає крок оновлення в цьому випадку,
# замість лякати повідомленням про "помилку" на КОЖНОМУ запуску.
function Test-CanAutoUpdate {
    return (Test-GitAvailable)
}

# Перетворює теку БЕЗ .git (розпакований ZIP-архів) на git-репозиторій, що
# стежить за публічним дзеркалом - викликається лише коли .git ще не існує.
# Без цього кроку копія, роздана НЕ через "git clone" (а саме так її й
# отримав реальний кінцевий користувач - developer завантажив ZIP і
# переслав) НІКОЛИ не змогла б оновитись: Test-CanAutoUpdate раніше саме
# через відсутність .git тихо пропускала оновлення взагалі, без жодного
# повідомлення - зовні виглядало так, ніби подія оновлення не спрацювала.
#
# Поточний вміст теки (той самий ZIP-знімок) зберігається одним початковим
# комітом - не тому, що він має якусь цінність сам собою, а щоб було що
# порівнювати з дзеркалом і, за потреби, "врятувати" комітом, якщо в ньому
# раптом є щось відмінне від дзеркала (той самий механізм, що й для
# звичайних локальних правок). Update-Repository, що йде одразу після,
# сам підтягне цю копію до актуальної версії дзеркала - навіть у ПЕРШИЙ
# запуск.
function Ensure-GitRepository {
    if (Test-Path (Join-Path $RepoRoot '.git')) {
        return $true
    }
    try {
        $initResult = Invoke-Git -GitArgs @('-C', $RepoRoot, 'init', '--quiet', '-b', $UpdateBranch)
        if ($initResult.ExitCode -ne 0) {
            throw "git init: $($initResult.Text)"
        }
        Invoke-Git -GitArgs @('-C', $RepoRoot, 'remote', 'add', $UpdateRemoteName, $PublicMirrorUrl) | Out-Null

        $addResult = Invoke-Git -GitArgs @('-C', $RepoRoot, 'add', '-A')
        if ($addResult.ExitCode -ne 0) {
            throw "git add: $($addResult.Text)"
        }
        $hasStagedContent = (Invoke-Git -GitArgs @('-C', $RepoRoot, 'diff', '--cached', '--quiet')).ExitCode -ne 0
        if ($hasStagedContent) {
            $commitResult = Invoke-Git -GitArgs @(
                '-C', $RepoRoot,
                '-c', 'user.email=update-script@local', '-c', 'user.name=Generator Update Script', '-c', 'commit.gpgsign=false',
                'commit', '--quiet', '-m', 'Адаптовано з ZIP-архіву'
            )
            if ($commitResult.ExitCode -ne 0) {
                throw "git commit: $($commitResult.Text)"
            }
        }
        return $true
    } catch {
        # Якщо щось пішло не так (напр. немає інтернету на САМЕ цій спробі
        # ще до фактичного fetch) - прибираємо напівстворений .git, щоб
        # наступний запуск почав спробу з чистого аркуша, а не застряг у
        # зламаному напівстані.
        Remove-Item -Path (Join-Path $RepoRoot '.git') -Recurse -Force -ErrorAction SilentlyContinue
        return $false
    }
}

# Чи це робоча копія САМОГО розробника (origin вказує на ПРИВАТНИЙ
# репозиторій з реальною багатосотенною історією комітів) - а не копія
# кінцевого користувача (де джерело оновлень - публічне дзеркало, окремий
# репозиторій без спільної історії з приватним). Порівняння нормалізує адресу
# (регулярка виймає лише "власник/репозиторій"), а не звіряє рядок дослівно -
# інакше перемикання origin розробника з https:// на git@github.com: (форма
# SSH) мовчки "розгубило" б цю перевірку, і reset --hard нижче міг би
# зачепити справжню робочу гілку розробника.
function Test-IsDeveloperClone {
    $originResult = Invoke-Git -GitArgs @('-C', $RepoRoot, 'remote', 'get-url', 'origin')
    if ($originResult.ExitCode -ne 0) {
        return $false
    }
    $originUrl = $originResult.Text.Trim()
    if ($originUrl -match '(?i)github\.com[:/]+([^/]+/[^/.]+?)(\.git)?/?$') {
        return ($Matches[1] -ieq $PrivateRepoOwnerSlashName)
    }
    return $false
}

# Ідемпотентно: якщо remote "update-source" (джерело оновлень - публічне
# дзеркало) ще не існує чи вказує на іншу адресу - додає/виправляє його.
# Викликається лише для копій кінцевих користувачів (Test-IsDeveloperClone
# вже відсіяв копію розробника раніше в Update-Repository), тому тут НІКОЛИ
# не чіпає "origin" - той лишається чим був (чи його й зовсім немає).
function Ensure-UpdateRemote {
    $existing = Invoke-Git -GitArgs @('-C', $RepoRoot, 'remote', 'get-url', $UpdateRemoteName)
    if ($existing.ExitCode -ne 0) {
        Invoke-Git -GitArgs @('-C', $RepoRoot, 'remote', 'add', $UpdateRemoteName, $PublicMirrorUrl) | Out-Null
    } elseif ($existing.Text.Trim() -ne $PublicMirrorUrl) {
        Invoke-Git -GitArgs @('-C', $RepoRoot, 'remote', 'set-url', $UpdateRemoteName, $PublicMirrorUrl) | Out-Null
    }
}

# Перевіряє й, за потреби, оновлює сам код проєкту з ПУБЛІЧНОГО дзеркала (не
# з "origin" - той, якщо взагалі є, лишається недоторканим) - НЕ чіпає
# resources/output (вони поза git, .gitignore).
#
# За задумом користувач НЕ повинен ні знати, ні думати про те, що це git, що
# в нього можуть бути "незбережені зміни", чи що таке "конфлікт" - людина, для
# якої зроблено весь цей .bat-файл, не програміст і взагалі не має чути слово
# "git". Тому: якщо є власні зміни у файлах проєкту - скрипт МОВЧКИ зберігає
# їх звичайним commit (без жодного повідомлення користувачу про це), тоді
# ПОВНІСТЮ приймає дерево файлів з дзеркала (git reset --hard) - розробник має
# бути ВПЕВНЕНИЙ, що правка, яку він запушив у constants.py, дійде до кожного
# користувача при наступному відкритті START.bat, навіть якщо в когось
# локально випадково лишилась застаріла правка того самого рядка (а не тому,
# що користувач має свідомо редагувати ці файли сам - для цього нема
# легітимного сценарію: усі реальні дані користувача живуть поза git, у
# resources/output).
#
# Чому reset --hard, а НЕ merge: кожна синхронізація дзеркала (окремий
# GitHub Actions workflow у приватному репозиторії) - це НОВИЙ orphan-коміт
# без спільного предка з тим, що вже є в користувача локально (навмисно - щоб
# дзеркало ніколи не показувало git-історію приватного репозиторію). "git
# merge" без спільного предка або взагалі відмовляється працювати, або (з
# --allow-unrelated-histories) трактує "файл, якого нема в новій версії, але
# є локально" як "додано лише в нас" - НЕ як конфлікт, тобто такі файли
# накопичувались б назавжди й ніколи не видалялись. "reset --hard" такого
# недоліку не має: робочий каталог стає ТОЧНОЮ копією опублікованого дерева,
# незалежно від походження коміту.
function Update-Repository {
    if (-not (Test-PathWritable $RepoRoot)) {
        throw "Немає доступу на запис у теку проєкту ($RepoRoot).$(Get-AccessDeniedGuidance -Path $RepoRoot)"
    }

    if (-not (Test-Path (Join-Path $RepoRoot '.git'))) {
        # Розпакований ZIP-архів (без .git) - перетворюємо на git-репозиторій,
        # що стежить за публічним дзеркалом. Якщо не вдалось (напр. немає
        # інтернету саме зараз) - мовчки виходимо, як і раніше: копія лишається
        # робочою, просто без оновлення цього разу.
        if (-not (Ensure-GitRepository)) {
            return
        }
    } elseif (Test-IsDeveloperClone) {
        # На власній машині розробника (origin = приватний репозиторій)
        # НІКОЛИ не робимо reset гілки main до чужого дзеркала - там живе
        # справжня історія, яку розробник пушить у origin вручну; reset
        # --hard до непов'язаного дзеркала зламав би її.
        return
    }

    Ensure-UpdateRemote

    Write-Info "Перевіряю, чи є оновлення..."
    Invoke-WithRetry -Description "перевірку оновлень" -Action {
        $result = Invoke-Git -GitArgs @('-C', $RepoRoot, 'fetch', '--quiet', $UpdateRemoteName, $UpdateBranch)
        if ($result.ExitCode -ne 0) {
            if (Test-NoInternetMessage $result.Text) {
                throw "Не вдалося перевірити оновлення - немає з'єднання з інтернетом (або GitHub тимчасово недоступний). Перевірте інтернет і спробуйте ще раз.`n$($result.Text)"
            }
            throw "Не вдалося перевірити оновлення.`n$($result.Text)"
        }
    }

    $remoteRef = "$UpdateRemoteName/$UpdateBranch"
    $localTree = (Invoke-Git -GitArgs @('-C', $RepoRoot, 'rev-parse', 'HEAD^{tree}')).Text.Trim()
    $remoteTreeResult = Invoke-Git -GitArgs @('-C', $RepoRoot, 'rev-parse', "$remoteRef^{tree}")
    if ($remoteTreeResult.ExitCode -ne 0) {
        throw "Не вдалося перевірити оновлення. Зверніться по допомогу."
    }
    $remoteTree = $remoteTreeResult.Text.Trim()

    # Порівняння хешів ДЕРЕВ (а не комітів/предків) - навмисно: кожна
    # синхронізація дзеркала - це НОВИЙ orphan-коміт без спільного предка з
    # тим, що вже є локально, тому HEAD ніколи не буде "нащадком" upstream у
    # звичайному git-розумінні. Хеш дерева натомість залежить ЛИШЕ від
    # вмісту файлів, тож два незалежні orphan-коміти з однаковим вмістом
    # дають ОДНАКОВИЙ хеш дерева - це і є коректний критерій "нічого нового
    # немає". ВАЖЛИВО: якщо дерева співпадають - виходимо ОДРАЗУ, НЕ чіпаючи
    # жодних локальних незбережених правок (навіть комітом) - бо коміт правок
    # тут зсунув би локальне дерево ВБІК від дзеркала, і НАСТУПНИЙ запуск (без
    # жодного реального оновлення) хибно прочитав би це як "є нова версія" й
    # стер би саме ті правки, які щойно "врятував" би цей коміт.
    if ($localTree -eq $remoteTree) {
        Write-Ok "Скрипт оновлено до актуальної версії."
        return
    }

    Write-Info "Знайдено нову версію - оновлюю..."

    $hasLocalChanges = [bool]((Invoke-Git -GitArgs @('-C', $RepoRoot, 'status', '--porcelain')).Output | Where-Object { $_ })
    if ($hasLocalChanges) {
        $addResult = Invoke-Git -GitArgs @('-C', $RepoRoot, 'add', '-A')
        if ($addResult.ExitCode -ne 0) {
            throw "Не вдалося оновити.`n$($addResult.Text)"
        }
        # -c user.email/user.name (не глобальний конфіг git) - щоб працювало
        # навіть на щойно встановленому git, де особу автора ще не задано.
        # -c commit.gpgsign=false - якщо на цьому комп'ютері десь глобально
        # увімкнено підпис комітів (напр. корпоративна політика чи власні
        # налаштування розробника) - без цього git чекав би пароль GPG-ключа
        # у спливному вікні (pinentry) САМЕ на цьому внутрішньому,
        # автоматичному коміті, де підпис і не потрібен.
        $commitResult = Invoke-Git -GitArgs @(
            '-C', $RepoRoot,
            '-c', 'user.email=update-script@local', '-c', 'user.name=Generator Update Script', '-c', 'commit.gpgsign=false',
            'commit', '--quiet', '-m', "Автозбереження перед оновленням ($(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))"
        )
        if ($commitResult.ExitCode -ne 0) {
            throw "Не вдалося оновити.`n$($commitResult.Text)"
        }
    }

    Invoke-WithRetry -Description "застосування оновлення" -Action {
        $resetResult = Invoke-Git -GitArgs @('-C', $RepoRoot, 'reset', '--hard', '--quiet', $remoteRef)
        if ($resetResult.ExitCode -ne 0) {
            if (Test-NoInternetMessage $resetResult.Text) {
                throw "Не вдалося завантажити оновлення - немає з'єднання з інтернетом. Перевірте інтернет і спробуйте ще раз.`n$($resetResult.Text)"
            }
            throw "Не вдалося оновити.`n$($resetResult.Text)"
        }
    }

    Write-Ok "Оновлено до останньої версії."
}

# Викликається З launcher.ps1 ПЕРЕД Ensure-PythonEnvironment (щоб код і
# requirements.txt були вже актуальні до встановлення бібліотек) - НІКОЛИ не
# кидає виняток далі (на відміну від Update-Repository): невдале оновлення
# (немає інтернету, GitHub тимчасово недоступний тощо) не повинно заблокувати
# сам ЗАПУСК програми - людина має змогу далі працювати з тією версією, що
# вже встановлена, а не впиратись у стіну через тимчасову мережеву проблему.
function Update-IfPossible {
    try {
        Ensure-GitAvailable
    } catch {
        # Git не знайдено і не вдалось встановити (напр. немає інтернету на
        # ЦЬОМУ комп'ютері ВЗАГАЛІ, ще до першого запуску) - без git
        # оновлювати нічим, тихо пропускаємо крок оновлення й одразу йдемо
        # до самої програми, а не блокуємо запуск через це.
        return
    }
    if (-not (Test-CanAutoUpdate)) {
        return
    }
    try {
        Update-Repository
    } catch {
        Write-Warn "Не вдалося перевірити/встановити оновлення - запускаю з тією версією, що вже встановлена."
        Write-Info $_.Exception.Message
        Start-Sleep -Seconds 2
    }
}

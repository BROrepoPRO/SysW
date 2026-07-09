<#
.SYNOPSIS
    Inicializaciya Git-repozitoriya dlya proekta SysW i otpravka na GitHub.
.DESCRIPTION
    Skript vypolnyaet polnuyu nastrojku Git-repozitoriya.
.NOTES
    Avtor: SourceCraft
    Versiya: 1.0
#>

# ============================================================
# Funkciya dlya vyvoda soobsheniya ob oshibke i vyhoda
# ============================================================
function Write-ErrorAndExit {
    param([string]$Message)
    Write-Host "[OSHIBKA] $Message" -ForegroundColor Red
    exit 1
}

# ============================================================
# Funkciya dlya vyvoda informacionnogo soobsheniya
# ============================================================
function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

# ============================================================
# Funkciya dlya vyvoda soobsheniya ob uspehe
# ============================================================
function Write-Success {
    param([string]$Message)
    Write-Host "[OK] $Message" -ForegroundColor Green
}

# ============================================================
# Shag 0: Proverka nalichiya Git v sisteme
# ============================================================
Write-Info "Proverka nalichiya Git..."

try {
    $gitVersion = git --version
    Write-Success "Git nayden: $gitVersion"
} catch {
    Write-ErrorAndExit "Git ne ustanovlen ili ne dobavlen v PATH. Ustanovite Git s https://git-scm.com/ i povtorite popytku."
}

# ============================================================
# Shag 1: Proverka i perehod v direktoriyu L:\PROject\SysW
# ============================================================
$targetDir = "L:\PROject\SysW"

Write-Info "Proverka direktori: $targetDir"

if (-not (Test-Path $targetDir -PathType Container)) {
    Write-ErrorAndExit "Papka '$targetDir' ne naidena. Ubedites, chto put sushestvuet."
}

Set-Location -Path $targetDir
Write-Success "Tekushaya direktoriya: $(Get-Location)"

# ============================================================
# Shag 2: Inicializaciya Git-repozitoriya
# ============================================================
Write-Info "Inicializaciya Git-repozitoriya..."

if (Test-Path ".git" -PathType Container) {
    Write-Host "[WARN] Repozitoriy uze inicializirovan (.git sushestvuet). Propuskaem git init." -ForegroundColor Yellow
} else {
    git init
    if ($LASTEXITCODE -ne 0) {
        Write-ErrorAndExit "Ne udalos inicializirovat Git-repozitoriy."
    }
    Write-Success "Git-repozitoriy inicializirovan."
}

# ============================================================
# Shag 3: Sozdanie fayla .gitignore
# ============================================================
Write-Info "Sozdanie fayla .gitignore..."

$gitignoreContent = @"
# ============================================
# .gitignore dlya VBA / Excel proektov
# ============================================

# Vremennye fayly Excel (lock-fayly)
~$*.xls*
*.~lock.*

# Vremennye fayly
*.tmp
*.temp

# Fayly blokirovki Microsoft Office
~*.*

# Log-fayly
*.log

# Bazy dannyh SQLite (esli ispolzuyutsya kak kesh)
*.db
*.sqlite

# Sistemnye fayly Windows
Thumbs.db
Desktop.ini
$RECYCLE.BIN/

# Papki sborok i kesha (esli poyavyatsya)
bin/
obj/
Debug/
Release/
.vs/
*.user
*.suo
*.sln.docstates

# Papka node_modules (na sluchay, esli poyavitsya)
node_modules/

# Fayly konfiguracii IDE
.idea/
.vscode/
*.swp
*.swo

# Fayly s sekretami/klyuchami (esli poyavyatsya)
*.key
*.pem
secrets/
"@

try {
    [System.IO.File]::WriteAllText((Join-Path $targetDir ".gitignore"), $gitignoreContent, [System.Text.Encoding]::UTF8)
    Write-Success "Fayl .gitignore sozdan."
} catch {
    Write-ErrorAndExit "Ne udalos sozdat .gitignore: $_"
}

# ============================================================
# Shag 4: Pereimenovanie vetki v main
# ============================================================
Write-Info "Pereimenovanie vetki v 'main'..."

try {
    git branch -M main
    if ($LASTEXITCODE -ne 0) {
        throw "Oshibka pri pereimenovanii vetki."
    }
    Write-Success "Vetka pereimenovana v 'main'."
} catch {
    Write-ErrorAndExit "Ne udalos pereimenovat vetku: $_"
}

# ============================================================
# Shag 5: Dobavlenie udalonnogo repozitoriya
# ============================================================
Write-Info "Dobavlenie udalennogo repozitoriya: origin..."

try {
    $existingRemote = git remote get-url origin 2>$null
    if ($existingRemote) {
        Write-Host "[WARN] Udalonny repozitoriy 'origin' uze sushestvuet: $existingRemote" -ForegroundColor Yellow
        if ($existingRemote -ne "https://github.com/BROrepoPRO/SysW") {
            Write-Host "[WARN] URL otlichaetsya ot ozhidaemogo. Menyaem na https://github.com/BROrepoPRO/SysW..." -ForegroundColor Yellow
            git remote set-url origin "https://github.com/BROrepoPRO/SysW"
            Write-Success "URL udalonnogo repozitoriya obnovlen."
        }
    } else {
        git remote add origin "https://github.com/BROrepoPRO/SysW"
        if ($LASTEXITCODE -ne 0) {
            throw "Oshibka pri dobavlenii udalonnogo repozitoriya."
        }
        Write-Success "Udalonny repozitoriy dobavlen: https://github.com/BROrepoPRO/SysW"
    }
} catch {
    Write-ErrorAndExit "Ne udalos dobavit udalonny repozitoriy: $_"
}

# ============================================================
# Shag 6: Dobavlenie vseh faylov v indeks
# ============================================================
Write-Info "Dobavlenie vseh faylov v indeks (git add .)..."

try {
    git add .
    if ($LASTEXITCODE -ne 0) {
        throw "Oshibka pri vypolnenii git add."
    }
    Write-Success "Vse fayly dobavleny v indeks."
} catch {
    Write-ErrorAndExit "Ne udalos dobavit fayly: $_"
}

# ============================================================
# Shag 6.5: Nastrojka Git-konfiga (user.name / user.email)
# ============================================================
Write-Info "Proverka i nastrojka Git-konfiga (user.name / user.email)..."

try {
    $currentName = git config user.name
    $currentEmail = git config user.email

    if (-not $currentName) {
        git config user.name "BROrepoPRO"
        Write-Host "[WARN] user.name ne byl ustanovlen. Ustanovlen: 'BROrepoPRO'" -ForegroundColor Yellow
    } else {
        Write-Success "user.name: $currentName"
    }

    if (-not $currentEmail) {
        git config user.email "brorepopro@users.noreply.github.com"
        Write-Host "[WARN] user.email ne byl ustanovlen. Ustanovlen: 'brorepopro@users.noreply.github.com'" -ForegroundColor Yellow
    } else {
        Write-Success "user.email: $currentEmail"
    }
} catch {
    Write-ErrorAndExit "Ne udalos nastroit Git-konfig: $_"
}

# ============================================================
# Shag 7: Sozdanie pervogo commita
# ============================================================
Write-Info "Sozdanie pervogo commita..."

try {
    $status = git status --porcelain
    if (-not $status) {
        Write-Host "[WARN] Net izmeneniy dlya commita. Vozmozno, commit uze sozdan." -ForegroundColor Yellow
    } else {
        git commit -m "Initial commit: struktura proekta SysW"
        if ($LASTEXITCODE -ne 0) {
            throw "Oshibka pri sozdanii commita."
        }
        Write-Success "Pervy commit sozdan."
    }
} catch {
    Write-ErrorAndExit "Ne udalos sozdat commit: $_"
}

# ============================================================
# Shag 8: Otpravka vetki main na GitHub
# ============================================================
Write-Info "Otpravka vetki 'main' na GitHub (git push -u origin main)..."

Write-Host ""
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host "  VNIMANIE: Mozhet potrebovatsya autentifikaciya GitHub!" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Yellow
Write-Host ""
Write-Host "Esli Git zaprosit login/parol - ispolzuyte:" -ForegroundColor White
Write-Host "  - Login: vash GitHub username" -ForegroundColor White
Write-Host "  - Parol: Personal Access Token (ne parol ot akkounta!)" -ForegroundColor White
Write-Host ""
Write-Host "Kak sozdat Personal Access Token:" -ForegroundColor White
Write-Host "  1. GitHub -> Settings -> Developer settings -> Personal access tokens -> Tokens (classic)" -ForegroundColor White
Write-Host "  2. Nagmite 'Generate new token (classic)'" -ForegroundColor White
Write-Host "  3. Vyberite scope: repo (polny dostup)" -ForegroundColor White
Write-Host "  4. Skopiruyte token i ispolzuyte ego kak parol" -ForegroundColor White
Write-Host ""

try {
    git push -u origin main
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "[OSHIBKA] Ne udalos otpravit vetku na GitHub." -ForegroundColor Red
        Write-Host "Vozmoznye prichiny:" -ForegroundColor Yellow
        Write-Host "  1. Ne proydena autentifikaciya - vvedite login i Personal Access Token." -ForegroundColor Yellow
        Write-Host "  2. Repozitoriy https://github.com/BROrepoPRO/SysW ne sushestvuet - sozdajte ego vruchnuyu." -ForegroundColor Yellow
        Write-Host "  3. Net dostupa k internetu ili GitHub zablokirovan." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Posle resheniya problemy vypolnite komandu vruchnuyu:" -ForegroundColor White
        Write-Host "  git push -u origin main" -ForegroundColor White
        exit 1
    }
    Write-Success "Vetka 'main' uspeshno otpravlena na GitHub!"
} catch {
    Write-Host ""
    Write-Host "[OSHIBKA] Isklyuchenie pri push: $_" -ForegroundColor Red
    Write-Host "Posle resheniya problemy vypolnite komandu vruchnuyu:" -ForegroundColor White
    Write-Host "  git push -u origin main" -ForegroundColor White
    exit 1
}

# ============================================================
# Zavershenie
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Inicializaciya Git-repozitoriya uspeshno zavershena!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Repozitoriy: https://github.com/BROrepoPRO/SysW" -ForegroundColor White
Write-Host "Vetka: main" -ForegroundColor White
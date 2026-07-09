<#
.SYNOPSIS
    Inicializaciya Git-repozitoriya dlya proekta SysW i otpravka na GitHub.
.DESCRIPTION
    Skript vypolnyaet polnuyu nastrojku Git-repozitoriya soglasno promtu.
.NOTES
    Avtor: SourceCraft Code Assistant
    Versiya: 1.1
#>

# ============================================================
# Funkciya dlya vyvoda soobsheniya ob oshibke i vyhoda
# ============================================================
function Write-ErrorAndExit {
    param([string]$Message)
    Write-Host ""
    Write-Host "[OSHIBKA] $Message" -ForegroundColor Red
    Write-Host ""
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
# Shag 1: Perehod v papku L:\PROject\SysW
# ============================================================
Write-Info "Shag 1. Perehod v papku L:\PROject\SysW..."

$targetDir = "L:\PROject\SysW"

if (-not (Test-Path $targetDir -PathType Container)) {
    Write-ErrorAndExit "Papka '$targetDir' ne sushestvuet. Skript ostanovlen."
}

Set-Location -Path $targetDir
Write-Success "Tekushaya direktoriya: $(Get-Location)"
Write-Host ""

# ============================================================
# Shag 2: Proverka, chto Git ustanovlen
# ============================================================
Write-Info "Shag 2. Proverka nalichiya Git v sisteme..."

try {
    $gitVersion = git --version
    Write-Success "Git nayden: $gitVersion"
} catch {
    Write-ErrorAndExit "Git ne ustanovlen ili ne dobavlen v PATH."
}
Write-Host ""

# ============================================================
# Shag 3: Inicializaciya Git-repozitoriya
# ============================================================
Write-Info "Shag 3. Inicializaciya Git-repozitoriya (git init)..."

if (Test-Path ".git" -PathType Container) {
    Write-Host "[WARN] Repozitoriy uze inicializirovan (.git sushestvuet)." -ForegroundColor Yellow
} else {
    git init
    if ($LASTEXITCODE -ne 0) {
        Write-ErrorAndExit "Ne udalos inicializirovat Git-repozitoriy."
    }
    Write-Success "Git-repozitoriy inicializirovan."
}
Write-Host ""

# ============================================================
# Shag 4: Sozdanie fayla .gitignore
# ============================================================
Write-Info "Shag 4. Sozdanie fayla .gitignore..."

$gitignoreContent = @"
~$*.xls*
*.tmp
~*.*
*.log
*.db
*.sqlite
Thumbs.db
Desktop.ini
node_modules/
.yca/
"@

try {
    [System.IO.File]::WriteAllText((Join-Path $targetDir ".gitignore"), $gitignoreContent, [System.Text.Encoding]::UTF8)
    Write-Success "Fayl .gitignore sozdan."
} catch {
    Write-ErrorAndExit "Ne udalos sozdat .gitignore: $_"
}
Write-Host ""

# ============================================================
# Shag 5: Nastrojka lokalnyh parametrov polzovatelya Git
# ============================================================
Write-Info "Shag 5. Nastrojka lokalnyh parametrov polzovatelya Git..."

try {
    $currentName = git config user.name
    $currentEmail = git config user.email

    if (-not $currentName) {
        git config user.name "BROrepoPRO"
        Write-Success "user.name ustanovlen: 'BROrepoPRO'"
    } else {
        Write-Host "[INFO] user.name uze zadan: $currentName" -ForegroundColor White
    }

    if (-not $currentEmail) {
        git config user.email "karambos1984@gmail.com"
        Write-Success "user.email ustanovlen: 'karambos1984@gmail.com'"
    } else {
        Write-Host "[INFO] user.email uze zadan: $currentEmail" -ForegroundColor White
    }
} catch {
    Write-ErrorAndExit "Ne udalos nastroit Git-konfig: $_"
}
Write-Host ""

# ============================================================
# Shag 6: Pereimenovanie tekushej vetki v main
# ============================================================
Write-Info "Shag 6. Pereimenovanie tekushej vetki v 'main'..."

try {
    git branch -M main
    if ($LASTEXITCODE -ne 0) {
        throw "Oshibka pri pereimenovanii vetki."
    }
    Write-Success "Tekushaya vetka pereimenovana v 'main'."
} catch {
    Write-ErrorAndExit "Ne udalos pereimenovat vetku: $_"
}
Write-Host ""

# ============================================================
# Shag 7: Dobavlenie udalonnogo repozitoriya
# ============================================================
Write-Info "Shag 7. Dobavlenie udalonnogo repozitoriya origin..."

$remoteUrl = "https://github.com/BROrepoPRO/SysW"

try {
    $existingRemote = git remote get-url origin 2>$null
    if ($existingRemote) {
        Write-Host "[WARN] Udalonny repozitoriy 'origin' uze sushestvuet: $existingRemote" -ForegroundColor Yellow
        if ($existingRemote -ne $remoteUrl) {
            Write-Host "[WARN] URL otlichaetsya. Menyaem na $remoteUrl ..." -ForegroundColor Yellow
            git remote set-url origin $remoteUrl
            Write-Success "URL udalonnogo repozitoriya obnovlen."
        } else {
            Write-Host "[INFO] URL uze sootvetstvuet ozhidaemomu." -ForegroundColor White
        }
    } else {
        git remote add origin $remoteUrl
        if ($LASTEXITCODE -ne 0) {
            throw "Oshibka pri dobavlenii udalonnogo repozitoriya."
        }
        Write-Success "Udalonny repozitoriy dobavlen: $remoteUrl"
    }
} catch {
    Write-ErrorAndExit "Ne udalos dobavit udalonny repozitoriy: $_"
}
Write-Host ""

# ============================================================
# Shag 8: Sozdanie README.md
# ============================================================
Write-Info "Shag 8. Sozdanie README.md..."

$readmeContent = @"
# SysW - Sistema avtomatizacii obrabotki zakaz-naryadov avtoremonta

**Naznachenie:** import, analiz i uchet dannyh zakaz-naryadov iz Excel v SQLite.

**Tehnicheskiy sostav:**
- Moduli VBA dlya Excel (import, parsing, vygruzka)
- Baza dannyh SQLite (hranenie zakaz-naryadov, rabot, zapchastey)
- PowerShell-skripty avtomatizacii
- Integraciya s Git i GitHub dlya kontrolya versiy

**Komanda proekta:**

| Rol | Obyazannosti |
|------|-------------|
| **Nachalnik mira** | Generaciya idey, postanovka zadach, klyuchevye resheniya, arhitektura i prioritety. |
| **Glavny pomoshnik** | Koordinaciya, dekompoziciya zadach, promty dlya Code Assistant, kontrol kachestva, integraciya rezultatov. |
| **Code Assistant (DevOps-komanda)** | Napisanie koda, dokumentacii, testov, avtomatizaciya po gotovym promtam. |

**Ispolzuemye tehnologii:** VBA, ADO, SQLite, Excel, Git, PowerShell, SourceCraft Code Assistant.
"@

try {
    [System.IO.File]::WriteAllText((Join-Path $targetDir "README.md"), $readmeContent, [System.Text.Encoding]::UTF8)
    Write-Success "Fayl README.md sozdan."
} catch {
    Write-ErrorAndExit "Ne udalos sozdat README.md: $_"
}
Write-Host ""

# ============================================================
# Shag 9: Dobavlenie vseh faylov v indeks
# ============================================================
Write-Info "Shag 9. Dobavlenie vseh faylov v indeks (git add .)..."

try {
    git add .
    if ($LASTEXITCODE -ne 0) {
        throw "Oshibka pri vypolnenii git add."
    }
    Write-Success "Vse fayly dobavleny v indeks."
} catch {
    Write-ErrorAndExit "Ne udalos dobavit fayly: $_"
}
Write-Host ""

# ============================================================
# Shag 10: Pervy commit
# ============================================================
Write-Info "Shag 10. Sozdanie pervogo commita..."

try {
    $status = git status --porcelain
    if (-not $status) {
        Write-Host "[WARN] Net izmenenij dlya commita. Vozmozhno, commit uze sozdan." -ForegroundColor Yellow
    } else {
        git commit -m "Initial commit: struktura proekta SysW i README"
        if ($LASTEXITCODE -ne 0) {
            throw "Oshibka pri sozdanii commita."
        }
        Write-Success "Pervyj commit sozdan."
    }
} catch {
    Write-ErrorAndExit "Ne udalos sozdat commit: $_"
}
Write-Host ""

# ============================================================
# Shag 11: Proverka dostupnosti udalonnogo repozitoriya i push
# ============================================================
Write-Info "Shag 11. Proverka dostupnosti udalonnogo repozitoriya (git ls-remote origin)..."

try {
    $lsRemote = git ls-remote origin 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "[OSHIBKA] Ne udalos podklyuchitsya k udalennomu repozitoriyu." -ForegroundColor Red
        Write-Host "Trebuetsya autentifikaciya na GitHub. Vypolnite vhod i povtorite push vruchnuyu." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Dlya ruchnogo pusha vypolnite:" -ForegroundColor White
        Write-Host "  git push -u origin main" -ForegroundColor White
        Write-Host ""
        exit 1
    }
    Write-Success "Udalonny repozitoriy dostupen."

    # Vypolnyaem push
    Write-Info "Otpravka vetki 'main' na GitHub (git push -u origin main)..."
    git push -u origin main
    if ($LASTEXITCODE -ne 0) {
        Write-Host ""
        Write-Host "[OSHIBKA] Ne udalos otpravit vetku na GitHub." -ForegroundColor Red
        Write-Host "Trebuetsya autentifikaciya na GitHub. Vypolnite vhod i povtorite push vruchnuyu." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Dlya ruchnogo pusha vypolnite:" -ForegroundColor White
        Write-Host "  git push -u origin main" -ForegroundColor White
        Write-Host ""
        exit 1
    }
    Write-Success "Vetka 'main' uspeshno otpravlena na GitHub!"
} catch {
    Write-Host ""
    Write-Host "[OSHIBKA] Isklyuchenie pri proverke ili otpravke: $_" -ForegroundColor Red
    Write-Host "Trebuetsya autentifikaciya na GitHub. Vypolnite vhod i povtorite push vruchnuyu." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Dlya ruchnogo pusha vypolnite:" -ForegroundColor White
    Write-Host "  git push -u origin main" -ForegroundColor White
    Write-Host ""
    exit 1
}
Write-Host ""

# ============================================================
# Shag 12: Sozdanie vetki dev ot main
# ============================================================
Write-Info "Shag 12. Sozdanie vetki 'dev' ot 'main'..."

try {
    $branchExists = git branch --list dev
    if ($branchExists) {
        Write-Host "[WARN] Vetka 'dev' uze sushestvuet. Perekljuchaemsya na nee." -ForegroundColor Yellow
        git checkout dev
    } else {
        git checkout -b dev
        if ($LASTEXITCODE -ne 0) {
            throw "Oshibka pri sozdanii vetki dev."
        }
        Write-Success "Vetka 'dev' sozdana ot 'main' i vypolneno perekljuchenie na nee."
    }
} catch {
    Write-ErrorAndExit "Ne udalos sozdat vetku dev: $_"
}
Write-Host ""

# ============================================================
# Shag 13: Vyvod tekushego statusa i loga
# ============================================================
Write-Info "Shag 13. Vyvod tekushego statusa i istorii..."

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  git status" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
git status

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  git log --oneline --all --graph" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
git log --oneline --all --graph

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Inicializaciya Git-repozitoriya uspeshno zavershena!" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Repozitoriy: https://github.com/BROrepoPRO/SysW" -ForegroundColor White
Write-Host "Vetka: dev (sozdana ot main)" -ForegroundColor White
Write-Host ""

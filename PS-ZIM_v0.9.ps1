#Requires -Version 5.1
<#
.SYNOPSIS
  PS-ZIM — Lecteur de fichiers ZIM en PowerShell pur, servi en HTTP local.

.DESCRIPTION
  Ouvre un fichier .zim (Wikipedia offline, Stack Exchange dump, etc.), monte
  un mini serveur HTTP local et lance le navigateur. Décompresse à la volée
  les clusters Zstandard (Kiwix moderne) et clusters non compressés.

  Fonctionnalités :
    - Parseur ZIM binaire complet (header, MIME list, dir entries, clusters)
    - Décompression Zstandard via libzstd.dll natif (P/Invoke, auto-téléchargée au 1er run)
    - Vérification SHA-256 optionnelle (épinglée) du binaire natif téléchargé
    - Cache LRU de clusters décompressés (configurable)
    - Serveur HTTP concurrent (pool de runspaces) avec accès flux thread-safe
    - ETag / 304 Not Modified sur les articles
    - Recherche par préfixe de titre (binary search) + repli URL pour ZIM v6.3+
    - Endpoint d'autocomplétion /api/suggest + suggestions dans la barre de nav
    - Article aléatoire, page d'accueil avec métadonnées, top-bar injectée
    - Mode foreground ou background (-Background) avec PID file
    - Compatible ZIM v5 (legacy namespaces) et v6 (C namespace)

  Auteur : Eric Guiffault
  Version : 0.9.0

.PARAMETER ZimPath
  Chemin du fichier .zim à servir.

.PARAMETER Port
  Port HTTP local (défaut : 8642).

.PARAMETER BindAddress
  Interface d'écoute (défaut : 127.0.0.1). Mettre 0.0.0.0 pour exposer au LAN.

.PARAMETER CacheSize
  Nombre de clusters décompressés gardés en mémoire (défaut : 64).
  Chaque cluster fait en moyenne 1–4 Mo après décompression.

.PARAMETER MaxThreads
  Nombre maximal de requêtes traitées en parallèle (défaut : auto = nb de cœurs).
  Ignoré si -Sequential est utilisé.

.PARAMETER Sequential
  Désactive la concurrence : traite les requêtes une par une (mode robuste de repli).

.PARAMETER DebugLog
  Active la journalisation détaillée dans ps-zim-debug.log (désactivée par défaut).

.PARAMETER NoBrowser
  Ne pas ouvrir le navigateur automatiquement.

.PARAMETER Background
  Lance le serveur dans un processus PowerShell détaché (fenêtre cachée).

.PARAMETER Stop
  Arrête une instance lancée précédemment avec -Background.

.EXAMPLE
  .\PS-ZIM.ps1 .\wikipedia_fr_all_nopic_2024-06.zim

.EXAMPLE
  .\PS-ZIM.ps1 .\wikipedia.zim -Background -Port 8888

.EXAMPLE
  .\PS-ZIM.ps1 .\wikipedia.zim -Sequential -DebugLog

.EXAMPLE
  .\PS-ZIM.ps1 -Stop

.NOTES
  CHANGELOG
  ---------
  0.9.0  - Numérotation SemVer, auteur Eric Guiffault.
         - Journalisation passée en option (-DebugLog), désactivée par défaut.
         - Serveur HTTP concurrent via pool de runspaces (+ -Sequential / -MaxThreads).
         - Accès au flux ZIM rendu thread-safe (verrou réentrant).
         - ETag / 304 Not Modified sur les articles.
         - Vérification SHA-256 optionnelle de libzstd.dll.
         - Endpoint /api/suggest + autocomplétion dans la barre de navigation.
         - Source aléatoire thread-safe (Get-Random), bannière de démarrage alignée.
         - Détection de l'OS (Windows = cible testée ; autres = expérimental).
#>
[CmdletBinding(DefaultParameterSetName='Serve')]
param(
    [Parameter(ParameterSetName='Serve', Mandatory=$true, Position=0)]
    [Alias('Path','File')]
    [string]$ZimPath,

    [Parameter(ParameterSetName='Serve')]
    [int]$Port = 8642,

    [Parameter(ParameterSetName='Serve')]
    [string]$BindAddress = '127.0.0.1',

    [Parameter(ParameterSetName='Serve')]
    [int]$CacheSize = 64,

    [Parameter(ParameterSetName='Serve')]
    [int]$MaxThreads = 0,

    [Parameter(ParameterSetName='Serve')]
    [switch]$Sequential,

    [Parameter(ParameterSetName='Serve')]
    [switch]$DebugLog,

    [Parameter(ParameterSetName='Serve')]
    [switch]$NoBrowser,

    [Parameter(ParameterSetName='Serve')]
    [switch]$Background,

    [Parameter(ParameterSetName='Stop', Mandatory=$true)]
    [switch]$Stop
)

$ErrorActionPreference = 'Stop'
# UTF-8 pour la sortie console — silencieux si pas de console valide (ISE, redirection, etc.)
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
try { $OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

# ─────────────────────────────────────────────────────────────────────────────
#  Identité / version
# ─────────────────────────────────────────────────────────────────────────────
$Global:PsZimVersion = 'PS-ZIM v0.9.0'
$Global:PsZimAuthor  = 'Eric Guiffault'

# ─────────────────────────────────────────────────────────────────────────────
#  Détection de l'OS (Windows = cible testée)
# ─────────────────────────────────────────────────────────────────────────────
$Global:IsWindowsOS = $true
try {
    $Global:IsWindowsOS = [System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform(
        [System.Runtime.InteropServices.OSPlatform]::Windows)
} catch {
    $Global:IsWindowsOS = ($env:OS -eq 'Windows_NT')
}

# ─────────────────────────────────────────────────────────────────────────────
#  Chemins / fichiers d'état
# ─────────────────────────────────────────────────────────────────────────────
$ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$LibDir    = Join-Path $ScriptDir 'lib'
$PidFile   = Join-Path $ScriptDir 'PS-ZIM.pid'

function Write-PszBox {
    param([string[]]$Lines, [string]$Color = 'Cyan')
    $w = 0
    foreach ($l in $Lines) { if ($l.Length -gt $w) { $w = $l.Length } }
    Write-Host ('  ┌' + ('─' * ($w + 2)) + '┐') -ForegroundColor $Color
    foreach ($l in $Lines) {
        $pad = ' ' * ($w - $l.Length)
        Write-Host ('  │ ' + $l + $pad + ' │') -ForegroundColor $Color
    }
    Write-Host ('  └' + ('─' * ($w + 2)) + '┘') -ForegroundColor $Color
}

Write-Host ""
Write-PszBox @("$Global:PsZimVersion — par $Global:PsZimAuthor")
Write-Host ""

# ─────────────────────────────────────────────────────────────────────────────
#  -Stop : arrêter une instance en arrière-plan
# ─────────────────────────────────────────────────────────────────────────────
if ($Stop) {
    if (-not (Test-Path $PidFile)) {
        Write-Host "Aucune instance PS-ZIM enregistrée (pas de PS-ZIM.pid)." -ForegroundColor Yellow
        return
    }
    $pidVal = [int](Get-Content $PidFile -Raw).Trim()
    try {
        Stop-Process -Id $pidVal -Force -ErrorAction Stop
        Write-Host "PS-ZIM (PID $pidVal) arrêté." -ForegroundColor Green
    } catch {
        Write-Host "Impossible d'arrêter le PID $pidVal : $($_.Exception.Message)" -ForegroundColor Red
    }
    Remove-Item $PidFile -Force -ErrorAction SilentlyContinue
    return
}

# ─────────────────────────────────────────────────────────────────────────────
#  -Background : relance le script dans un processus détaché puis quitte
# ─────────────────────────────────────────────────────────────────────────────
if ($Background) {
    $scriptPath = $MyInvocation.MyCommand.Path
    $resolvedZim = (Resolve-Path $ZimPath).Path
    $argString = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -ZimPath `"$resolvedZim`" -Port $Port -BindAddress $BindAddress -CacheSize $CacheSize -MaxThreads $MaxThreads"
    if ($NoBrowser)  { $argString += ' -NoBrowser' }
    if ($Sequential) { $argString += ' -Sequential' }
    if ($DebugLog)   { $argString += ' -DebugLog' }

    $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $argString -WindowStyle Hidden -PassThru
    Start-Sleep -Milliseconds 400
    $proc.Id | Out-File -FilePath $PidFile -Encoding ascii -Force

    Write-Host ""
    Write-Host "  PS-ZIM démarré en arrière-plan" -ForegroundColor Green
    Write-Host "  ──────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  URL     : " -NoNewline; Write-Host "http://${BindAddress}:$Port/" -ForegroundColor Cyan
    Write-Host "  PID     : $($proc.Id)"
    Write-Host "  Arrêter : " -NoNewline; Write-Host ".\PS-ZIM.ps1 -Stop" -ForegroundColor Yellow
    Write-Host ""
    if (-not $NoBrowser) {
        Start-Sleep -Milliseconds 800
        Start-Process "http://${BindAddress}:$Port/" | Out-Null
    }
    return
}

# ─────────────────────────────────────────────────────────────────────────────
#  Logger de debug (optionnel — activé par -DebugLog)
# ─────────────────────────────────────────────────────────────────────────────
$Global:DebugEnabled = [bool]$DebugLog
$Global:DebugLogFile = Join-Path $ScriptDir 'ps-zim-debug.log'
if ($Global:DebugEnabled) {
    # Remise à zéro du log à chaque démarrage
    "" | Out-File -FilePath $Global:DebugLogFile -Encoding utf8 -Force
}

function Write-DebugLog {
    param([string]$msg)
    if (-not $Global:DebugEnabled) { return }
    $line = "{0} {1}" -f (Get-Date -Format 'HH:mm:ss.fff'), $msg
    # Append best-effort, tolérant à la concurrence
    for ($i = 0; $i -lt 5; $i++) {
        try { Add-Content -Path $Global:DebugLogFile -Value $line -Encoding utf8; break }
        catch { Start-Sleep -Milliseconds 5 }
    }
}

function Write-DebugError {
    param([string]$where, $errRec)
    $ex = $errRec.Exception
    $info = $errRec.InvocationInfo
    $lines = @(
        "═══════ ERREUR @ $where ═══════"
        "Type      : $($ex.GetType().FullName)"
        "Message   : $($ex.Message)"
        "Line#     : $($info.ScriptLineNumber)"
        "Position  : $($info.PositionMessage -replace '\r?\n', ' / ')"
        "ScriptName: $($info.ScriptName)"
        "Command   : $($info.MyCommand)"
        "Inner     : $(if ($ex.InnerException) { $ex.InnerException.Message } else { '(none)' })"
        "StackTrace:"
        ($errRec.ScriptStackTrace -split "`n" | ForEach-Object { "  $_" }) -join "`n"
        "════════════════════════════════"
    ) -join "`n"
    Write-DebugLog $lines
    Write-Host $lines -ForegroundColor Red
    return $lines
}

if ($Global:DebugEnabled) {
    Write-Host "  Log debug : $Global:DebugLogFile" -ForegroundColor DarkGray
}

# ─────────────────────────────────────────────────────────────────────────────
#  Vérification du fichier ZIM
# ─────────────────────────────────────────────────────────────────────────────
if (-not (Test-Path $ZimPath)) { throw "Fichier ZIM introuvable : $ZimPath" }
$ZimPath = (Resolve-Path $ZimPath).Path

# ─────────────────────────────────────────────────────────────────────────────
#  Bootstrap : libzstd.dll natif (binaire officiel Facebook/Meta)
#  → P/Invoke pur, zéro dépendance NuGet, marche sur PS 5.1 comme sur 7.x
#
#  Sécurité : vérification SHA-256 optionnelle. Renseignez les empreintes
#  officielles ci-dessous (calculées sur la DLL extraite) pour activer un
#  contrôle d'intégrité strict. Laissées vides => contrôle ignoré (avec note).
# ─────────────────────────────────────────────────────────────────────────────
$Global:ExpectedZstdSha256 = @{
    # 'win64' = 'À_RENSEIGNER_HASH_SHA256_MAJUSCULES'
    # 'win32' = 'À_RENSEIGNER_HASH_SHA256_MAJUSCULES'
}

if (-not (Test-Path $LibDir)) { New-Item -ItemType Directory -Path $LibDir -Force | Out-Null }

# Nettoyage d'un éventuel ZstdSharp.dll d'une tentative précédente
$stale = Join-Path $LibDir 'ZstdSharp.dll'
if (Test-Path $stale) { Remove-Item $stale -Force -ErrorAction SilentlyContinue }

function Test-ZstdHash {
    param([string]$Path, [string]$Arch)
    $expected = $Global:ExpectedZstdSha256[$Arch]
    if ([string]::IsNullOrWhiteSpace($expected) -or $expected -like 'À_RENSEIGNER*') {
        Write-DebugLog "Test-ZstdHash: pas d'empreinte épinglée pour $Arch — contrôle ignoré"
        return $true
    }
    $actual = (Get-FileHash -Path $Path -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actual -ne $expected.ToUpperInvariant()) {
        throw "Empreinte SHA-256 de libzstd.dll invalide (attendu $expected, obtenu $actual). Téléchargement potentiellement corrompu ou altéré."
    }
    Write-DebugLog "Test-ZstdHash: empreinte $Arch vérifiée OK"
    return $true
}

function Install-LibZstd {
    param(
        [string]$Version = '1.5.6'
    )
    $target = Join-Path $LibDir 'libzstd.dll'

    if (-not $Global:IsWindowsOS) {
        # Plateforme non-Windows : on n'auto-télécharge pas (les releases zstd
        # GitHub ne fournissent que des binaires Windows). On s'appuie sur la
        # libzstd du système. Support expérimental.
        Write-Host "  [!] OS non-Windows détecté : support expérimental." -ForegroundColor Yellow
        Write-Host "      Installez zstd via votre gestionnaire de paquets (ex. apt install zstd / brew install zstd)." -ForegroundColor DarkYellow
        return $null
    }

    $arch = if ([Environment]::Is64BitProcess) { 'win64' } else { 'win32' }

    if (Test-Path $target) {
        # Re-vérifie l'empreinte d'une DLL déjà en cache si une empreinte est épinglée
        Test-ZstdHash -Path $target -Arch $arch | Out-Null
        return $target
    }

    $url  = "https://github.com/facebook/zstd/releases/download/v$Version/zstd-v$Version-$arch.zip"

    Write-Host "  Téléchargement de libzstd $Version ($arch) depuis github.com/facebook/zstd..." -ForegroundColor DarkCyan

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

    $tmpZip = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), '.zip')
    try {
        Invoke-WebRequest -Uri $url -OutFile $tmpZip -UseBasicParsing
    } catch {
        throw "Échec du téléchargement de libzstd $Version : $($_.Exception.Message)"
    }

    $extractDir = Join-Path $env:TEMP ("psz_" + [Guid]::NewGuid().ToString('N'))
    [System.IO.Compression.ZipFile]::ExtractToDirectory($tmpZip, $extractDir)

    $candidate = Get-ChildItem -Path $extractDir -Recurse -Filter 'libzstd.dll' -ErrorAction SilentlyContinue |
                 Select-Object -First 1
    if (-not $candidate) {
        Remove-Item $tmpZip, $extractDir -Recurse -Force -ErrorAction SilentlyContinue
        throw "libzstd.dll introuvable dans l'archive téléchargée"
    }
    Copy-Item $candidate.FullName $target -Force
    Remove-Item $tmpZip, $extractDir -Recurse -Force -ErrorAction SilentlyContinue

    # Contrôle d'intégrité (si empreinte épinglée)
    Test-ZstdHash -Path $target -Arch $arch | Out-Null

    Write-Host "  → libzstd.dll installée dans .\lib\" -ForegroundColor DarkGray
    return $target
}

$ZstdDll = Install-LibZstd

# Charge libzstd.dll explicitement depuis ./lib/ avant tout P/Invoke (Windows).
if ($Global:IsWindowsOS) {
    $kernel32 = Add-Type -Name PsZimKernel32 -Namespace PsZim -PassThru -MemberDefinition @'
[DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
public static extern IntPtr LoadLibrary(string lpFileName);
[DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
public static extern bool SetDllDirectory(string lpPathName);
'@
    [void]$kernel32::SetDllDirectory($LibDir)
    $loaded = $kernel32::LoadLibrary($ZstdDll)
    if ($loaded -eq [IntPtr]::Zero) {
        $procArch = if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }
        throw "Échec du chargement de libzstd.dll (LoadLibrary). Vérifiez que la DLL correspond à l'architecture du processus PowerShell ($procArch)."
    }
}

# Wrapper P/Invoke pour la décompression Zstandard (API streaming)
Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace PsZim {
    public static class Zstd {
        [StructLayout(LayoutKind.Sequential)]
        public struct Buffer_ {
            public IntPtr ptr;
            public UIntPtr size;
            public UIntPtr pos;
        }

        [DllImport("libzstd", CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr ZSTD_createDStream();

        [DllImport("libzstd", CallingConvention = CallingConvention.Cdecl)]
        public static extern UIntPtr ZSTD_freeDStream(IntPtr zds);

        [DllImport("libzstd", CallingConvention = CallingConvention.Cdecl)]
        public static extern UIntPtr ZSTD_initDStream(IntPtr zds);

        [DllImport("libzstd", CallingConvention = CallingConvention.Cdecl)]
        public static extern UIntPtr ZSTD_decompressStream(IntPtr zds, ref Buffer_ output, ref Buffer_ input);

        [DllImport("libzstd", CallingConvention = CallingConvention.Cdecl)]
        public static extern uint ZSTD_isError(UIntPtr code);

        [DllImport("libzstd", CallingConvention = CallingConvention.Cdecl)]
        public static extern IntPtr ZSTD_getErrorName(UIntPtr code);

        [DllImport("libzstd", CallingConvention = CallingConvention.Cdecl)]
        public static extern UIntPtr ZSTD_DStreamOutSize();

        public static byte[] Decompress(byte[] compressed) {
            IntPtr zds = ZSTD_createDStream();
            if (zds == IntPtr.Zero) throw new InvalidOperationException("ZSTD_createDStream a renvoyé NULL");
            try {
                UIntPtr initRet = ZSTD_initDStream(zds);
                if (ZSTD_isError(initRet) != 0) {
                    throw new InvalidOperationException("ZSTD_initDStream: " + Marshal.PtrToStringAnsi(ZSTD_getErrorName(initRet)));
                }

                ulong outChunkU = ZSTD_DStreamOutSize().ToUInt64();
                int outChunk = (outChunkU == 0 || outChunkU > 0x100000) ? 131072 : (int)outChunkU;

                GCHandle inHandle  = GCHandle.Alloc(compressed, GCHandleType.Pinned);
                byte[] outBuf      = new byte[outChunk];
                GCHandle outHandle = GCHandle.Alloc(outBuf, GCHandleType.Pinned);

                try {
                    Buffer_ inBuf = new Buffer_ {
                        ptr  = inHandle.AddrOfPinnedObject(),
                        size = (UIntPtr)compressed.Length,
                        pos  = UIntPtr.Zero
                    };

                    using (MemoryStream ms = new MemoryStream(compressed.Length * 3)) {
                        while (true) {
                            Buffer_ outBufStruct = new Buffer_ {
                                ptr  = outHandle.AddrOfPinnedObject(),
                                size = (UIntPtr)outChunk,
                                pos  = UIntPtr.Zero
                            };

                            UIntPtr ret = ZSTD_decompressStream(zds, ref outBufStruct, ref inBuf);
                            if (ZSTD_isError(ret) != 0) {
                                throw new InvalidOperationException("ZSTD_decompressStream: " + Marshal.PtrToStringAnsi(ZSTD_getErrorName(ret)));
                            }

                            int produced = (int)outBufStruct.pos.ToUInt64();
                            if (produced > 0) ms.Write(outBuf, 0, produced);

                            if (ret.ToUInt64() == 0) break;  // frame terminé

                            if (produced == 0 && inBuf.pos.ToUInt64() >= inBuf.size.ToUInt64()) {
                                throw new InvalidOperationException("ZSTD: fin de stream prématurée");
                            }
                        }
                        return ms.ToArray();
                    }
                } finally {
                    inHandle.Free();
                    outHandle.Free();
                }
            } finally {
                ZSTD_freeDStream(zds);
            }
        }
    }
}
'@ -ErrorAction Stop

# ─────────────────────────────────────────────────────────────────────────────
#  Lecteur ZIM (à base de hashtables — évite les erreurs de redéfinition de classe)
# ─────────────────────────────────────────────────────────────────────────────
$ZIM_MAGIC = 0x44D495A   # bytes 5A 49 4D 04 = "ZIM\x04" en little-endian

function New-ZimReader {
    param([string]$Path, [int]$CacheSize)

    $fs = [System.IO.File]::Open($Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::Read)
    $br = New-Object System.IO.BinaryReader($fs)

    $zim = [ordered]@{
        Path         = $Path
        Stream       = $fs
        Reader       = $br
        Header       = $null
        MimeList     = @()
        ClusterCache = @{}
        CacheOrder   = New-Object System.Collections.Generic.Queue[int]
        MaxCacheSize = $CacheSize
        Lock         = New-Object Object
    }

    # ── Header (80 octets, little-endian) ──
    $fs.Position = 0
    $magic = $br.ReadUInt32()
    if ($magic -ne $ZIM_MAGIC) {
        throw ("Fichier ZIM invalide (magic number 0x{0:X8})." -f $magic)
    }
    $zim.Header = [ordered]@{
        MajorVersion  = $br.ReadUInt16()
        MinorVersion  = $br.ReadUInt16()
        Uuid          = $br.ReadBytes(16)
        ArticleCount  = $br.ReadUInt32()
        ClusterCount  = $br.ReadUInt32()
        UrlPtrPos     = $br.ReadUInt64()
        TitlePtrPos   = $br.ReadUInt64()
        ClusterPtrPos = $br.ReadUInt64()
        MimeListPos   = $br.ReadUInt64()
        MainPage      = $br.ReadUInt32()
        LayoutPage    = $br.ReadUInt32()
        ChecksumPos   = $br.ReadUInt64()
    }

    # ── Liste des MIME types (chaînes UTF-8 nulles-terminées, liste finie par "") ──
    $fs.Position = $zim.Header.MimeListPos
    $mimes = New-Object System.Collections.Generic.List[string]
    while ($true) {
        $bytes = New-Object System.Collections.Generic.List[byte]
        while ($true) {
            $b = $br.ReadByte()
            if ($b -eq 0) { break }
            $bytes.Add($b)
        }
        if ($bytes.Count -eq 0) { break }
        $mimes.Add([System.Text.Encoding]::UTF8.GetString($bytes.ToArray()))
    }
    $zim.MimeList = $mimes.ToArray()

    return $zim
}

# Vérifie la santé de la Title Pointer List. Pour les ZIMs v6.x récents
# (depuis ~2023), titlePtrPos n'est plus garanti — la liste a été déplacée
# dans une entrée X/listing/titleOrdered. On teste 5 indices et on regarde
# si les UInt32 lus sont < ArticleCount. Si non, on bascule en mode URL-binsearch.
function Test-TitlePtrListHealth {
    param($zim)
    Write-DebugLog "Test-TitlePtrListHealth: ENTRY"
    $count = [System.Convert]::ToInt64($zim.Header.ArticleCount)
    Write-DebugLog "Test-TitlePtrListHealth: count=$count (type=$($count.GetType().FullName))"
    if ($count -le 0) { return $false }

    $c = [int64]$count
    $samples = New-Object System.Collections.Generic.List[int]
    [void]$samples.Add(0)
    [void]$samples.Add([int]($c / 4L))
    [void]$samples.Add([int]($c / 2L))
    [void]$samples.Add([int](($c * 3L) / 4L))
    [void]$samples.Add([int]($c - 1L))
    Write-DebugLog ("Test-TitlePtrListHealth: samples=" + ($samples -join ','))

    $bad = 0
    foreach ($i in $samples) {
        try {
            $idx = Get-ZimTitlePtr $zim ([uint32]$i)
            $idxL = [System.Convert]::ToInt64($idx)
            if ($idxL -ge $c -or $idxL -lt 0) { $bad++ }
        } catch { $bad++ }
    }
    Write-DebugLog "Test-TitlePtrListHealth: probe $bad/5 invalid"
    return ($bad -le 1)  # tolérance 1 sur 5
}

function Read-ZimCString {
    param($Reader)
    $bytes = New-Object System.Collections.Generic.List[byte]
    while ($true) {
        $b = $Reader.ReadByte()
        if ($b -eq 0) { break }
        $bytes.Add($b)
    }
    return [System.Text.Encoding]::UTF8.GetString($bytes.ToArray())
}

# ── Accès au flux : tous protégés par $zim.Lock (verrou réentrant Monitor) ──
function Get-ZimUrlPtr {
    param($zim, [uint32]$Index)
    [System.Threading.Monitor]::Enter($zim.Lock)
    try {
        $zim.Stream.Position = [int64]$zim.Header.UrlPtrPos + ([int64]$Index * 8L)
        return $zim.Reader.ReadUInt64()
    } finally { [System.Threading.Monitor]::Exit($zim.Lock) }
}

function Get-ZimTitlePtr {
    param($zim, [uint32]$Index)
    [System.Threading.Monitor]::Enter($zim.Lock)
    try {
        $zim.Stream.Position = [int64]$zim.Header.TitlePtrPos + ([int64]$Index * 4L)
        return $zim.Reader.ReadUInt32()
    } finally { [System.Threading.Monitor]::Exit($zim.Lock) }
}

function Get-ZimClusterPtr {
    param($zim, [uint32]$Index)
    [System.Threading.Monitor]::Enter($zim.Lock)
    try {
        $zim.Stream.Position = [int64]$zim.Header.ClusterPtrPos + ([int64]$Index * 8L)
        return $zim.Reader.ReadUInt64()
    } finally { [System.Threading.Monitor]::Exit($zim.Lock) }
}

function Read-ZimDirEntry {
    param($zim, [uint32]$Index)
    # Verrou tenu sur toute la lecture (séquence de reads dépendants du Position).
    # Monitor est réentrant : l'appel imbriqué à Get-ZimUrlPtr ne relâche pas le verrou.
    [System.Threading.Monitor]::Enter($zim.Lock)
    try {
        $pos = Get-ZimUrlPtr $zim $Index
        $streamPos = 0L
        try {
            $streamPos = [int64]$pos
        } catch {
            throw "ZIM-ENTRY-INVALID: index $Index → position aberrante $pos"
        }
        if ($streamPos -lt 0 -or $streamPos -ge $zim.Stream.Length) {
            throw "ZIM-ENTRY-INVALID: index $Index → position $streamPos hors fichier"
        }
        $zim.Stream.Position = $streamPos
        $br = $zim.Reader

        $mime = $br.ReadUInt16()
        $entry = [ordered]@{
            Index         = $Index
            Mime          = $mime
            IsRedirect    = ($mime -eq 0xFFFF)
            IsLinkTarget  = ($mime -eq 0xFFFE)
            IsDeleted     = ($mime -eq 0xFFFD)
            ParamLen      = 0
            Namespace     = ''
            Revision      = 0
            ClusterNumber = 0
            BlobNumber    = 0
            RedirectIndex = 0
            Url           = ''
            Title         = ''
        }
        $entry.ParamLen = $br.ReadByte()
        $entry.Namespace = [char]$br.ReadByte()
        $entry.Revision = $br.ReadUInt32()
        if ($entry.IsRedirect) {
            $entry.RedirectIndex = $br.ReadUInt32()
        } elseif (-not $entry.IsLinkTarget -and -not $entry.IsDeleted) {
            $entry.ClusterNumber = $br.ReadUInt32()
            $entry.BlobNumber = $br.ReadUInt32()
        }
        $entry.Url   = Read-ZimCString $br
        $entry.Title = Read-ZimCString $br
        if ([string]::IsNullOrEmpty($entry.Title)) { $entry.Title = $entry.Url }
        return $entry
    } finally { [System.Threading.Monitor]::Exit($zim.Lock) }
}

function Resolve-ZimEntry {
    param($zim, [uint32]$Index)
    $entry = Read-ZimDirEntry $zim $Index
    $hops = 0
    while ($entry.IsRedirect -and $hops -lt 16) {
        $entry = Read-ZimDirEntry $zim $entry.RedirectIndex
        $hops++
    }
    return $entry
}

# Comparateur (namespace, url) — ordinal sur octets UTF-8 comme dans le format ZIM
function Compare-ZimKey {
    param([string]$Ns1, [string]$Url1, [string]$Ns2, [string]$Url2)
    $c = [int][char]$Ns1 - [int][char]$Ns2
    if ($c -ne 0) { return $c }
    return [string]::CompareOrdinal($Url1, $Url2)
}

function Find-ZimByUrl {
    param($zim, [string]$Namespace, [string]$Url)
    $count = [int]$zim.Header.ArticleCount
    if ($count -eq 0) { return -1 }
    $lo = 0
    $hi = $count - 1
    while ($lo -le $hi) {
        $mid = [int](($lo + $hi) -shr 1)
        try {
            $entry = Read-ZimDirEntry $zim ([uint32]$mid)
        } catch {
            $hi = $mid - 1
            continue
        }
        $cmp = Compare-ZimKey $entry.Namespace $entry.Url $Namespace $Url
        if ($cmp -eq 0) { return $mid }
        if ($cmp -lt 0) { $lo = $mid + 1 } else { $hi = $mid - 1 }
    }
    return -1
}

# Résolution permissive d'un chemin HTTP en entrée ZIM.
function Find-ZimEntryByPath {
    param($zim, [string]$Path)
    if ([string]::IsNullOrEmpty($Path)) { return -1 }

    if ($zim.Header.MajorVersion -ge 6) {
        $idx = Find-ZimByUrl $zim 'C' $Path
        if ($idx -ge 0) { return $idx }
    }
    if ($Path -match '^([A-Za-z\-])/(.+)$') {
        $idx = Find-ZimByUrl $zim $matches[1] $matches[2]
        if ($idx -ge 0) { return $idx }
    }
    $idx = Find-ZimByUrl $zim 'A' $Path
    if ($idx -ge 0) { return $idx }
    $idx = Find-ZimByUrl $zim 'C' $Path
    if ($idx -ge 0) { return $idx }
    return -1
}

# ─────────────────────────────────────────────────────────────────────────────
#  Lecture / décompression de clusters (avec cache LRU)
# ─────────────────────────────────────────────────────────────────────────────
function Get-ZimCluster {
    param($zim, [uint32]$ClusterNum)

    [System.Threading.Monitor]::Enter($zim.Lock)
    try {
        if ($zim.ClusterCache.ContainsKey([int]$ClusterNum)) {
            return $zim.ClusterCache[[int]$ClusterNum]
        }
    } finally { [System.Threading.Monitor]::Exit($zim.Lock) }

    # Lecture brute des octets compressés sous verrou (accès Position dépendant)
    [System.Threading.Monitor]::Enter($zim.Lock)
    try {
        $pos     = Get-ZimClusterPtr $zim $ClusterNum
        $nextPos = if ($ClusterNum + 1 -lt $zim.Header.ClusterCount) {
            Get-ZimClusterPtr $zim ($ClusterNum + 1)
        } else {
            $zim.Header.ChecksumPos
        }
        $clusterSize = [int64]($nextPos - $pos)

        $zim.Stream.Position = [int64]$pos
        $info = $zim.Reader.ReadByte()
        $compression = $info -band 0x0F
        $extended    = ($info -band 0x10) -ne 0
        $remaining   = $clusterSize - 1
        $compressedData = $zim.Reader.ReadBytes([int]$remaining)
    } finally { [System.Threading.Monitor]::Exit($zim.Lock) }

    # Décompression hors verrou (étape coûteuse → parallélisable)
    $raw = $null
    switch ($compression) {
        1 { $raw = $compressedData }                  # pas de compression
        2 { throw "Cluster zlib non supporté (très ancien ZIM)." }
        3 { throw "Cluster bzip2 non supporté (très ancien ZIM)." }
        4 {
            throw "Cluster LZMA2/XZ détecté : ce ZIM est antérieur à 2020 et utilise un format non supporté par PS-ZIM. Re-téléchargez la version moderne sur https://library.kiwix.org."
        }
        5 {
            $raw = [PsZim.Zstd]::Decompress($compressedData)
        }
        default { throw "Type de compression inconnu : $compression" }
    }

    # Lecture des offsets internes du cluster
    $ms = New-Object System.IO.MemoryStream(,$raw)
    $br = New-Object System.IO.BinaryReader($ms)
    if ($extended) {
        $first = $br.ReadUInt64()
        $nOff  = [int]($first / 8)
        $offsets = New-Object uint64[] $nOff
        $offsets[0] = $first
        for ($i = 1; $i -lt $nOff; $i++) { $offsets[$i] = $br.ReadUInt64() }
    } else {
        $first = $br.ReadUInt32()
        $nOff  = [int]($first / 4)
        $offsets = New-Object uint32[] $nOff
        $offsets[0] = $first
        for ($i = 1; $i -lt $nOff; $i++) { $offsets[$i] = $br.ReadUInt32() }
    }
    $br.Dispose(); $ms.Dispose()

    $cluster = [ordered]@{ Offsets = $offsets; Data = $raw; Extended = $extended }

    [System.Threading.Monitor]::Enter($zim.Lock)
    try {
        if (-not $zim.ClusterCache.ContainsKey([int]$ClusterNum)) {
            $zim.ClusterCache[[int]$ClusterNum] = $cluster
            $zim.CacheOrder.Enqueue([int]$ClusterNum)
            while ($zim.ClusterCache.Count -gt $zim.MaxCacheSize) {
                $old = $zim.CacheOrder.Dequeue()
                if ($zim.ClusterCache.ContainsKey($old)) { $zim.ClusterCache.Remove($old) }
            }
        }
    } finally { [System.Threading.Monitor]::Exit($zim.Lock) }

    return $cluster
}

function Get-ZimBlob {
    param($zim, [uint32]$ClusterNum, [uint32]$BlobNum)
    $cluster = Get-ZimCluster $zim $ClusterNum
    $offsets = $cluster.Offsets
    if ($BlobNum -ge ($offsets.Length - 1)) { return ([byte[]]@()) }
    $start = [int64]$offsets[$BlobNum]
    $end   = [int64]$offsets[$BlobNum + 1]
    $len   = [int]($end - $start)
    $blob  = New-Object byte[] $len
    [System.Array]::Copy($cluster.Data, $start, $blob, 0, $len)
    return $blob
}

# ─────────────────────────────────────────────────────────────────────────────
#  Recherche par préfixe de titre
# ─────────────────────────────────────────────────────────────────────────────
function Search-ZimTitles {
    param($zim, [string]$Prefix, [int]$Limit = 50)

    Write-DebugLog "Search-ZimTitles: ENTRY Prefix='$Prefix' Limit=$Limit TitlePtrUsable=$($zim.TitlePtrUsable)"

    if (-not $zim.TitlePtrUsable) {
        return Search-ZimByUrl $zim $Prefix $Limit
    }

    $count = [int]$zim.Header.ArticleCount
    Write-DebugLog "Search-ZimTitles: count=$count"
    if ($count -eq 0 -or [string]::IsNullOrEmpty($Prefix)) {
        Write-DebugLog "Search-ZimTitles: early exit (count=0 or empty prefix)"
        return @()
    }

    $searchKey = if ($Prefix.Length -gt 0) {
        ([char]::ToUpperInvariant($Prefix[0])).ToString() + $Prefix.Substring(1)
    } else { $Prefix }
    Write-DebugLog "Search-ZimTitles: searchKey='$searchKey'"

    $lo = 0; $hi = $count - 1
    $iter = 0
    while ($lo -lt $hi) {
        $iter++
        $mid = [int](($lo + $hi) -shr 1)
        try {
            $urlIdx = Get-ZimTitlePtr $zim ([uint32]$mid)
            $entry  = Read-ZimDirEntry $zim ([uint32]$urlIdx)
        } catch {
            Write-DebugLog "Search-ZimTitles: BSEARCH iter=$iter mid=$mid CAUGHT: $($_.Exception.Message)"
            $hi = $mid - 1
            if ($hi -lt $lo) { break }
            continue
        }
        if ([string]::CompareOrdinal($entry.Title, $searchKey) -lt 0) {
            $lo = $mid + 1
        } else {
            $hi = $mid
        }
    }
    Write-DebugLog "Search-ZimTitles: BSEARCH done lo=$lo iter=$iter"

    $prefixLow = $Prefix.ToLowerInvariant()
    $results = New-Object System.Collections.Generic.List[object]
    $consecutiveMisses = 0
    $maxMisses = 200
    $scanned = 0
    $errors = 0
    for ($i = $lo; $i -lt $count -and $results.Count -lt $Limit; $i++) {
        $scanned++
        try {
            $urlIdx = Get-ZimTitlePtr $zim ([uint32]$i)
            $entry  = Read-ZimDirEntry $zim ([uint32]$urlIdx)
        } catch {
            $errors++
            if ($errors -le 3) {
                Write-DebugLog "Search-ZimTitles: SCAN i=$i CAUGHT: $($_.Exception.Message)"
            }
            continue
        }
        if ($entry.IsLinkTarget -or $entry.IsDeleted) { continue }
        if ($entry.Namespace -notin 'A','C') { continue }

        $titleLow = $entry.Title.ToLowerInvariant()
        if ($titleLow.StartsWith($prefixLow)) {
            [void]$results.Add($entry)
            $consecutiveMisses = 0
        } else {
            $consecutiveMisses++
            if ($consecutiveMisses -gt $maxMisses) { break }
        }
    }
    Write-DebugLog "Search-ZimTitles: EXIT scanned=$scanned errors=$errors results=$($results.Count)"
    return $results.ToArray()
}

# ─────────────────────────────────────────────────────────────────────────────
#  Recherche par URL — fallback quand la title pointer list est cassée.
# ─────────────────────────────────────────────────────────────────────────────
function Search-ZimByUrl {
    param($zim, [string]$Prefix, [int]$Limit = 50)

    Write-DebugLog "Search-ZimByUrl: ENTRY Prefix='$Prefix' Limit=$Limit"
    $count = [int]$zim.Header.ArticleCount
    if ($count -eq 0 -or [string]::IsNullOrEmpty($Prefix)) { return @() }

    $key = $Prefix -replace ' ', '_'
    if ($key.Length -gt 0) {
        $key = ([char]::ToUpperInvariant($key[0])).ToString() + $key.Substring(1)
    }

    $ns = if ($zim.Header.MajorVersion -ge 6) { 'C' } else { 'A' }
    Write-DebugLog "Search-ZimByUrl: searchKey='$key' ns='$ns'"

    $lo = 0; $hi = $count - 1
    $iter = 0
    while ($lo -lt $hi) {
        $iter++
        $mid = [int](($lo + $hi) -shr 1)
        try {
            $entry = Read-ZimDirEntry $zim ([uint32]$mid)
        } catch {
            Write-DebugLog "Search-ZimByUrl: BSEARCH iter=$iter mid=$mid CAUGHT: $($_.Exception.Message)"
            $hi = $mid - 1
            if ($hi -lt $lo) { break }
            continue
        }
        $cmp = Compare-ZimKey $entry.Namespace $entry.Url $ns $key
        if ($cmp -lt 0) { $lo = $mid + 1 } else { $hi = $mid }
    }
    Write-DebugLog "Search-ZimByUrl: BSEARCH done lo=$lo iter=$iter"

    $prefixLow = $Prefix.ToLowerInvariant() -replace ' ', '_'
    $results = New-Object System.Collections.Generic.List[object]
    $consecutiveMisses = 0
    $maxMisses = 500
    $scanned = 0
    $errors = 0
    for ($i = $lo; $i -lt $count -and $results.Count -lt $Limit; $i++) {
        $scanned++
        try {
            $entry = Read-ZimDirEntry $zim ([uint32]$i)
        } catch {
            $errors++
            if ($errors -le 3) {
                Write-DebugLog "Search-ZimByUrl: SCAN i=$i CAUGHT: $($_.Exception.Message)"
            }
            continue
        }
        if ($entry.IsLinkTarget -or $entry.IsDeleted) { continue }
        if ($entry.Namespace -notin 'A','C') { continue }

        $urlLow = $entry.Url.ToLowerInvariant()
        $matchUrl   = $urlLow.StartsWith($prefixLow)
        $matchTitle = $entry.Title.ToLowerInvariant().Contains($Prefix.ToLowerInvariant())
        if ($matchUrl -or $matchTitle) {
            [void]$results.Add($entry)
            $consecutiveMisses = 0
        } else {
            $consecutiveMisses++
            if ($consecutiveMisses -gt $maxMisses) { break }
        }
    }
    Write-DebugLog "Search-ZimByUrl: EXIT scanned=$scanned errors=$errors results=$($results.Count)"
    return $results.ToArray()
}

# ─────────────────────────────────────────────────────────────────────────────
#  Helpers HTTP
# ─────────────────────────────────────────────────────────────────────────────
function Test-IfNoneMatch {
    param($Req, [string]$ETag)
    if ([string]::IsNullOrEmpty($ETag)) { return $false }
    $inm = $Req.Headers['If-None-Match']
    if ([string]::IsNullOrEmpty($inm)) { return $false }
    $tokens = $inm -split ',' | ForEach-Object { $_.Trim() }
    return ($tokens -contains $ETag) -or ($tokens -contains '*')
}

function Send-HttpRaw {
    param($Res, [byte[]]$Body, [string]$ContentType, [int]$Status = 200, [string]$ETag = $null)
    try {
        $Res.StatusCode = $Status
        $Res.ContentType = $ContentType
        $Res.ContentLength64 = $Body.Length
        $Res.AddHeader('Cache-Control','public, max-age=3600')
        $Res.AddHeader('X-Server','PS-ZIM')
        if (-not [string]::IsNullOrEmpty($ETag)) { $Res.AddHeader('ETag', $ETag) }
        $Res.OutputStream.Write($Body, 0, $Body.Length)
    } catch { }
    finally { try { $Res.OutputStream.Close() } catch { } }
}

function Send-HttpNotModified {
    param($Res, [string]$ETag)
    try {
        $Res.StatusCode = 304
        $Res.ContentLength64 = 0
        if (-not [string]::IsNullOrEmpty($ETag)) { $Res.AddHeader('ETag', $ETag) }
        $Res.AddHeader('Cache-Control','public, max-age=3600')
        $Res.OutputStream.Close()
    } catch { }
}

function Send-HttpHtml {
    param($Res, [string]$Html, [int]$Status = 200)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Html)
    Send-HttpRaw -Res $Res -Body $bytes -ContentType 'text/html; charset=utf-8' -Status $Status
}

function Send-HttpText {
    param($Res, [string]$Text, [int]$Status = 200)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    Send-HttpRaw -Res $Res -Body $bytes -ContentType 'text/plain; charset=utf-8' -Status $Status
}

function Send-HttpRedirect {
    param($Res, [string]$Location, [int]$Status = 302)
    try {
        $Res.StatusCode = $Status
        $Res.RedirectLocation = $Location
        $Res.ContentLength64 = 0
        $Res.OutputStream.Close()
    } catch { }
}

function ConvertTo-SafeHtml {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($s)
}

function ConvertTo-JsonString {
    param([string]$s)
    if ($null -eq $s) { return '""' }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('"')
    foreach ($ch in $s.ToCharArray()) {
        switch ($ch) {
            '"'  { [void]$sb.Append('\"') }
            '\'  { [void]$sb.Append('\\') }
            "`b" { [void]$sb.Append('\b') }
            "`f" { [void]$sb.Append('\f') }
            "`n" { [void]$sb.Append('\n') }
            "`r" { [void]$sb.Append('\r') }
            "`t" { [void]$sb.Append('\t') }
            default {
                if ([int]$ch -lt 32) { [void]$sb.Append(('\u{0:x4}' -f [int]$ch)) }
                else { [void]$sb.Append($ch) }
            }
        }
    }
    [void]$sb.Append('"')
    return $sb.ToString()
}

function Get-EntryHttpUrl {
    param($zim, $entry)
    if ($zim.Header.MajorVersion -ge 6) {
        return '/' + $entry.Url
    } else {
        return '/' + $entry.Namespace + '/' + $entry.Url
    }
}

function Get-MimeType {
    param($zim, [int]$mimeIdx)
    if ($mimeIdx -ge 0 -and $mimeIdx -lt $zim.MimeList.Count) {
        return $zim.MimeList[$mimeIdx]
    }
    return 'application/octet-stream'
}

# ─────────────────────────────────────────────────────────────────────────────
#  Top-bar injectée dans tous les articles HTML servis (+ autocomplétion)
# ─────────────────────────────────────────────────────────────────────────────
$Global:TopBarHtml = @'
<style id="psz-style">
  #psz-bar, #psz-bar * { box-sizing: border-box; font-family: ui-serif, "Iowan Old Style", "Apple Garamond", "Palatino Linotype", "Source Serif Pro", Georgia, serif; }
  #psz-bar {
    position: fixed; top: 0; left: 0; right: 0; height: 44px;
    display: flex; align-items: center; gap: 10px;
    padding: 0 18px;
    background: #14110d;
    color: #f0e6d2;
    border-bottom: 1px solid #3a2f1f;
    z-index: 2147483647;
    box-shadow: 0 2px 10px rgba(0,0,0,.18);
    font-size: 14px;
  }
  #psz-bar .psz-brand {
    font-family: ui-serif, Georgia, serif;
    font-weight: 600; font-size: 15px;
    letter-spacing: .04em;
    color: #d9b370;
    margin-right: 4px;
  }
  #psz-bar .psz-brand small { font-weight: 400; font-size: 11px; color: #8a7a5a; letter-spacing: .12em; margin-left: 6px; text-transform: uppercase; }
  #psz-bar a.psz-link {
    color: #f0e6d2; text-decoration: none; padding: 6px 10px;
    border-radius: 3px; line-height: 1; font-size: 15px;
    transition: background .15s;
  }
  #psz-bar a.psz-link:hover { background: #2a221a; color: #f7d98a; }
  #psz-bar form.psz-search { flex: 1; display: flex; max-width: 520px; position: relative; }
  #psz-bar input.psz-q {
    flex: 1; background: #1f1a13; color: #f0e6d2;
    border: 1px solid #3a2f1f; border-radius: 3px;
    padding: 6px 12px; font-size: 13px;
    font-family: ui-monospace, "SF Mono", "JetBrains Mono", Consolas, monospace;
    outline: none; transition: border-color .15s;
  }
  #psz-bar input.psz-q:focus { border-color: #d9b370; }
  #psz-bar input.psz-q::placeholder { color: #6b5d44; }
  #psz-suggest {
    position: absolute; top: 38px; left: 0; right: 0;
    background: #1f1a13; border: 1px solid #3a2f1f; border-top: none;
    border-radius: 0 0 4px 4px; max-height: 360px; overflow-y: auto;
    display: none; z-index: 2147483647;
  }
  #psz-suggest a {
    display: block; padding: 8px 12px; color: #f0e6d2; text-decoration: none;
    font-size: 13px; border-bottom: 1px solid #2a221a;
  }
  #psz-suggest a:last-child { border-bottom: none; }
  #psz-suggest a:hover, #psz-suggest a.psz-active { background: #2a221a; color: #f7d98a; }
  html { scroll-padding-top: 52px; }
  body { padding-top: 48px !important; }
</style>
<div id="psz-bar">
  <a class="psz-link" href="/" title="Accueil">&#9737;</a>
  <span class="psz-brand">PS-ZIM<small>archive</small></span>
  <form class="psz-search" action="/search" method="get" role="search" autocomplete="off">
    <input class="psz-q" id="psz-q" name="q" placeholder="Rechercher un article…" autocomplete="off" spellcheck="false">
    <div id="psz-suggest"></div>
  </form>
  <a class="psz-link" href="/random" title="Article aléatoire">&#9879;</a>
</div>
<script>
(function(){
  var inp = document.getElementById('psz-q');
  var box = document.getElementById('psz-suggest');
  if(!inp || !box) return;
  var t = null, items = [], active = -1;
  function hide(){ box.style.display='none'; active=-1; }
  function render(){
    if(!items.length){ hide(); return; }
    box.innerHTML = items.map(function(it,i){
      var cls = (i===active) ? ' class="psz-active"' : '';
      return '<a'+cls+' href="'+it.url+'">'+it.title.replace(/</g,'&lt;')+'</a>';
    }).join('');
    box.style.display='block';
  }
  inp.addEventListener('input', function(){
    var q = inp.value.trim();
    if(t) clearTimeout(t);
    if(q.length < 2){ hide(); return; }
    t = setTimeout(function(){
      fetch('/api/suggest?q='+encodeURIComponent(q))
        .then(function(r){ return r.json(); })
        .then(function(d){ items = d || []; active=-1; render(); })
        .catch(function(){ hide(); });
    }, 160);
  });
  inp.addEventListener('keydown', function(e){
    if(box.style.display!=='block') return;
    if(e.key==='ArrowDown'){ active=Math.min(active+1, items.length-1); render(); e.preventDefault(); }
    else if(e.key==='ArrowUp'){ active=Math.max(active-1, 0); render(); e.preventDefault(); }
    else if(e.key==='Enter' && active>=0){ window.location.href = items[active].url; e.preventDefault(); }
    else if(e.key==='Escape'){ hide(); }
  });
  document.addEventListener('click', function(e){ if(!box.contains(e.target) && e.target!==inp) hide(); });
})();
</script>
'@

function Add-TopBar {
    param([string]$Html)
    if ([string]::IsNullOrEmpty($Html)) { return $Global:TopBarHtml }
    $rx = [regex]'(?i)<body\b[^>]*>'
    $m  = $rx.Match($Html)
    if ($m.Success) {
        return $Html.Substring(0, $m.Index + $m.Length) + $Global:TopBarHtml + $Html.Substring($m.Index + $m.Length)
    }
    return $Global:TopBarHtml + $Html
}

# ─────────────────────────────────────────────────────────────────────────────
#  Pages générées : accueil, résultats de recherche, 404
# ─────────────────────────────────────────────────────────────────────────────
$Global:PageBaseCss = @'
<style>
  :root {
    --psz-bg: #f4ecd8;
    --psz-bg-sub: #ede2c5;
    --psz-ink: #2a2118;
    --psz-ink-mute: #6b5d44;
    --psz-accent: #8b1c1c;
    --psz-rule: #c9b582;
    --psz-paper: #faf3e0;
  }
  html { background: var(--psz-bg); }
  body {
    background: var(--psz-bg);
    color: var(--psz-ink);
    font-family: ui-serif, "Iowan Old Style", "Apple Garamond", "Palatino Linotype", "Source Serif Pro", Georgia, serif;
    max-width: 720px;
    margin: 0 auto;
    padding: 70px 32px 80px;
    line-height: 1.55;
    background-image:
      radial-gradient(circle at 20% 10%, rgba(139, 28, 28, .04), transparent 40%),
      radial-gradient(circle at 80% 90%, rgba(139, 28, 28, .03), transparent 40%);
  }
  .psz-mast {
    text-align: center;
    border-bottom: 1px solid var(--psz-rule);
    padding-bottom: 24px;
    margin-bottom: 36px;
    position: relative;
  }
  .psz-mast .eyebrow {
    font-family: ui-monospace, "SF Mono", Consolas, monospace;
    font-size: 11px; letter-spacing: .3em; text-transform: uppercase;
    color: var(--psz-ink-mute);
    margin-bottom: 16px;
  }
  .psz-mast h1 {
    font-family: ui-serif, "Apple Garamond", Georgia, serif;
    font-weight: 400; font-size: 56px; letter-spacing: -.01em;
    margin: 0 0 6px; color: var(--psz-ink);
    font-style: italic;
  }
  .psz-mast h1 .amp { color: var(--psz-accent); font-style: normal; }
  .psz-mast .filename {
    font-family: ui-monospace, "SF Mono", Consolas, monospace;
    font-size: 12px; color: var(--psz-ink-mute);
    word-break: break-all;
  }
  .psz-search-box {
    display: flex; gap: 0; margin: 30px 0 40px;
    border: 1px solid var(--psz-rule);
    background: var(--psz-paper);
    border-radius: 2px;
    box-shadow: 0 1px 0 rgba(0,0,0,.04), inset 0 1px 0 rgba(255,255,255,.4);
  }
  .psz-search-box input {
    flex: 1; padding: 16px 20px;
    font-size: 17px; font-family: inherit;
    border: none; background: transparent;
    color: var(--psz-ink); outline: none;
  }
  .psz-search-box input::placeholder { color: var(--psz-ink-mute); font-style: italic; }
  .psz-search-box button {
    padding: 16px 24px; font-size: 14px;
    font-family: ui-monospace, Consolas, monospace;
    letter-spacing: .2em; text-transform: uppercase;
    background: var(--psz-ink); color: var(--psz-bg);
    border: none; cursor: pointer;
    transition: background .15s;
  }
  .psz-search-box button:hover { background: var(--psz-accent); }
  .psz-meta {
    display: grid; grid-template-columns: 1fr 1fr;
    gap: 4px 24px; margin-top: 40px;
    border-top: 1px solid var(--psz-rule); padding-top: 24px;
    font-size: 13px;
  }
  .psz-meta dt { color: var(--psz-ink-mute); font-family: ui-monospace, Consolas, monospace;
                  font-size: 11px; letter-spacing: .15em; text-transform: uppercase; }
  .psz-meta dd { margin: 0 0 12px; font-family: ui-serif, Georgia, serif; }
  .psz-meta dd code { font-family: ui-monospace, Consolas, monospace; font-size: 12px; color: var(--psz-accent); }
  .psz-actions { display: flex; gap: 14px; justify-content: center; margin-top: 16px; }
  .psz-actions a {
    color: var(--psz-ink-mute); text-decoration: none;
    font-family: ui-monospace, Consolas, monospace;
    font-size: 11px; letter-spacing: .2em; text-transform: uppercase;
    border-bottom: 1px solid transparent; padding-bottom: 2px;
    transition: color .15s, border-color .15s;
  }
  .psz-actions a:hover { color: var(--psz-accent); border-color: var(--psz-accent); }
  .psz-msg { background: #fff8e2; border-left: 3px solid var(--psz-accent);
              padding: 12px 16px; margin-bottom: 24px; font-style: italic; font-size: 14px; }

  .psz-results { list-style: none; padding: 0; margin: 0; }
  .psz-results li {
    padding: 16px 0;
    border-bottom: 1px dotted var(--psz-rule);
  }
  .psz-results li:last-child { border-bottom: none; }
  .psz-results a {
    color: var(--psz-ink); text-decoration: none;
    font-size: 18px; display: block;
    transition: color .15s;
  }
  .psz-results a:hover { color: var(--psz-accent); }
  .psz-results .ns {
    font-family: ui-monospace, Consolas, monospace;
    font-size: 10px; color: var(--psz-ink-mute);
    letter-spacing: .2em; text-transform: uppercase;
    margin-left: 8px; vertical-align: middle;
  }
  .psz-results .empty { color: var(--psz-ink-mute); font-style: italic; }
  .psz-section-title {
    font-family: ui-monospace, Consolas, monospace;
    font-size: 11px; letter-spacing: .3em; text-transform: uppercase;
    color: var(--psz-ink-mute); margin: 0 0 18px;
    padding-bottom: 8px; border-bottom: 1px solid var(--psz-rule);
  }
  .psz-count { color: var(--psz-accent); }
</style>
'@

function Render-Homepage {
    param($zim, [string]$Message)
    $articles = "{0:N0}" -f [int64]$zim.Header.ArticleCount
    $clusters = "{0:N0}" -f [int64]$zim.Header.ClusterCount
    $version  = "$($zim.Header.MajorVersion).$($zim.Header.MinorVersion)"
    $uuid     = ($zim.Header.Uuid | ForEach-Object { $_.ToString('x2') }) -join ''
    $filename = ConvertTo-SafeHtml ([System.IO.Path]::GetFileName($zim.Path))
    $fileSize = "{0:N1} Go" -f ((Get-Item $zim.Path).Length / 1GB)
    $msgHtml  = if ($Message) { "<div class='psz-msg'>$(ConvertTo-SafeHtml $Message)</div>" } else { '' }

    return @"
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PS-ZIM &mdash; Archive</title>
$Global:PageBaseCss
</head>
<body>
$Global:TopBarHtml
<header class="psz-mast">
  <div class="eyebrow">PS-ZIM &middot; Lecteur d'archive offline</div>
  <h1>Archive <span class="amp">&amp;</span> Mémoire</h1>
  <div class="filename">$filename</div>
</header>

$msgHtml

<form class="psz-search-box" action="/search" method="get" role="search">
  <input name="q" placeholder="Saisissez un titre, un mot-clé&hellip;" autocomplete="off" autofocus>
  <button type="submit">Chercher</button>
</form>

<dl class="psz-meta">
  <dt>Articles</dt><dd>$articles</dd>
  <dt>Clusters</dt><dd>$clusters</dd>
  <dt>Version ZIM</dt><dd>$version</dd>
  <dt>Taille</dt><dd>$fileSize</dd>
  <dt>UUID</dt><dd><code>$uuid</code></dd>
  <dt>Cache</dt><dd>$($zim.ClusterCache.Count) / $($zim.MaxCacheSize) clusters</dd>
</dl>

<div class="psz-actions">
  <a href="/random">&#9879;&nbsp;&nbsp;Article aléatoire</a>
  <a href="/api/info">&#9881;&nbsp;&nbsp;API info</a>
</div>
</body>
</html>
"@
}

function Render-SearchResults {
    param($zim, [string]$Query, $Results)
    $qHtml = ConvertTo-SafeHtml $Query
    $count = $Results.Count
    $itemsHtml = if ($count -eq 0) {
        "<li class='empty'>Aucun résultat pour &laquo;&nbsp;$qHtml&nbsp;&raquo;.</li>"
    } else {
        ($Results | ForEach-Object {
            $url = Get-EntryHttpUrl $zim $_
            $t   = ConvertTo-SafeHtml $_.Title
            "<li><a href=`"$url`">$t<span class='ns'>$($_.Namespace)</span></a></li>"
        }) -join "`n"
    }
    return @"
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Recherche &laquo; $qHtml &raquo; &mdash; PS-ZIM</title>
$Global:PageBaseCss
</head>
<body>
$Global:TopBarHtml
<header class="psz-mast">
  <div class="eyebrow">Recherche</div>
  <h1>&laquo;&nbsp;<em>$qHtml</em>&nbsp;&raquo;</h1>
  <div class="filename"><span class="psz-count">$count</span> résultat(s)</div>
</header>

<form class="psz-search-box" action="/search" method="get" role="search">
  <input name="q" value="$qHtml" autocomplete="off" autofocus>
  <button type="submit">Chercher</button>
</form>

<h2 class="psz-section-title">Articles trouvés</h2>
<ul class="psz-results">
$itemsHtml
</ul>
</body>
</html>
"@
}

function Render-NotFound {
    param([string]$Path)
    $p = ConvertTo-SafeHtml $Path
    return @"
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>404 &mdash; PS-ZIM</title>
$Global:PageBaseCss
</head>
<body>
$Global:TopBarHtml
<header class="psz-mast">
  <div class="eyebrow">Erreur 404</div>
  <h1>Introuvable</h1>
  <div class="filename">$p</div>
</header>
<div class="psz-actions">
  <a href="/">&#9737;&nbsp;&nbsp;Retour à l'accueil</a>
  <a href="/random">&#9879;&nbsp;&nbsp;Article aléatoire</a>
</div>
</body>
</html>
"@
}

# ─────────────────────────────────────────────────────────────────────────────
#  Dispatcher de requête
# ─────────────────────────────────────────────────────────────────────────────
function Handle-Request {
    param($zim, $ctx)
    $req = $ctx.Request
    $res = $ctx.Response

    try {
        $rawPath = $req.Url.AbsolutePath
        $path = [Uri]::UnescapeDataString($rawPath.TrimStart('/'))

        # / → page principale du ZIM si dispo, sinon homepage générée
        if ([string]::IsNullOrEmpty($path)) {
            if ($zim.Header.MainPage -ne 0xFFFFFFFF -and $zim.Header.MainPage -lt $zim.Header.ArticleCount) {
                $entry = Resolve-ZimEntry $zim ([uint32]$zim.Header.MainPage)
                Send-HttpRedirect $res (Get-EntryHttpUrl $zim $entry)
                return
            }
            Send-HttpHtml $res (Render-Homepage $zim)
            return
        }

        # /home → toujours homepage générée
        if ($path -eq 'home') {
            Send-HttpHtml $res (Render-Homepage $zim)
            return
        }

        # /search?q=
        if ($path -eq 'search') {
            $q = $req.QueryString['q']
            if ([string]::IsNullOrWhiteSpace($q)) {
                Send-HttpHtml $res (Render-Homepage $zim)
                return
            }
            Write-DebugLog "=== /search?q=$q START ==="
            $results = $null
            try {
                $results = Search-ZimTitles $zim $q 60
            } catch {
                throw "PHASE=Search-ZimTitles q='$q' -> $($_.Exception.GetType().Name): $($_.Exception.Message)"
            }
            Write-DebugLog "=== /search: Search-ZimTitles returned $($results.Count) results ==="
            $html = $null
            try {
                $html = Render-SearchResults $zim $q $results
            } catch {
                throw "PHASE=Render-SearchResults q='$q' resultsCount=$($results.Count) -> $($_.Exception.GetType().Name): $($_.Exception.Message)"
            }
            Send-HttpHtml $res $html
            Write-DebugLog "=== /search?q=$q END ==="
            return
        }

        # /api/suggest?q= → JSON pour l'autocomplétion
        if ($path -eq 'api/suggest') {
            $q = $req.QueryString['q']
            $arr = @()
            if (-not [string]::IsNullOrWhiteSpace($q)) {
                try {
                    $r = Search-ZimTitles $zim $q 10
                    $arr = @($r | ForEach-Object {
                        '{"title":' + (ConvertTo-JsonString $_.Title) + ',"url":' + (ConvertTo-JsonString (Get-EntryHttpUrl $zim $_)) + '}'
                    })
                } catch {
                    Write-DebugLog "api/suggest CAUGHT: $($_.Exception.Message)"
                    $arr = @()
                }
            }
            $json = '[' + ($arr -join ',') + ']'
            Send-HttpRaw $res ([System.Text.Encoding]::UTF8.GetBytes($json)) 'application/json; charset=utf-8'
            return
        }

        # /random
        if ($path -eq 'random') {
            for ($try = 0; $try -lt 30; $try++) {
                $idx = Get-Random -Minimum 0 -Maximum ([int]$zim.Header.ArticleCount)
                $entry = Read-ZimDirEntry $zim ([uint32]$idx)
                if (-not $entry.IsLinkTarget -and -not $entry.IsDeleted -and
                    ($entry.Namespace -eq 'A' -or $entry.Namespace -eq 'C')) {
                    if ($entry.IsRedirect) {
                        $entry = Resolve-ZimEntry $zim $entry.Index
                    }
                    $mime = Get-MimeType $zim $entry.Mime
                    if ($mime -like 'text/html*') {
                        Send-HttpRedirect $res (Get-EntryHttpUrl $zim $entry)
                        return
                    }
                }
            }
            Send-HttpText $res "Aucun article aléatoire trouvé (essayez de relancer)." 404
            return
        }

        # /api/info → JSON
        if ($path -eq 'api/info') {
            $info = [ordered]@{
                file          = [System.IO.Path]::GetFileName($zim.Path)
                articleCount  = [int64]$zim.Header.ArticleCount
                clusterCount  = [int64]$zim.Header.ClusterCount
                zimVersion    = "$($zim.Header.MajorVersion).$($zim.Header.MinorVersion)"
                mainPage      = [int64]$zim.Header.MainPage
                mimeTypes     = $zim.MimeList.Count
                cacheClusters = $zim.ClusterCache.Count
                cacheMaxSize  = $zim.MaxCacheSize
            } | ConvertTo-Json -Compress
            Send-HttpRaw $res ([System.Text.Encoding]::UTF8.GetBytes($info)) 'application/json; charset=utf-8'
            return
        }

        # /favicon.ico → 204 (pas de favicon embarquée)
        if ($path -eq 'favicon.ico') {
            Send-HttpRaw $res ([byte[]]@()) 'image/x-icon' 204
            return
        }

        # ── Résolution dans le ZIM ──
        $idx = Find-ZimEntryByPath $zim $path
        if ($idx -lt 0) {
            Send-HttpHtml $res (Render-NotFound $path) 404
            return
        }

        $entry = Resolve-ZimEntry $zim ([uint32]$idx)
        if ($entry.IsLinkTarget -or $entry.IsDeleted) {
            Send-HttpHtml $res (Render-NotFound $path) 404
            return
        }

        # ETag basé sur l'identité de l'entrée (cluster/blob/revision/mime)
        $etag = '"{0}-{1}-{2}-{3}"' -f $entry.ClusterNumber, $entry.BlobNumber, $entry.Revision, $entry.Mime
        if (Test-IfNoneMatch $req $etag) {
            Send-HttpNotModified $res $etag
            return
        }

        $blob = Get-ZimBlob $zim $entry.ClusterNumber $entry.BlobNumber
        $mime = Get-MimeType $zim $entry.Mime

        # Injecter la top-bar dans les pages HTML
        if ($mime -like 'text/html*' -and $blob.Length -gt 0) {
            $html = [System.Text.Encoding]::UTF8.GetString($blob)
            $html = Add-TopBar $html
            $blob = [System.Text.Encoding]::UTF8.GetBytes($html)
        }

        Send-HttpRaw $res $blob $mime 200 $etag
    }
    catch {
        $details = Write-DebugError ("HTTP " + $req.Url.AbsolutePath) $_
        $details = "[$Global:PsZimVersion]`n" + $details
        try { Send-HttpText $res $details 500 } catch { }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
#  Pool de runspaces (serveur HTTP concurrent)
# ─────────────────────────────────────────────────────────────────────────────
function New-PszRunspacePool {
    param([int]$MaxThreads, [hashtable]$SharedVars, [string[]]$FunctionNames, $PsHost)

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()

    foreach ($fn in $FunctionNames) {
        $cmd = Get-Command $fn -CommandType Function -ErrorAction SilentlyContinue
        if ($null -eq $cmd) { continue }
        $entry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry($fn, $cmd.Definition)
        $iss.Commands.Add($entry)
    }

    foreach ($k in $SharedVars.Keys) {
        $ve = New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry($k, $SharedVars[$k], '')
        $iss.Variables.Add($ve)
    }

    $pool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads, $iss, $PsHost)
    $pool.ApartmentState = 'MTA'
    $pool.Open()
    return $pool
}

# Liste des fonctions à rendre disponibles dans chaque runspace de travail
$Global:PszWorkerFunctions = @(
    'Write-DebugLog','Write-DebugError',
    'Read-ZimCString','Get-ZimUrlPtr','Get-ZimTitlePtr','Get-ZimClusterPtr',
    'Read-ZimDirEntry','Resolve-ZimEntry','Compare-ZimKey','Find-ZimByUrl','Find-ZimEntryByPath',
    'Get-ZimCluster','Get-ZimBlob','Search-ZimTitles','Search-ZimByUrl',
    'Test-IfNoneMatch','Send-HttpRaw','Send-HttpNotModified','Send-HttpHtml','Send-HttpText','Send-HttpRedirect',
    'ConvertTo-SafeHtml','ConvertTo-JsonString','Get-EntryHttpUrl','Get-MimeType','Add-TopBar',
    'Render-Homepage','Render-SearchResults','Render-NotFound','Handle-Request'
)

# Variables globales partagées avec les runspaces
$Global:PszSharedVars = @{
    PsZimVersion = $Global:PsZimVersion
    TopBarHtml   = $Global:TopBarHtml
    PageBaseCss  = $Global:PageBaseCss
    DebugLogFile = $Global:DebugLogFile
    DebugEnabled = $Global:DebugEnabled
}

# ─────────────────────────────────────────────────────────────────────────────
#  Démarrage du serveur
# ─────────────────────────────────────────────────────────────────────────────
$zim = New-ZimReader -Path $ZimPath -CacheSize $CacheSize
$zim.TitlePtrUsable = Test-TitlePtrListHealth $zim
Write-DebugLog "BOOT: TitlePtrUsable=$($zim.TitlePtrUsable) UrlPtrPos=$($zim.Header.UrlPtrPos) TitlePtrPos=$($zim.Header.TitlePtrPos)"

if ($MaxThreads -le 0) {
    $MaxThreads = [Math]::Max(4, [Environment]::ProcessorCount)
}

Write-Host ""
Write-PszBox @("P S - Z I M   server") 'DarkYellow'
Write-Host "  Fichier  : " -NoNewline; Write-Host ([System.IO.Path]::GetFileName($ZimPath)) -ForegroundColor White
Write-Host "  Articles : " -NoNewline; Write-Host ("{0:N0}" -f [int64]$zim.Header.ArticleCount) -ForegroundColor White
Write-Host "  Clusters : " -NoNewline; Write-Host ("{0:N0}" -f [int64]$zim.Header.ClusterCount) -ForegroundColor White
Write-Host "  ZIM ver. : " -NoNewline; Write-Host "$($zim.Header.MajorVersion).$($zim.Header.MinorVersion)" -ForegroundColor White
$searchMode = if ($zim.TitlePtrUsable) { 'par titre (TitlePtrList OK)' } else { 'par URL (TitlePtrList absente — ZIM v6.3+)' }
Write-Host "  Recherche: " -NoNewline; Write-Host $searchMode -ForegroundColor White
$concMode = if ($Sequential) { 'séquentiel (1 requête à la fois)' } else { "concurrent ($MaxThreads threads max)" }
Write-Host "  Mode     : " -NoNewline; Write-Host $concMode -ForegroundColor White
Write-Host "  Cache    : $CacheSize clusters" -ForegroundColor DarkGray
Write-Host "  URL      : " -NoNewline; Write-Host "http://${BindAddress}:$Port/" -ForegroundColor Cyan
Write-Host "  Ctrl+C pour arrêter." -ForegroundColor DarkGray
Write-Host ""

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://${BindAddress}:$Port/")
try { $listener.Start() } catch {
    throw "Impossible de démarrer le serveur sur le port $Port : $($_.Exception.Message). Essayez un autre port ou exécutez en administrateur (URL ACL)."
}

if (-not $NoBrowser) {
    Start-Sleep -Milliseconds 250
    try { Start-Process "http://${BindAddress}:$Port/" | Out-Null } catch { }
}

$pool = $null
$inflight = New-Object System.Collections.ArrayList

try {
    if ($Sequential) {
        # ── Boucle synchrone (mode robuste de repli) ──
        while ($listener.IsListening) {
            $ctx = $listener.GetContext()
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            Handle-Request $zim $ctx
            $sw.Stop()
            Write-Host ("  {0,3} {1,6:N0}ms  {2}" -f $ctx.Response.StatusCode, $sw.ElapsedMilliseconds, $ctx.Request.Url.AbsolutePath) -ForegroundColor DarkGray
        }
    } else {
        # ── Boucle concurrente (pool de runspaces) ──
        $pool = New-PszRunspacePool -MaxThreads $MaxThreads -SharedVars $Global:PszSharedVars -FunctionNames $Global:PszWorkerFunctions -PsHost $Host

        while ($listener.IsListening) {
            $ctx = $listener.GetContext()

            $ps = [powershell]::Create()
            $ps.RunspacePool = $pool
            [void]$ps.AddCommand('Handle-Request').AddArgument($zim).AddArgument($ctx)
            $async = $ps.BeginInvoke()

            [void]$inflight.Add([pscustomobject]@{
                PS    = $ps
                Async = $async
                Ctx   = $ctx
                SW    = [System.Diagnostics.Stopwatch]::StartNew()
            })

            # Récupère les requêtes terminées (libère les runspaces + log)
            for ($i = $inflight.Count - 1; $i -ge 0; $i--) {
                $job = $inflight[$i]
                if ($job.Async.IsCompleted) {
                    try { $job.PS.EndInvoke($job.Async) | Out-Null } catch { }
                    $job.SW.Stop()
                    try {
                        Write-Host ("  {0,3} {1,6:N0}ms  {2}" -f $job.Ctx.Response.StatusCode, $job.SW.ElapsedMilliseconds, $job.Ctx.Request.Url.AbsolutePath) -ForegroundColor DarkGray
                    } catch { }
                    $job.PS.Dispose()
                    $inflight.RemoveAt($i)
                }
            }
        }
    }
} catch [System.Net.HttpListenerException] {
    # Listener arrêté proprement
} catch {
    Write-Host "Erreur fatale : $($_.Exception.Message)" -ForegroundColor Red
} finally {
    try { $listener.Stop(); $listener.Close() } catch { }
    foreach ($job in $inflight) {
        try { $job.PS.Dispose() } catch { }
    }
    if ($pool) { try { $pool.Close(); $pool.Dispose() } catch { } }
    try { $zim.Reader.Dispose(); $zim.Stream.Dispose() } catch { }
    if (Test-Path $PidFile) { Remove-Item $PidFile -Force -ErrorAction SilentlyContinue }
    Write-Host "PS-ZIM arrêté." -ForegroundColor Yellow
}

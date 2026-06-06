# PS-ZIM

**Lecteur de fichiers ZIM en PowerShell pur, servi en HTTP local.**

PS-ZIM ouvre un fichier `.zim` (Wikipédia offline, dump Stack Exchange, etc.), monte un mini serveur HTTP local et lance votre navigateur. Il décompresse à la volée les clusters Zstandard des ZIM modernes (Kiwix) — le tout sans installer Kiwix, sans dépendance externe lourde, et compatible aussi bien avec PowerShell 5.1 qu'avec PowerShell 7.x.

---

## Sommaire

- [Caractéristiques](#caractéristiques)
- [Prérequis](#prérequis)
- [Installation](#installation)
- [Utilisation](#utilisation)
- [Paramètres](#paramètres)
- [Où trouver des fichiers ZIM](#où-trouver-des-fichiers-zim)
- [Fonctionnement](#fonctionnement)
- [Compatibilité des fichiers ZIM](#compatibilité-des-fichiers-zim)
- [Dépannage](#dépannage)
- [Limitations connues](#limitations-connues)
- [Licence](#licence)

---

## Caractéristiques

- **Parseur ZIM binaire complet** : header, liste MIME, entrées d'annuaire (dir entries), clusters.
- **Décompression Zstandard** via `libzstd.dll` natif (P/Invoke pur). La DLL officielle de Facebook/Meta est téléchargée automatiquement au premier lancement.
- **Cache LRU de clusters** décompressés, taille configurable.
- **Recherche par préfixe de titre** (binary search sur la title pointer list), avec repli automatique sur la recherche par URL pour les ZIM récents.
- **Article aléatoire**, page d'accueil avec métadonnées, et barre de navigation injectée dans chaque article.
- **Mode premier plan ou arrière-plan** (`-Background`) avec fichier PID.
- **Compatible ZIM v5** (namespaces legacy) **et v6** (namespace `C`).
- **Zéro dépendance NuGet** : uniquement `libzstd.dll` auto-téléchargée.

---

## Prérequis

- **Windows** avec PowerShell **5.1** ou **PowerShell 7.x**.
- Accès Internet **au premier lancement uniquement** (pour télécharger `libzstd.dll` depuis GitHub). Une fois la DLL en cache dans `./lib/`, PS-ZIM fonctionne entièrement hors-ligne.
- Un fichier `.zim`.

---

## Installation

```powershell
# Cloner le dépôt
git clone https://github.com/<votre-compte>/PS-ZIM.git
cd PS-ZIM
```

Aucune compilation n'est nécessaire. La bibliothèque `libzstd.dll` est récupérée automatiquement et placée dans le sous-dossier `lib/` lors de la première exécution.

> **Note sur la politique d'exécution** : si vous n'avez jamais exécuté de script PowerShell, lancez le script avec `-ExecutionPolicy Bypass` (voir les exemples ci-dessous) ou ajustez votre politique d'exécution.

---

## Utilisation

Lancer le serveur sur le port par défaut (8642) et ouvrir le navigateur :

```powershell
.\PS-ZIM.ps1 .\wikipedia_fr_all_nopic_2024-06.zim
```

Lancer en arrière-plan sur un port personnalisé :

```powershell
.\PS-ZIM.ps1 .\wikipedia.zim -Background -Port 8888
```

Arrêter une instance lancée en arrière-plan :

```powershell
.\PS-ZIM.ps1 -Stop
```

Contourner la politique d'exécution si nécessaire :

```powershell
powershell -ExecutionPolicy Bypass -File .\PS-ZIM.ps1 .\wikipedia.zim
```

Une fois démarré, ouvrez `http://127.0.0.1:8642/` dans votre navigateur. Vous y trouverez :

- une page d'accueil avec les métadonnées du fichier et une barre de recherche ;
- une barre de navigation fixe en haut de chaque article (accueil, recherche, article aléatoire) ;
- `/random` pour un article aléatoire ;
- `/api/info` pour les métadonnées du ZIM au format JSON.

---

## Paramètres

| Paramètre        | Type     | Défaut        | Description |
|------------------|----------|---------------|-------------|
| `-ZimPath`       | string   | *(requis)*    | Chemin du fichier `.zim` à servir. Alias : `-Path`, `-File`. Position 0. |
| `-Port`          | int      | `8642`        | Port HTTP local. |
| `-BindAddress`   | string   | `127.0.0.1`   | Interface d'écoute. Mettre `0.0.0.0` pour exposer au LAN. |
| `-CacheSize`     | int      | `64`          | Nombre de clusters décompressés gardés en mémoire (≈ 1–4 Mo chacun). |
| `-NoBrowser`     | switch   | *(off)*       | Ne pas ouvrir le navigateur automatiquement. |
| `-Background`    | switch   | *(off)*       | Lance le serveur dans un processus PowerShell détaché (fenêtre cachée), avec fichier PID. |
| `-Stop`          | switch   | *(off)*       | Arrête une instance lancée précédemment avec `-Background`. |

---

## Où trouver des fichiers ZIM

Les fichiers ZIM modernes (compatibles avec PS-ZIM) sont disponibles sur la bibliothèque Kiwix :

- **https://library.kiwix.org**

Vous y trouverez des dumps de Wikipédia, Wiktionnaire, Stack Exchange, Project Gutenberg, et bien d'autres, dans de nombreuses langues et tailles.

> Privilégiez les versions récentes (≥ 2020) : voir la section [Compatibilité](#compatibilité-des-fichiers-zim).

---

## Fonctionnement

PS-ZIM implémente directement le format binaire ZIM en PowerShell :

1. **Bootstrap** — au premier lancement, `libzstd.dll` (binaire officiel Facebook/Meta correspondant à l'architecture du processus) est téléchargée et placée dans `./lib/`. Elle est ensuite chargée explicitement via `LoadLibrary` avant tout appel P/Invoke.
2. **Lecture du header** — le magic number est vérifié, puis les positions des différentes tables (URL pointer list, title pointer list, cluster pointer list, MIME list) sont lues.
3. **Résolution d'un chemin HTTP** — le chemin demandé est résolu en entrée ZIM via une recherche dichotomique (binary search) sur les clés `(namespace, url)`, avec gestion des namespaces v5 et v6 et des redirections.
4. **Décompression de cluster** — le cluster contenant l'article est lu, décompressé (Zstandard ou non compressé) et mis en cache LRU. Les offsets internes (blobs) sont ensuite extraits.
5. **Injection de la barre de navigation** — pour les pages `text/html`, une barre fixe est insérée juste après la balise `<body>`.
6. **Recherche** — la title pointer list est testée au démarrage : si elle est saine, la recherche se fait par titre ; sinon (ZIM v6.3+ récents), PS-ZIM bascule automatiquement sur une recherche par URL.

Le serveur HTTP est synchrone et basé sur `System.Net.HttpListener`, ce qui le rend simple et robuste.

---

## Compatibilité des fichiers ZIM

| Compression du cluster | Support |
|------------------------|---------|
| Non compressé          | ✅ Oui |
| Zstandard (ZIM modernes / Kiwix) | ✅ Oui (via `libzstd.dll`) |
| zlib                   | ❌ Non (très ancien ZIM) |
| bzip2                  | ❌ Non (très ancien ZIM) |
| LZMA2 / XZ             | ❌ Non (ZIM antérieurs à 2020) |

| Version ZIM | Support |
|-------------|---------|
| v5 (namespaces legacy `A`, `I`, …) | ✅ Oui |
| v6 (namespace `C`)                 | ✅ Oui |
| v6.3+ (title pointer list déplacée) | ✅ Oui (repli automatique sur recherche par URL) |

> Si vous obtenez une erreur de type LZMA2/XZ, re-téléchargez une version moderne du fichier depuis https://library.kiwix.org.

---

## Dépannage

**Échec du chargement de `libzstd.dll` (LoadLibrary)**
La DLL ne correspond pas à l'architecture du processus PowerShell. Vérifiez si vous êtes en 32 ou 64 bits et supprimez `./lib/libzstd.dll` pour forcer un nouveau téléchargement adapté.

**Échec du téléchargement de libzstd**
Vérifiez votre connexion Internet et l'accès à `github.com`. Vous pouvez aussi placer manuellement le bon `libzstd.dll` dans le dossier `lib/`.

**« Impossible de démarrer le serveur sur le port … »**
Le port est déjà occupé, ou une réservation d'URL ACL est requise. Essayez un autre port (`-Port`) ou exécutez PowerShell en administrateur.

**La recherche ne renvoie rien sur un ZIM récent**
PS-ZIM bascule normalement en recherche par URL. Au démarrage, la console indique le mode utilisé (« par titre » ou « par URL »). Wikipédia normalisant les URL en `Titre_avec_underscores`, essayez des mots-clés proches du titre exact.

**Logs de debug**
Un fichier `ps-zim-debug.log` est généré dans le dossier du script à chaque démarrage (remis à zéro à chaque lancement). Il contient le détail des recherches et des erreurs rencontrées.

**Exposer au réseau local**
Utilisez `-BindAddress 0.0.0.0`. Attention : cela rend l'archive accessible à toute machine du réseau. Une réservation d'URL ACL ou des droits administrateur peuvent être nécessaires.

---

## Limitations connues

- Windows uniquement (dépend de `HttpListener`, de `kernel32.dll` et d'un `libzstd.dll` Windows).
- Pas de support des anciennes compressions zlib / bzip2 / LZMA2.
- Serveur HTTP synchrone (mono-requête à la fois) — conçu pour un usage local personnel, pas pour servir de nombreux clients simultanés.

---

## Licence

À définir par l'auteur du dépôt (par exemple MIT).

PS-ZIM télécharge et utilise `libzstd.dll`, distribuée par le projet [zstd](https://github.com/facebook/zstd) (Facebook/Meta) sous licence BSD/GPLv2. Le format ZIM est une spécification ouverte du projet [openZIM](https://openzim.org/).


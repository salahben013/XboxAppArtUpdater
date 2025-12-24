# XboxAppArtUpdater (Win32/C++)

## Overview
XboxAppArtUpdater is a Windows tool to scan, back up, and update third-party game icons in the Xbox app. It supports multiple stores, profiles for icon sets, and can auto-download high-quality icons from SteamGridDB and a search through Google images.

## Features
- **Scan Xbox app third-party library cache** for game icons:
  - `%LOCALAPPDATA%\Packages\Microsoft.GamingApp_8wekyb3d8bbwe\LocalState\ThirdPartyLibraries`
- **List detected icons** for supported stores:
  - Steam
  - Epic
  - GOG
  - Ubisoft Connect (Ubi)
- **Backup and restore** original icons before making changes
- **Profile support**: Manage and switch between different sets of icons using the `profiles/` directory
- **SteamGridDB integration** (requires API key):
  - Automatically resolve game titles and download high-quality square PNG icons
- **Configurable settings** via environment variable or `settings.ini`:
  - API key, timeouts, preferred icon size, backup toggle
- **Planned UI**:
  - Main: Scan, grid list with thumbnails
  - Settings: API key, timeouts, icon size, backup toggle
  - Auto update: Pick highest-res icon automatically
  - Manual update: Show candidate icons and let user choose

## Usage Notes
- **Close the Xbox app before replacing images** to avoid file locks.
- Provide your SteamGridDB API key via:
  - Environment variable: `STEAMGRIDDB_API_KEY`
  - Or a local config file: `settings.ini` next to the executable

## Build Instructions
- Run `build.bat` (requires MSVC Build Tools)

## Directory Structure
- `backupArt/` — Backups of original icons by store
- `profiles/` — Profile folders for different icon sets
- `icons/` — Custom or downloaded icons

## Supported Stores
- Steam
- Epic
- GOG
- Ubisoft Connect (Ubi)

> **Note:** Battle.net (Bnet) is not supported yet.

## Profile Support
Profiles allow you to maintain multiple sets of icons. Each profile is stored in its own subdirectory under `profiles/` (e.g., `profiles/20251223041550-test1/`). You can switch between profiles to quickly change the appearance of your Xbox app library.

# XboxAppArtUpdater (Win32/C++)

## Overview
XboxAppArtUpdater is a Windows tool to scan, back up, and update third-party game icons in the Xbox app. It supports multiple stores, profiles for icon sets, and can auto-download high-quality icons from SteamGridDB or Google Images.

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
   - Automatically resolve game titles and download high-quality PNG icons
- **Automated Web Search**: Fetches game covers from Google Images (no API key required)
- **Configurable settings** via Config menu:
   - API key, timeouts, preferred icon size, backup toggle

## Quick Start - Automated Web Search
Quick Apply now includes automatic web search using Google Images. When enabled, it fetches game covers directly from the web without needing a SteamGridDB API key.

**Important notes about web search:**
- Web results are NOT curated and may be wrong or low quality
- Results can be unpredictable and vary between searches
- You can choose 'Random' or select a specific result number
- First results (1-3) are usually more accurate than later ones
- Use specific result numbers for more consistent selection
- For best quality, use SteamGridDB sources instead

## How to Use This App

  You can skip step 1 and 2 if you don't want to use SteamGridDB, but in some cases where the name is not correctly resolved is a good source in this case SteamGridDb.

1. **Get a SteamGridDB API Key** (Optional for web-only mode)
    - Go to https://www.steamgriddb.com
    - Create an account or sign in
    - Go to your Profile > Preferences > API
    - Copy your API key

2. **Set Up the App**
    - Click the "Config" button (top right)
    - Paste your SteamGridDB API key
    - Optionally set a custom ThirdPartyLibraries path
    - Choose Theme (Dark/Light) and Parallel Downloads
    - Click OK to save

3. **Scan for Games**
    - Click the "Scan" button to find games
    - The app will scan Steam, Epic, GOG, and Ubisoft
    - Games will appear in the list with their current art

4. **Update Game Art**
    - Double-click any game in the list
    - Browse available art from SteamGridDB or Google Images
    - Use tabs to switch between Web, Grids, Heroes, Logos, Icons
    - Select an image and click "Apply" to update

5. **Restore Original Art**
    - Open the art window for a game
    - Click "Restore Original" to revert changes
    - Only works if a backup was created

## QUICK START - AUTOMATED WEB SEARCH:

    Quick Apply now includes automatic web search using Google Images.
    When enabled, it fetches game covers directly from the web without
    needing a SteamGridDB API key, you can also add the SteamGridDB as a source in this case the search choses randomly between the selected sources.

## IMPORTANT NOTES ABOUT WEB SEARCH:
    - Web results are NOT curated and may be wrong or low quality
    - Results can be unpredictable and vary between searches
    - You can choose 'Random' or select a specific result number
    - First results (1-3) are usually more accurate than later ones
    - Use specific result numbers for more consistent selection

## Filters
- **Store**: Filter by game store (Steam, Epic, GOG, Ubisoft)
- **Art**: Show all games, only games with art, or missing art
- **Search**: Type to filter games by title
- **Unresolved titles**: Check the box to get only games with unresolved titles (where the ID wasn't enough to identify the game's title).

## Tips
- Games with "(missing)" have no custom art yet
- Art is cached locally for faster loading
- Switch between List and Grid view using the Layout dropdown
- Change icon sizes with the Icon size dropdown

## Parallel Downloads
- Configure how many images download simultaneously
- Set in Config > Parallel Downloads dropdown
- "Auto" uses your CPU thread count for best performance
- Lower values (2-4) may help on slower connections

## Minimum Image Size
- Filter out small/blurry images from search results
- Set in Config > Min Image Size (px) field
- Default is 200px - images smaller than this are hidden
- Set to 0 to show all images regardless of size
- Affects both manual art selection and Quick Apply

## Quick Apply
- Select one or more games in the list
- Click "Quick Apply" to apply art automatically
- Choose sources: Web Search, Grids, Heroes, Logos, Icons
- Web Search: Automatic Google Images search (may be unpredictable)
   * Random: Picks a random result from search
   * Result #: Select specific result (1=first, usually best)
   * First 3 results are typically most accurate
- SteamGridDB: Curated, high-quality community art
- Shows results with success/failure details
- Debug window (single game) shows URLs and selection
- Uses parallel processing for faster batch updates

## Read-Only Protection
- Applied art is automatically marked as read-only
- Prevents Xbox app from overwriting your custom art
- Files are unlocked temporarily when updating art
- Restoring original art also removes read-only flag

## Cache Management
- Thumbnails are cached locally for faster loading
- Clear all cache: Config > "Clear Image Cache" button
- Clear per-game: Art window > "Clear Cache" button
- Cache is stored in %TEMP%\XboxAppUpdaterArtCache
- Safe to clear while browsing - images auto re-download
- Per-game clear only affects that game's thumbnails

## Profile Support
Profiles allow you to maintain multiple sets of icons. Each profile is stored in its own subdirectory under `profiles/` (e.g., `profiles/20251223041550-test1/`). You can switch between profiles to quickly change the appearance of your Xbox app library.

## Paths
- **Default ThirdPartyLibraries:**
   C:\Users\%USERNAME%\AppData\Local\Packages\Microsoft.GamingApp_8wekyb3d8bbwe\LocalState\ThirdPartyLibraries
- **Art Backups Location:**
   C:\XboxAppArtUpdater\backupArt

## Build Instructions
- Run `build.bat` (requires MSVC Build Tools)

## Directory Structure
- `backupArt/` — Backups of original icons by store
- `profiles/` — Profile folders for different icon sets
- `icons/` — Custom or downloaded icons

// --- Utility function for unresolved checkbox (file-scope, after globals) ---
// (Implementation moved after all globals and type definitions)
// XboxAppArtUpdater.cpp
#define NOMINMAX
#include <windows.h>
#include <commctrl.h>
#include <shlobj.h>

#pragma comment(linker,"/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#include <shellapi.h>
#include <gdiplus.h>
#include <winhttp.h>
#include <richedit.h>

#include <string>
#include <vector>
#include <filesystem>
#include <fstream>
#include <regex>
#include <unordered_map>
#include <algorithm>
#include <cstring>
#include <cstdlib>
#include <ctime>
#include <thread>
#include <mutex>
#include <atomic>
#include <queue>

#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Winhttp.lib")
#pragma comment(lib, "Gdiplus.lib")
#pragma comment(lib, "Dwmapi.lib")
#pragma comment(lib, "UxTheme.lib")

#include <dwmapi.h>
#include <uxtheme.h>

// Dark mode support (Windows 10 1809+)
#ifndef DWMWA_USE_IMMERSIVE_DARK_MODE
#define DWMWA_USE_IMMERSIVE_DARK_MODE 20
#endif

// Undocumented dark mode APIs from uxtheme.dll (ordinal exports)
// These are used by Windows itself for dark mode support
enum PreferredAppMode { Default, AllowDark, ForceDark, ForceLight, Max };
using fnSetPreferredAppMode = PreferredAppMode(WINAPI*)(PreferredAppMode appMode);
using fnAllowDarkModeForWindow = bool(WINAPI*)(HWND hWnd, bool allow);
using fnRefreshImmersiveColorPolicyState = void(WINAPI*)();
using fnFlushMenuThemes = void(WINAPI*)();

static fnSetPreferredAppMode SetPreferredAppMode = nullptr;
static fnAllowDarkModeForWindow AllowDarkModeForWindow = nullptr;
static fnRefreshImmersiveColorPolicyState RefreshImmersiveColorPolicyState = nullptr;
static fnFlushMenuThemes FlushMenuThemes = nullptr;
static bool gDarkModeApisLoaded = false;

static void LoadDarkModeApis() {
    if (gDarkModeApisLoaded) return;
    gDarkModeApisLoaded = true;
    
    HMODULE hUxTheme = GetModuleHandleW(L"uxtheme.dll");
    if (!hUxTheme) hUxTheme = LoadLibraryW(L"uxtheme.dll");
    if (!hUxTheme) return;
    
    // These are exported by ordinal, not by name
    SetPreferredAppMode = (fnSetPreferredAppMode)GetProcAddress(hUxTheme, MAKEINTRESOURCEA(135));
    AllowDarkModeForWindow = (fnAllowDarkModeForWindow)GetProcAddress(hUxTheme, MAKEINTRESOURCEA(133));
    RefreshImmersiveColorPolicyState = (fnRefreshImmersiveColorPolicyState)GetProcAddress(hUxTheme, MAKEINTRESOURCEA(104));
    FlushMenuThemes = (fnFlushMenuThemes)GetProcAddress(hUxTheme, MAKEINTRESOURCEA(136));
}

// Forward declaration - defined after gDarkMode is declared
static void DrawDarkButton(DRAWITEMSTRUCT* dis);

using namespace Gdiplus;
namespace fs = std::filesystem;

// ------------------------------------------------------------
// Globals / constants
// ------------------------------------------------------------

static const wchar_t* APP_TITLE = L"XboxAppArtUpdater";

// ...existing code...

static const int IDC_BTN_SCAN       = 1001;
static const int IDC_LIST           = 1002;
static const int IDC_CMB_SIZE       = 1003;
static const int IDC_CMB_LAYOUT     = 1004;
static const int IDC_BTN_CONFIG     = 1005;
static const int IDC_STATUS_BAR     = 1006;
static const int IDC_CMB_STORE      = 1007;
static const int IDC_CMB_ART_FILTER = 1008;
static const int IDC_SEARCH_EDIT    = 1009;
static const int IDC_BTN_HELP       = 1010;
static const int IDC_BTN_QUICK_APPLY = 1011;
static const int IDC_BTN_SELECT_ALL = 1012;
static const int IDC_BTN_PROFILES   = 1013;

static HWND gList = nullptr;
static HWND gNoScanMessage = nullptr;  // Overlay message when no scan has been done
static HWND gCmbSize = nullptr;
static HWND gCmbLayout = nullptr;
static HWND gStatusBar = nullptr;
static HWND gCmbStore = nullptr;
static HWND gCmbArtFilter = nullptr;
static HWND gSearchEdit = nullptr;
static HWND gBtnConfig = nullptr;
static HWND gBtnScan = nullptr;
static HWND gBtnHelp = nullptr;
static HWND gBtnQuickApply = nullptr;
static HWND gBtnSelectAll = nullptr;
static HWND gBtnProfiles = nullptr;
static HWND gListInstructionNote = nullptr;
static HWND gMainWnd = nullptr;
static HWND gTooltip = nullptr;

static ULONG_PTR gGdiplusToken = 0;

static HIMAGELIST gImgList = nullptr;
static int gPlaceholderImgIndex = -1;

static int  gIconSize = 64;      // 32/64/128
static bool gGridLayout = false; // false=list, true=grid
static bool gDarkMode = true;    // true=dark, false=light
static bool gHasScanned = false; // true after first scan (changes Scan button to Refresh)
static bool gQuickApplyRunning = false; // true while Quick Apply is processing

// Quick Apply configuration
struct QuickApplyConfig {
    bool useWeb = true;
    std::wstring webResolution = L"1024x1024";
    bool webRandom = false;  // true = random, false = use specific index
    int webResultIndex = 1;  // 1-based index (1 = first result)
    bool useGrids = false;
    bool useHeroes = false;
    bool useLogos = false;
    bool useIcons = false;
    bool hasShownQuickApplyIntro = false;  // Track if we've shown the intro dialog
};
static QuickApplyConfig gQuickApplyConfig;

// Draw an owner-draw button in dark mode style
static void DrawDarkButton(DRAWITEMSTRUCT* dis) {
    HDC hdc = dis->hDC;
    RECT rc = dis->rcItem;
    
    bool isPressed = (dis->itemState & ODS_SELECTED) != 0;
    bool isFocused = (dis->itemState & ODS_FOCUS) != 0;
    bool isDisabled = (dis->itemState & ODS_DISABLED) != 0;
    
    // Check if this is the Scan button (for special red color before first scan)
    bool isScanButton = (dis->CtlID == 1001);  // IDC_BTN_SCAN
    // Red buttons: Scan (before first scan) and Config button in art dialog - in both modes
    bool showRed = (isScanButton && !gHasScanned) || (dis->CtlID == 4013);  // IDC_ART_CONFIG_BTN
    
    // Check if this is an "action" button that should be cyan blue
    bool isCyanButton = (dis->CtlID == 3006 ||   // IDC_CFG_OK_BTN
                         dis->CtlID == 4002 ||   // IDC_ART_APPLY_BTN
                         dis->CtlID == 1011);    // IDC_BTN_QUICK_APPLY
    
    // Colors
    COLORREF bgColor, textColor, borderColor;
    if (gDarkMode) {
        if (isDisabled) {
            bgColor = RGB(50, 50, 50);
            textColor = RGB(128, 128, 128);
            borderColor = RGB(70, 70, 70);
        } else if (showRed) {
            // Red color for special buttons
            bgColor = isPressed ? RGB(180, 50, 50) : RGB(150, 40, 40);
            textColor = RGB(255, 255, 255);
            borderColor = RGB(200, 60, 60);
        } else if (isCyanButton) {
            // Cyan blue for OK/Apply buttons
            bgColor = isPressed ? RGB(0, 140, 180) : RGB(0, 120, 160);
            textColor = RGB(255, 255, 255);
            borderColor = RGB(0, 160, 200);
        } else if (isPressed) {
            bgColor = RGB(80, 80, 80);
            textColor = RGB(255, 255, 255);
            borderColor = RGB(100, 100, 100);
        } else {
            bgColor = RGB(60, 60, 60);
            textColor = RGB(255, 255, 255);
            borderColor = RGB(100, 100, 100);
        }
    } else {
        // Light mode - clean white buttons
        if (isDisabled) {
            bgColor = RGB(240, 240, 240);
            textColor = RGB(160, 160, 160);
            borderColor = RGB(200, 200, 200);
        } else if (showRed) {
            // Red color for special buttons in light mode too
            bgColor = isPressed ? RGB(200, 60, 60) : RGB(180, 50, 50);
            textColor = RGB(255, 255, 255);
            borderColor = RGB(150, 40, 40);
        } else if (isCyanButton) {
            // Cyan blue for OK/Apply buttons in light mode
            bgColor = isPressed ? RGB(0, 140, 180) : RGB(0, 120, 160);
            textColor = RGB(255, 255, 255);
            borderColor = RGB(0, 100, 140);
        } else if (isPressed) {
            bgColor = RGB(230, 230, 230);
            textColor = RGB(0, 0, 0);
            borderColor = RGB(160, 160, 160);
        } else {
            bgColor = RGB(255, 255, 255);
            textColor = RGB(0, 0, 0);
            borderColor = RGB(180, 180, 180);
        }
    }
    
    // Fill background
    HBRUSH bgBrush = CreateSolidBrush(bgColor);
    FillRect(hdc, &rc, bgBrush);
    DeleteObject(bgBrush);
    
    // Draw border
    HPEN borderPen = CreatePen(PS_SOLID, 1, borderColor);
    HPEN oldPen = (HPEN)SelectObject(hdc, borderPen);
    HBRUSH oldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, 4, 4);
    SelectObject(hdc, oldPen);
    SelectObject(hdc, oldBrush);
    DeleteObject(borderPen);
    
    // Focus rectangle
    if (isFocused) {
        RECT focusRc = { rc.left + 2, rc.top + 2, rc.right - 2, rc.bottom - 2 };
        HPEN focusPen = CreatePen(PS_DOT, 1, gDarkMode ? RGB(150, 150, 150) : RGB(100, 100, 100));
        oldPen = (HPEN)SelectObject(hdc, focusPen);
        oldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
        Rectangle(hdc, focusRc.left, focusRc.top, focusRc.right, focusRc.bottom);
        SelectObject(hdc, oldPen);
        SelectObject(hdc, oldBrush);
        DeleteObject(focusPen);
    }
    
    // Get button text
    wchar_t text[256]{};
    GetWindowTextW(dis->hwndItem, text, _countof(text));
    
    // Draw text
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textColor);
    
    RECT textRc = rc;
    if (isPressed) {
        OffsetRect(&textRc, 1, 1);
    }
    DrawTextW(hdc, text, -1, &textRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
}

// Draw an owner-draw tab in dark mode style
static void DrawDarkTab(DRAWITEMSTRUCT* dis, HWND hTab) {
    HDC hdc = dis->hDC;
    RECT rc = dis->rcItem;
    
    bool isSelected = (dis->itemState & ODS_SELECTED) != 0;
    
    // Get tab text
    wchar_t text[256]{};
    TCITEMW tci{};
    tci.mask = TCIF_TEXT;
    tci.pszText = text;
    tci.cchTextMax = _countof(text);
    TabCtrl_GetItem(hTab, dis->itemID, &tci);
    
    // Colors
    COLORREF bgColor, textColor, borderColor;
    if (gDarkMode) {
        if (isSelected) {
            bgColor = RGB(45, 45, 45);
            textColor = RGB(255, 255, 255);
            borderColor = RGB(200, 200, 200);  // Strong light border for active
        } else {
            bgColor = RGB(32, 32, 32);
            textColor = RGB(180, 180, 180);
            borderColor = RGB(60, 60, 60);  // Faint dark border for inactive
        }
    } else {
        // Light mode
        if (isSelected) {
            bgColor = RGB(255, 255, 255);  // White background for active
            textColor = RGB(0, 0, 0);
            borderColor = RGB(0, 0, 0);  // Strong black border for active
        } else {
            bgColor = RGB(255, 255, 255);
            textColor = RGB(100, 100, 100);
            borderColor = RGB(200, 200, 200);  // Faint light grey border for inactive
        }
    }
    
    // Fill background - inflate to cover any artifacts
    RECT fillRc = rc;
    InflateRect(&fillRc, 2, 2);
    HBRUSH bgBrush = CreateSolidBrush(bgColor);
    FillRect(hdc, &fillRc, bgBrush);
    DeleteObject(bgBrush);
    
    // Draw border around the tab
    HPEN borderPen = CreatePen(PS_SOLID, 1, borderColor);
    HPEN oldPen = (HPEN)SelectObject(hdc, borderPen);
    
    // Draw all four sides
    MoveToEx(hdc, rc.left, rc.bottom - 1, nullptr);
    LineTo(hdc, rc.left, rc.top);
    LineTo(hdc, rc.right - 1, rc.top);
    LineTo(hdc, rc.right - 1, rc.bottom - 1);
    LineTo(hdc, rc.left, rc.bottom - 1);
    
    SelectObject(hdc, oldPen);
    DeleteObject(borderPen);
    
    // Draw text with padding
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textColor);
    
    RECT textRc = rc;
    InflateRect(&textRc, -4, -2);  // Add padding: 4px horizontal, 2px vertical
    DrawTextW(hdc, text, -1, &textRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
}

// Draw an owner-draw combobox item in dark mode style
static void DrawDarkComboItem(DRAWITEMSTRUCT* dis) {
    if (dis->itemID == (UINT)-1) return;
    
    HDC hdc = dis->hDC;
    RECT rc = dis->rcItem;
    
    bool isSelected = (dis->itemState & ODS_SELECTED) != 0;
    bool isFocused = (dis->itemState & ODS_FOCUS) != 0;
    bool isComboEdit = (dis->itemState & ODS_COMBOBOXEDIT) != 0;
    
    // Colors
    COLORREF bgColor, textColor;
    if (gDarkMode) {
        if (isSelected && !isComboEdit) {
            bgColor = RGB(60, 60, 60);
            textColor = RGB(255, 255, 255);
        } else {
            bgColor = RGB(45, 45, 45);
            textColor = RGB(255, 255, 255);
        }
    } else {
        bgColor = GetSysColor(isSelected ? COLOR_HIGHLIGHT : COLOR_WINDOW);
        textColor = GetSysColor(isSelected ? COLOR_HIGHLIGHTTEXT : COLOR_WINDOWTEXT);
    }
    
    // Fill background
    HBRUSH bgBrush = CreateSolidBrush(bgColor);
    FillRect(hdc, &rc, bgBrush);
    DeleteObject(bgBrush);
    
    // Get item text
    wchar_t text[256]{};
    SendMessageW(dis->hwndItem, CB_GETLBTEXT, dis->itemID, (LPARAM)text);
    
    // Draw text
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, textColor);
    
    RECT textRc = rc;
    textRc.left += 4;
    DrawTextW(hdc, text, -1, &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
    
    // Draw focus rect if needed
    if (isFocused && !isComboEdit) {
        DrawFocusRect(hdc, &rc);
    }
}

// Edit control subclass for placeholder text (works in both dark and light mode)
static WNDPROC gOriginalEditProc = nullptr;
static std::unordered_map<HWND, std::wstring> gEditPlaceholders;

static LRESULT CALLBACK EditSubclassProc(HWND hWnd, UINT msg, WPARAM w, LPARAM l) {
    if (msg == WM_PAINT) {
        // Call original paint first
        LRESULT result = CallWindowProcW(gOriginalEditProc, hWnd, msg, w, l);
        
        // Check if edit is empty and not focused
        int textLen = GetWindowTextLengthW(hWnd);
        HWND focusWnd = GetFocus();
        
        auto it = gEditPlaceholders.find(hWnd);
        if (textLen == 0 && focusWnd != hWnd && it != gEditPlaceholders.end() && !it->second.empty()) {
            HDC hdc = GetDC(hWnd);
            RECT rc;
            GetClientRect(hWnd, &rc);
            
            // Draw placeholder text in gray
            SetBkMode(hdc, TRANSPARENT);
            SetTextColor(hdc, RGB(128, 128, 128));
            
            HFONT hFont = (HFONT)SendMessageW(hWnd, WM_GETFONT, 0, 0);
            HFONT oldFont = nullptr;
            if (hFont) oldFont = (HFONT)SelectObject(hdc, hFont);
            
            rc.left += 4;  // Padding
            DrawTextW(hdc, it->second.c_str(), -1, &rc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
            
            if (oldFont) SelectObject(hdc, oldFont);
            ReleaseDC(hWnd, hdc);
        }
        return result;
    }
    else if (msg == WM_SETFOCUS || msg == WM_KILLFOCUS) {
        // Redraw to show/hide placeholder
        LRESULT result = CallWindowProcW(gOriginalEditProc, hWnd, msg, w, l);
        InvalidateRect(hWnd, nullptr, TRUE);
        return result;
    }
    return CallWindowProcW(gOriginalEditProc, hWnd, msg, w, l);
}

static void SetEditPlaceholder(HWND hEdit, const std::wstring& placeholder) {
    gEditPlaceholders[hEdit] = placeholder;
    
    // Subclass if not already
    WNDPROC currentProc = (WNDPROC)GetWindowLongPtrW(hEdit, GWLP_WNDPROC);
    if (currentProc != EditSubclassProc) {
        if (!gOriginalEditProc) gOriginalEditProc = currentProc;
        SetWindowLongPtrW(hEdit, GWLP_WNDPROC, (LONG_PTR)EditSubclassProc);
    }
}

// Combobox subclass for custom dropdown button styling
static WNDPROC gOriginalComboProc = nullptr;

static LRESULT CALLBACK ComboSubclassProc(HWND hWnd, UINT msg, WPARAM w, LPARAM l) {
    if (msg == WM_PAINT) {
        // Call original paint first
        LRESULT result = CallWindowProcW(gOriginalComboProc, hWnd, msg, w, l);
        
        HDC hdc = GetDC(hWnd);
        RECT rc;
        GetClientRect(hWnd, &rc);
        
        // Colors based on mode
        COLORREF bgColor = gDarkMode ? RGB(45, 45, 45) : RGB(255, 255, 255);
        COLORREF borderColor = gDarkMode ? RGB(100, 100, 100) : RGB(180, 180, 180);
        
        // Cover the entire border area with background first
        HBRUSH bgBrush = CreateSolidBrush(bgColor);
        
        // Top edge (full width, 3 pixels tall)
        RECT topEdge = { rc.left, rc.top, rc.right, rc.top + 3 };
        FillRect(hdc, &topEdge, bgBrush);
        
        // Bottom edge (full width, 3 pixels tall)
        RECT bottomEdge = { rc.left, rc.bottom - 3, rc.right, rc.bottom };
        FillRect(hdc, &bottomEdge, bgBrush);
        
        // Left edge (full height, 3 pixels wide)
        RECT leftEdge = { rc.left, rc.top, rc.left + 3, rc.bottom };
        FillRect(hdc, &leftEdge, bgBrush);
        
        // Right edge up to button (3 pixels wide)
        int btnWidth = GetSystemMetrics(SM_CXVSCROLL);
        RECT rightEdge = { rc.right - btnWidth - 3, rc.top, rc.right - btnWidth, rc.bottom };
        FillRect(hdc, &rightEdge, bgBrush);
        
        DeleteObject(bgBrush);
        
        // Now draw flat 1px border
        HPEN borderPen = CreatePen(PS_SOLID, 1, borderColor);
        HPEN oldPen = (HPEN)SelectObject(hdc, borderPen);
        HBRUSH oldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
        Rectangle(hdc, rc.left, rc.top, rc.right, rc.bottom);
        SelectObject(hdc, oldPen);
        SelectObject(hdc, oldBrush);
        DeleteObject(borderPen);
        
        // Dropdown button area
        RECT btnRc = { rc.right - btnWidth, rc.top + 1, rc.right - 1, rc.bottom - 1 };
        
        // Fill dropdown button with cyan blue
        HBRUSH cyanBrush = CreateSolidBrush(RGB(0, 120, 160));
        FillRect(hdc, &btnRc, cyanBrush);
        DeleteObject(cyanBrush);
        
        // Draw the dropdown arrow
        int arrowSize = 5;
        int cx = btnRc.left + (btnRc.right - btnRc.left) / 2;
        int cy = btnRc.top + (btnRc.bottom - btnRc.top) / 2;
        
        POINT pts[3] = {
            { cx - arrowSize, cy - arrowSize / 2 },
            { cx + arrowSize, cy - arrowSize / 2 },
            { cx, cy + arrowSize / 2 + 1 }
        };
        
        HBRUSH arrowBrush = CreateSolidBrush(RGB(255, 255, 255));
        HPEN arrowPen = CreatePen(PS_SOLID, 1, RGB(255, 255, 255));
        oldBrush = (HBRUSH)SelectObject(hdc, arrowBrush);
        oldPen = (HPEN)SelectObject(hdc, arrowPen);
        Polygon(hdc, pts, 3);
        SelectObject(hdc, oldBrush);
        SelectObject(hdc, oldPen);
        DeleteObject(arrowBrush);
        DeleteObject(arrowPen);
        
        ReleaseDC(hWnd, hdc);
        return result;
    }
    else if (msg == WM_NCPAINT) {
        // Suppress non-client painting to prevent Windows from drawing any border
        return 0;
    }
    else if (msg == WM_ERASEBKGND) {
        // Fill with appropriate background
        HDC hdc = (HDC)w;
        RECT rc;
        GetClientRect(hWnd, &rc);
        HBRUSH bgBrush = CreateSolidBrush(gDarkMode ? RGB(45, 45, 45) : RGB(255, 255, 255));
        FillRect(hdc, &rc, bgBrush);
        DeleteObject(bgBrush);
        return 1;
    }
    return CallWindowProcW(gOriginalComboProc, hWnd, msg, w, l);
}

// Dark mode brushes (created once)
static HBRUSH gDarkBgBrush = nullptr;
static HBRUSH gDarkEditBrush = nullptr;
static HBRUSH gLightBgBrush = nullptr;  // White background for light mode

// Header subclass for dark mode
static WNDPROC gOriginalHeaderProc = nullptr;

// Status bar subclass for dark mode
static WNDPROC gOriginalStatusBarProc = nullptr;

// Tab control subclass for dark mode
static WNDPROC gOriginalTabProc = nullptr;

// Filter state
static std::wstring gFilterStore = L"";      // empty = all
static int gFilterArt = 0;                    // 0=all, 1=has art, 2=missing art
static std::wstring gFilterSearch = L"";

// Cache titles by stable key (steam:123, gog:123, etc)
static std::unordered_map<std::wstring, std::wstring> gTitleCache;

// Cache image indices by filepath (for current icon size only)
static std::unordered_map<std::wstring, int> gImageIndexCache;

// Ubisoft id->name map (loaded once from GitHub)
static bool gUbiMapLoaded = false;
static std::unordered_map<std::wstring, std::wstring> gUbiIdToName;

// Last displayed items (for instant redraw on size/layout change)
struct GameEntry {
    std::wstring store;
    std::wstring idStr;
    std::wstring filePath;         // actual png path if exists
    std::wstring fileName;         // actual png filename if exists
    std::wstring expectedFileName; // expected filename derived from manifest id
    std::wstring title;            // resolved title if possible
    std::wstring customTitle;      // user-set custom title (optional)
    bool hasArt = false;
    // Helper: get display title (custom if set, else resolved)
    std::wstring DisplayTitle() const { return !customTitle.empty() ? customTitle : title; }
};
static std::vector<GameEntry> gLastItems;
static std::vector<GameEntry> gFilteredItems;  // Items after filtering

// Unresolved filter state and control
static HWND gChkUnresolved = nullptr;
static bool gFilterUnresolved = false;

// -------------------- Config --------------------

struct AppConfig {
    std::wstring steamGridDbKey;       // can be empty
    std::wstring thirdPartyBasePath;   // optional override; if empty use default
    bool darkMode = true;              // dark mode by default
    int parallelDownloads = 0;         // 0 = auto (use CPU thread count), or 2,4,6,8...
    int minImageSize = 200;            // minimum image width/height in pixels (filter out small/blurry images)
};
static AppConfig gConfig;

// Config window control IDs
static const int IDC_CFG_KEY_LABEL   = 3001;
static const int IDC_CFG_KEY_EDIT    = 3002;
static const int IDC_CFG_PATH_LABEL  = 3003;
static const int IDC_CFG_PATH_EDIT   = 3004;
static const int IDC_CFG_BROWSE_BTN  = 3005;
static const int IDC_CFG_OK_BTN      = 3006;
static const int IDC_CFG_CANCEL_BTN  = 3007;
static const int IDC_CFG_THEME_LABEL = 3008;
static const int IDC_CFG_CMB_THEME   = 3009;
static const int IDC_CFG_PARALLEL_LABEL = 3010;
static const int IDC_CFG_CMB_PARALLEL   = 3011;
static const int IDC_CFG_CLEAR_CACHE_BTN = 3012;
static const int IDC_CFG_MINSIZE_LABEL = 3013;
static const int IDC_CFG_MINSIZE_EDIT  = 3014;
static const int IDC_CFG_RESTORE_ART_BTN = 3015;

// Art Update window control IDs
static const int IDC_ART_LIST        = 4001;
static const int IDC_ART_APPLY_BTN   = 4002;
static const int IDC_ART_CANCEL_BTN  = 4003;
static const int IDC_ART_STATUS      = 4004;
static const int IDC_ART_RESTORE_BTN = 4005;
static const int IDC_ART_CURRENT_IMG = 4006;
static const int IDC_ART_INFO_LABEL  = 4007;
static const int IDC_ART_DEBUG       = 4008;
static const int IDC_ART_PREVIEW_IMG = 4009;
static const int IDC_ART_PREVIEW_LABEL = 4010;
static const int IDC_ART_TAB         = 4011;
static const int IDC_ART_DEBUG_LABEL = 4012;
static const int IDC_ART_CONFIG_BTN  = 4013;
static const int IDC_ART_CLEAR_CACHE_BTN = 4014;
static const int IDC_ART_WEB_RES_COMBO = 4015;
static const int IDC_ART_WEB_PREV_BTN = 4016;
static const int IDC_ART_WEB_NEXT_BTN = 4017;

static HWND gCfgWnd = nullptr;
static HWND gCfgKeyEdit = nullptr;
static HWND gCfgPathEdit = nullptr;
static HWND gCfgTheme = nullptr;
static HWND gCfgParallel = nullptr;
static HWND gCfgParent = nullptr;
static HWND gCfgMinSizeEdit = nullptr;

// Art Update window globals
static HWND gArtWnd = nullptr;
static HWND gArtList = nullptr;
static HWND gArtParent = nullptr;
static HWND gArtStatus = nullptr;
static HWND gArtApiWarning = nullptr;
static HWND gArtCurrentImg = nullptr;
static HWND gArtPreviewImg = nullptr;
static HWND gArtInfoLabel = nullptr;
static HWND gArtDebug = nullptr;
static HWND gArtTab = nullptr;
static HWND gArtConfigBtn = nullptr;
static GameEntry gArtCurrentEntry;
static HIMAGELIST gArtImgList = nullptr;
static HBITMAP gArtCurrentBitmap = nullptr;
static HBITMAP gArtPreviewBitmap = nullptr;
static int gArtSelectedIndex = -1;
static std::wstring gArtGameId;  // Cached game ID for tab switching

// Track if a custom title was changed in the manual art window
static bool sCustomTitleChanged = false;

// Tab indices for art types
enum ArtType { ART_WEB = 0, ART_GRIDS = 1, ART_HEROES = 2, ART_LOGOS = 3, ART_ICONS = 4 };
static int gArtCurrentTab = ART_WEB;

// Web search state
static HWND gArtWebResolutionCombo = nullptr;  // Resolution dropdown for web search
static HWND gArtWebSizeLabel = nullptr;        // "Image Size:" label
static HWND gArtWebPrevBtn = nullptr;          // "< Prev" button
static HWND gArtWebNextBtn = nullptr;          // "Next >" button
static int gArtWebPageIndex = 0;               // Current page (0, 1, 2, ...)
static bool gArtWebNoMoreResults = false;      // True when no more results available
static std::wstring gArtWebResolution = L"1024x1024";  // Current resolution selection

// Loading state flags
static bool gArtIsLoading = false;      // True while loading thumbnails
static bool gArtCancelLoading = false;  // Set to true to cancel current loading
static int gArtLoadGeneration = 0;      // Incremented on each new load to detect stale loads

// Splitter for preview panel
static int gArtSplitterPos = 200;  // Width of preview panel (from right edge)
static bool gArtSplitterDragging = false;
static const int SPLITTER_WIDTH = 6;

struct SteamGridDbImage {
    std::wstring id;
    std::wstring url;
    std::wstring thumb;
    std::wstring author;
    int width = 0;
    int height = 0;
};
static std::vector<SteamGridDbImage> gArtImages;

// Filter images by minimum size (removes small/blurry images)
static std::vector<SteamGridDbImage> FilterImagesBySize(const std::vector<SteamGridDbImage>& images, int minSize) {
    if (minSize <= 0) return images;  // No filtering if minSize is 0 or negative
    
    std::vector<SteamGridDbImage> filtered;
    for (const auto& img : images) {
        // Keep images where both width and height are >= minSize
        // Also keep images with unknown dimensions (width=0 or height=0)
        if ((img.width == 0 && img.height == 0) || 
            (img.width >= minSize && img.height >= minSize)) {
            filtered.push_back(img);
        }
    }
    return filtered;
}

// ------------------------------------------------------------
// Utilities
// ------------------------------------------------------------

static std::wstring GetKnownFolderPath(REFKNOWNFOLDERID folderId) {
    PWSTR p = nullptr;
    std::wstring out;
    if (SUCCEEDED(SHGetKnownFolderPath(folderId, 0, nullptr, &p)) && p) {
        out = p;
        CoTaskMemFree(p);
    }
    return out;
}

static std::wstring GetDefaultThirdPartyLibrariesPath() {
    std::wstring local = GetKnownFolderPath(FOLDERID_LocalAppData);
    return local + L"\\Packages\\Microsoft.GamingApp_8wekyb3d8bbwe\\LocalState\\ThirdPartyLibraries";
}

static std::wstring GetThirdPartyLibrariesPath() {
    // If user provided an override and it exists, prefer it.
    if (!gConfig.thirdPartyBasePath.empty()) {
        fs::path p = gConfig.thirdPartyBasePath;
        if (fs::exists(p) && fs::is_directory(p)) {
            return p.wstring();
        }
    }
    return GetDefaultThirdPartyLibrariesPath();
}

// Get the profiles folder path
static std::wstring GetProfilesPath() {
    wchar_t exePath[MAX_PATH]{};
    GetModuleFileNameW(nullptr, exePath, MAX_PATH);
    fs::path exeDir = fs::path(exePath).parent_path();
    return (exeDir / L"profiles").wstring();
}

// Profile information structure
struct ProfileInfo {
    std::wstring fullName;      // Full folder name: 20251223143052-MySetup
    std::wstring baseName;      // Base name: MySetup
    std::wstring dateTimeStr;   // Date time: 20251223143052
    int artCount;               // Number of art files in profile
    std::wstring displayName;   // Display: MySetup
    SYSTEMTIME creationTime;    // Creation time
};

// Count art files in a profile
static int CountArtFiles(const std::wstring& profilePath) {
    int count = 0;
    if (!fs::exists(profilePath)) return 0;
    
    std::vector<std::wstring> stores = {L"steam", L"epic", L"gog", L"ubi"};
    for (const auto& store : stores) {
        fs::path storePath = fs::path(profilePath) / store;
        if (fs::exists(storePath) && fs::is_directory(storePath)) {
            try {
                for (const auto& entry : fs::directory_iterator(storePath)) {
                    if (entry.is_regular_file()) {
                        count++;
                    }
                }
            } catch (...) {}
        }
    }
    return count;
}

// Parse profile folder name
static bool ParseProfileName(const std::wstring& folderName, ProfileInfo& info) {
    // Expected format: YYYYMMDDHHMMSS-Name
    std::wregex pattern(L"(\\d{14})-(.+)");
    std::wsmatch matches;
    if (!std::regex_match(folderName, matches, pattern)) {
        return false;
    }
    
    info.fullName = folderName;
    info.dateTimeStr = matches[1].str();
    info.baseName = matches[2].str();
    info.displayName = info.baseName;
    info.artCount = 0;  // Will be set when needed
    
    // Parse datetime: YYYYMMDDHHMMSS
    if (info.dateTimeStr.length() == 14) {
        ZeroMemory(&info.creationTime, sizeof(SYSTEMTIME));
        info.creationTime.wYear = _wtoi(info.dateTimeStr.substr(0, 4).c_str());
        info.creationTime.wMonth = _wtoi(info.dateTimeStr.substr(4, 2).c_str());
        info.creationTime.wDay = _wtoi(info.dateTimeStr.substr(6, 2).c_str());
        info.creationTime.wHour = _wtoi(info.dateTimeStr.substr(8, 2).c_str());
        info.creationTime.wMinute = _wtoi(info.dateTimeStr.substr(10, 2).c_str());
        info.creationTime.wSecond = _wtoi(info.dateTimeStr.substr(12, 2).c_str());
    }
    
    return true;
}

// Get all profiles
static std::vector<ProfileInfo> GetAllProfiles() {
    std::vector<ProfileInfo> profiles;
    std::wstring profilesPath = GetProfilesPath();
    
    if (!fs::exists(profilesPath)) return profiles;
    
    for (const auto& entry : fs::directory_iterator(profilesPath)) {
        if (!entry.is_directory()) continue;
        
        ProfileInfo info;
        if (ParseProfileName(entry.path().filename().wstring(), info)) {
            info.artCount = CountArtFiles(entry.path().wstring());
            profiles.push_back(info);
        }
    }
    
    // Sort by creation time (newest first)
    std::sort(profiles.begin(), profiles.end(), [](const ProfileInfo& a, const ProfileInfo& b) {
        return a.dateTimeStr > b.dateTimeStr;
    });
    
    return profiles;
}

// Create a new profile
static bool CreateProfile(const std::wstring& profileName) {
    if (profileName.empty()) return false;
    
    // Get current timestamp
    SYSTEMTIME st;
    GetLocalTime(&st);
    wchar_t timestamp[32];
    swprintf(timestamp, 32, L"%04d%02d%02d%02d%02d%02d",
        st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    
    // Create folder name
    std::wstring folderName = std::wstring(timestamp) + L"-" + profileName;
    std::wstring profilePath = GetProfilesPath();
    fs::path fullPath = fs::path(profilePath) / folderName;
    
    // Create directory
    if (!fs::exists(profilePath)) {
        fs::create_directories(profilePath);
    }
    fs::create_directories(fullPath);
    
    // Copy library folders
    std::wstring librariesPath = GetThirdPartyLibrariesPath();
    if (!fs::exists(librariesPath)) return false;
    
    std::vector<std::wstring> stores = {L"steam", L"epic", L"gog", L"ubi"};
    for (const auto& store : stores) {
        fs::path srcPath = fs::path(librariesPath) / store;
        if (fs::exists(srcPath) && fs::is_directory(srcPath)) {
            fs::path destPath = fullPath / store;
            fs::create_directories(destPath);
            
            // Copy files except manifest files
            try {
                for (const auto& entry : fs::directory_iterator(srcPath)) {
                    std::wstring filename = entry.path().filename().wstring();
                    // Skip manifest files
                    if (filename.find(L".manifest") != std::wstring::npos) {
                        continue;
                    }
                    fs::path dest = destPath / filename;
                    if (entry.is_regular_file()) {
                        fs::copy_file(entry.path(), dest, fs::copy_options::overwrite_existing);
                    }
                }
            } catch (...) {}
        }
    }
    
    return true;
}

// Restore a profile
static bool RestoreProfile(const ProfileInfo& profile) {
    std::wstring profilesPath = GetProfilesPath();
    fs::path profilePath = fs::path(profilesPath) / profile.fullName;
    
    if (!fs::exists(profilePath)) return false;
    
    std::wstring librariesPath = GetThirdPartyLibrariesPath();
    if (librariesPath.empty()) return false;
    
    // Copy each store folder back
    std::vector<std::wstring> stores = {L"steam", L"epic", L"gog", L"ubi"};
    for (const auto& store : stores) {
        fs::path srcPath = profilePath / store;
        if (fs::exists(srcPath) && fs::is_directory(srcPath)) {
            fs::path destPath = fs::path(librariesPath) / store;
            try {
                // Only delete non-manifest files from destination
                if (fs::exists(destPath)) {
                    for (const auto& entry : fs::directory_iterator(destPath)) {
                        if (entry.is_regular_file()) {
                            std::wstring filename = entry.path().filename().wstring();
                            // Keep manifest files, delete everything else
                            if (filename.find(L".manifest") == std::wstring::npos) {
                                fs::remove(entry.path());
                            }
                        }
                    }
                }
                
                // Create destination if it doesn't exist
                if (!fs::exists(destPath)) {
                    fs::create_directories(destPath);
                }
                
                // Copy art files from profile
                for (const auto& entry : fs::directory_iterator(srcPath)) {
                    if (entry.is_regular_file()) {
                        fs::path dest = destPath / entry.path().filename();
                        fs::copy_file(entry.path(), dest, fs::copy_options::overwrite_existing);
                    }
                }
            } catch (...) {
                return false;
            }
        }
    }
    
    return true;
}

// Delete a profile
static bool DeleteProfile(const ProfileInfo& profile) {
    std::wstring profilesPath = GetProfilesPath();
    fs::path profilePath = fs::path(profilesPath) / profile.fullName;
    
    if (!fs::exists(profilePath)) return false;
    
    try {
        fs::remove_all(profilePath);
        return true;
    } catch (...) {
        return false;
    }
}

// Get the backup folder path for original artwork
static std::wstring GetBackupArtPath() {
    wchar_t exePath[MAX_PATH]{};
    GetModuleFileNameW(nullptr, exePath, MAX_PATH);
    fs::path exeDir = fs::path(exePath).parent_path();
    return (exeDir / L"backupArt").wstring();
}

// Get the backup file path for a specific game's art
static std::wstring GetBackupFilePath(const std::wstring& store, const std::wstring& fileName) {
    fs::path backupDir = fs::path(GetBackupArtPath()) / store;
    return (backupDir / fileName).wstring();
}

// Check if a backup exists for a game's art
static bool HasBackup(const std::wstring& store, const std::wstring& fileName) {
    if (fileName.empty()) return false;
    std::wstring backupPath = GetBackupFilePath(store, fileName);
    return fs::exists(backupPath);
}

// Create a backup of the original art file
static bool BackupOriginalArt(const std::wstring& store, const std::wstring& originalPath, const std::wstring& fileName) {
    if (originalPath.empty() || fileName.empty()) return false;
    if (!fs::exists(originalPath)) return false;

    fs::path backupDir = fs::path(GetBackupArtPath()) / store;
    if (!fs::exists(backupDir)) {
        fs::create_directories(backupDir);
    }

    fs::path backupPath = backupDir / fileName;

    // Only backup if we don't already have one (preserve the original)
    if (!fs::exists(backupPath)) {
        try {
            fs::copy_file(originalPath, backupPath, fs::copy_options::overwrite_existing);
            return true;
        } catch (...) {
            return false;
        }
    }
    return true; // Already have a backup
}

// Restore all backed up art files and remove art that didn't exist before
static int RestoreAllOriginalArt() {
    int restoredCount = 0;
    std::wstring backupBasePath = GetBackupArtPath();
    if (!fs::exists(backupBasePath)) return 0;
    
    // Iterate through store folders (steam, epic, gog, ubi)
    for (const auto& storeEntry : fs::directory_iterator(backupBasePath)) {
        if (!storeEntry.is_directory()) continue;
        std::wstring storeName = storeEntry.path().filename().wstring();
        
        // Iterate through backup files in this store folder
        for (const auto& backupFile : fs::directory_iterator(storeEntry.path())) {
            if (!backupFile.is_regular_file()) continue;
            
            std::wstring backupPath = backupFile.path().wstring();
            std::wstring fileName = backupFile.path().filename().wstring();
            
            // Find the original art path by searching current games
            bool restored = false;
            for (const auto& game : gLastItems) {
                if (game.store == storeName && game.fileName == fileName && !game.filePath.empty()) {
                    // Remove read-only attribute if present
                    if (fs::exists(game.filePath)) {
                        DWORD attrs = GetFileAttributesW(game.filePath.c_str());
                        if (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_READONLY)) {
                            SetFileAttributesW(game.filePath.c_str(), attrs & ~FILE_ATTRIBUTE_READONLY);
                        }
                    }
                    
                    // Copy backup over current file
                    try {
                        fs::copy_file(backupPath, game.filePath, fs::copy_options::overwrite_existing);
                        restoredCount++;
                        restored = true;
                    } catch (...) {}
                    break;
                }
            }
        }
    }
    
    // Remove art files that don't have backups (newly created)
    for (const auto& game : gLastItems) {
        if (game.hasArt && !game.filePath.empty() && !game.fileName.empty()) {
            if (!HasBackup(game.store, game.fileName) && fs::exists(game.filePath)) {
                // This art was newly created (no backup exists), remove it
                DWORD attrs = GetFileAttributesW(game.filePath.c_str());
                if (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_READONLY)) {
                    SetFileAttributesW(game.filePath.c_str(), attrs & ~FILE_ATTRIBUTE_READONLY);
                }
                try {
                    fs::remove(game.filePath);
                    restoredCount++;
                } catch (...) {}
            }
        }
    }
    
    return restoredCount;
}

// Set a file as read-only to prevent Xbox app from modifying it
static bool SetFileReadOnly(const std::wstring& path) {
    if (path.empty() || !fs::exists(path)) return false;
    DWORD attrs = GetFileAttributesW(path.c_str());
    if (attrs == INVALID_FILE_ATTRIBUTES) return false;
    return SetFileAttributesW(path.c_str(), attrs | FILE_ATTRIBUTE_READONLY) != 0;
}

// Remove read-only attribute from a file
static bool RemoveFileReadOnly(const std::wstring& path) {
    if (path.empty() || !fs::exists(path)) return true;  // No file = success
    DWORD attrs = GetFileAttributesW(path.c_str());
    if (attrs == INVALID_FILE_ATTRIBUTES) return true;
    if (!(attrs & FILE_ATTRIBUTE_READONLY)) return true;  // Already writable
    return SetFileAttributesW(path.c_str(), attrs & ~FILE_ATTRIBUTE_READONLY) != 0;
}

// Restore art from backup
static bool RestoreArtFromBackup(const std::wstring& store, const std::wstring& fileName) {
    if (fileName.empty()) return false;

    std::wstring backupPath = GetBackupFilePath(store, fileName);
    if (!fs::exists(backupPath)) return false;

    std::wstring basePath = GetThirdPartyLibrariesPath();
    fs::path destPath = fs::path(basePath) / store / fileName;

    try {
        // Remove read-only attribute before overwriting
        RemoveFileReadOnly(destPath.wstring());
        
        fs::copy_file(backupPath, destPath, fs::copy_options::overwrite_existing);
        
        // Set as read-only to prevent Xbox app from modifying
        SetFileReadOnly(destPath.wstring());
        return true;
    } catch (...) {
        return false;
    }
}

static bool IsAllDigits(const std::wstring& s) {
    if (s.empty()) return false;
    for (wchar_t c : s) if (c < L'0' || c > L'9') return false;
    return true;
}

static std::string ReadAllTextA(const std::wstring& path) {
    HANDLE h = CreateFileW(path.c_str(), GENERIC_READ,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);

    if (h == INVALID_HANDLE_VALUE) return {};

    LARGE_INTEGER sz{};
    if (!GetFileSizeEx(h, &sz) || sz.QuadPart <= 0 || sz.QuadPart > 50 * 1024 * 1024) {
        CloseHandle(h);
        return {};
    }

    std::string data;
    data.resize((size_t)sz.QuadPart);

    DWORD read = 0;
    BOOL ok = ReadFile(h, data.data(), (DWORD)data.size(), &read, nullptr);
    CloseHandle(h);

    if (!ok) return {};
    data.resize(read);
    return data;
}

static bool WriteAllTextA(const std::wstring& path, const std::string& data) {
    HANDLE h = CreateFileW(path.c_str(), GENERIC_WRITE,
        FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return false;

    DWORD written = 0;
    BOOL ok = WriteFile(h, data.data(), (DWORD)data.size(), &written, nullptr);
    CloseHandle(h);
    return ok && written == data.size();
}

// Apply dark mode to a window (Windows 10 1809+)
static void ApplyDarkModeToWindow(HWND hWnd) {
    LoadDarkModeApis();
    
    BOOL useDarkMode = gDarkMode ? TRUE : FALSE;
    DwmSetWindowAttribute(hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &useDarkMode, sizeof(useDarkMode));
    
    // Also try the pre-release attribute value (19) for older Windows 10 versions
    DwmSetWindowAttribute(hWnd, 19, &useDarkMode, sizeof(useDarkMode));
    
    // Use undocumented API for window-level dark mode
    if (AllowDarkModeForWindow) {
        AllowDarkModeForWindow(hWnd, gDarkMode);
    }
    
    // Create brushes if needed
    if (!gDarkBgBrush) gDarkBgBrush = CreateSolidBrush(RGB(32, 32, 32));
    if (!gDarkEditBrush) gDarkEditBrush = CreateSolidBrush(RGB(45, 45, 45));
    if (!gLightBgBrush) gLightBgBrush = CreateSolidBrush(RGB(255, 255, 255));
}

// Apply dark mode theme to common controls
static void ApplyDarkModeToControls(HWND hWnd) {
    LoadDarkModeApis();
    
    if (gDarkMode) {
        SetWindowTheme(hWnd, L"DarkMode_Explorer", nullptr);
    } else {
        SetWindowTheme(hWnd, L"Explorer", nullptr);
    }
    
    // Apply to all child windows
    EnumChildWindows(hWnd, [](HWND child, LPARAM) -> BOOL {
        wchar_t className[64]{};
        GetClassNameW(child, className, _countof(className));
        
        // Allow dark mode for this window
        if (AllowDarkModeForWindow) {
            AllowDarkModeForWindow(child, gDarkMode);
        }
        
        if (gDarkMode) {
            // Different themes for different control types
            if (wcscmp(className, L"SysListView32") == 0) {
                SetWindowTheme(child, L"DarkMode_Explorer", nullptr);
                ListView_SetBkColor(child, RGB(32, 32, 32));
                ListView_SetTextBkColor(child, RGB(32, 32, 32));
                ListView_SetTextColor(child, RGB(255, 255, 255));
                // Apply immersive dark mode for scrollbars
                BOOL useDarkMode = TRUE;
                DwmSetWindowAttribute(child, DWMWA_USE_IMMERSIVE_DARK_MODE, &useDarkMode, sizeof(useDarkMode));
            } else if (wcscmp(className, L"SysTabControl32") == 0) {
                SetWindowTheme(child, L"DarkMode_Explorer", nullptr);
            } else if (wcscmp(className, L"msctls_statusbar32") == 0) {
                // Status bar dark mode
                SendMessageW(child, SB_SETBKCOLOR, 0, (LPARAM)RGB(32, 32, 32));
            } else if (wcscmp(className, L"Edit") == 0) {
                SetWindowTheme(child, L"DarkMode_Explorer", nullptr);
                // Apply immersive dark mode for scrollbars
                BOOL useDarkMode = TRUE;
                DwmSetWindowAttribute(child, DWMWA_USE_IMMERSIVE_DARK_MODE, &useDarkMode, sizeof(useDarkMode));
            } else if (wcscmp(className, L"ComboBox") == 0) {
                // Remove WS_EX_CLIENTEDGE to eliminate bevel border
                LONG_PTR exStyle = GetWindowLongPtrW(child, GWL_EXSTYLE);
                if (exStyle & WS_EX_CLIENTEDGE) {
                    SetWindowLongPtrW(child, GWL_EXSTYLE, exStyle & ~WS_EX_CLIENTEDGE);
                    SetWindowPos(child, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
                }
                // Disable theming completely for flat look
                SetWindowTheme(child, L"", L"");
                // Subclass for cyan dropdown button and flat border
                WNDPROC oldProc = (WNDPROC)GetWindowLongPtrW(child, GWLP_WNDPROC);
                if (oldProc != ComboSubclassProc) {
                    if (!gOriginalComboProc) gOriginalComboProc = oldProc;
                    SetWindowLongPtrW(child, GWLP_WNDPROC, (LONG_PTR)ComboSubclassProc);
                }
            } else if (wcscmp(className, L"Button") == 0) {
                SetWindowTheme(child, L"DarkMode_Explorer", nullptr);
            } else if (wcscmp(className, L"Static") == 0) {
                // No theme for static - we handle via WM_CTLCOLORSTATIC
            } else {
                SetWindowTheme(child, L"DarkMode_Explorer", nullptr);
            }
        } else {
            SetWindowTheme(child, nullptr, nullptr);
            if (wcscmp(className, L"SysListView32") == 0) {
                ListView_SetBkColor(child, RGB(255, 255, 255));
                ListView_SetTextBkColor(child, RGB(255, 255, 255));
                ListView_SetTextColor(child, RGB(0, 0, 0));
            } else if (wcscmp(className, L"msctls_statusbar32") == 0) {
                SendMessageW(child, SB_SETBKCOLOR, 0, (LPARAM)CLR_DEFAULT);
            } else if (wcscmp(className, L"ComboBox") == 0) {
                // Remove WS_EX_CLIENTEDGE to eliminate bevel border in light mode too
                LONG_PTR exStyle = GetWindowLongPtrW(child, GWL_EXSTYLE);
                if (exStyle & WS_EX_CLIENTEDGE) {
                    SetWindowLongPtrW(child, GWL_EXSTYLE, exStyle & ~WS_EX_CLIENTEDGE);
                    SetWindowPos(child, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
                }
                // Disable theming completely for flat look
                SetWindowTheme(child, L"", L"");
                // Subclass for cyan dropdown button and flat border
                WNDPROC oldProc = (WNDPROC)GetWindowLongPtrW(child, GWLP_WNDPROC);
                if (oldProc != ComboSubclassProc) {
                    if (!gOriginalComboProc) gOriginalComboProc = oldProc;
                    SetWindowLongPtrW(child, GWLP_WNDPROC, (LONG_PTR)ComboSubclassProc);
                }
            }
        }
        
        // Force redraw
        InvalidateRect(child, nullptr, TRUE);
        
        return TRUE;
    }, 0);
    
    // Flush menu themes to apply dark mode to context menus
    if (FlushMenuThemes) {
        FlushMenuThemes();
    }
}

// Dark mode aware message box using TaskDialog
// TaskDialog respects system dark mode settings better than MessageBox
// Custom message box dialog
struct MessageBoxData {
    std::wstring message;
    std::wstring title;
    UINT type;
    int result;
};

static int DarkMessageBox(HWND hParent, const std::wstring& message, const std::wstring& title, UINT type) {
    static bool msgBoxClassRegistered = false;
    static const wchar_t* MSGBOX_WND_CLASS = L"DarkMessageBoxClass";
    
    MessageBoxData* pData = new MessageBoxData{ message, title, type, IDOK };
    
    if (!msgBoxClassRegistered) {
        WNDCLASSEXW wc = { sizeof(WNDCLASSEXW) };
        wc.lpfnWndProc = [](HWND hWnd, UINT msg, WPARAM w, LPARAM l) -> LRESULT {
            MessageBoxData* pData = (MessageBoxData*)GetWindowLongPtrW(hWnd, GWLP_USERDATA);
            
            switch (msg) {
            case WM_CREATE: {
                CREATESTRUCTW* cs = (CREATESTRUCTW*)l;
                pData = (MessageBoxData*)cs->lpCreateParams;
                SetWindowLongPtrW(hWnd, GWLP_USERDATA, (LONG_PTR)pData);
                
                // Get client area
                RECT rc;
                GetClientRect(hWnd, &rc);
                int width = rc.right - rc.left;
                int height = rc.bottom - rc.top;
                
                // Create icon
                int iconY = 20;
                LPCWSTR iconType = IDI_INFORMATION;
                if (pData->type & MB_ICONERROR) iconType = IDI_ERROR;
                else if (pData->type & MB_ICONWARNING) iconType = IDI_WARNING;
                else if (pData->type & MB_ICONQUESTION) iconType = IDI_QUESTION;
                
                HICON hIcon = LoadIconW(nullptr, iconType);
                HWND hIconCtrl = CreateWindowExW(0, L"STATIC", nullptr,
                    WS_CHILD | WS_VISIBLE | SS_ICON,
                    20, iconY, 32, 32, hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                SendMessageW(hIconCtrl, STM_SETICON, (WPARAM)hIcon, 0);
                
                // Create message text
                HWND hText = CreateWindowExW(0, L"STATIC", pData->message.c_str(),
                    WS_CHILD | WS_VISIBLE | SS_LEFT,
                    65, 20, width - 85, height - 80, hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                SendMessageW(hText, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);
                
                // Create buttons
                int btnWidth = 80;
                int btnHeight = 28;
                int btnY = height - btnHeight - 15;
                
                if (pData->type & MB_YESNO) {
                    int totalWidth = btnWidth * 2 + 10;
                    int startX = (width - totalWidth) / 2;
                    CreateWindowExW(0, L"BUTTON", L"Yes", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                        startX, btnY, btnWidth, btnHeight, hWnd, (HMENU)IDYES, GetModuleHandleW(nullptr), nullptr);
                    CreateWindowExW(0, L"BUTTON", L"No", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                        startX + btnWidth + 10, btnY, btnWidth, btnHeight, hWnd, (HMENU)IDNO, GetModuleHandleW(nullptr), nullptr);
                } else if (pData->type & MB_OKCANCEL) {
                    int totalWidth = btnWidth * 2 + 10;
                    int startX = (width - totalWidth) / 2;
                    CreateWindowExW(0, L"BUTTON", L"OK", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                        startX, btnY, btnWidth, btnHeight, hWnd, (HMENU)IDOK, GetModuleHandleW(nullptr), nullptr);
                    CreateWindowExW(0, L"BUTTON", L"Cancel", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                        startX + btnWidth + 10, btnY, btnWidth, btnHeight, hWnd, (HMENU)IDCANCEL, GetModuleHandleW(nullptr), nullptr);
                } else {
                    // Just OK button
                    int startX = (width - btnWidth) / 2;
                    CreateWindowExW(0, L"BUTTON", L"OK", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW,
                        startX, btnY, btnWidth, btnHeight, hWnd, (HMENU)IDOK, GetModuleHandleW(nullptr), nullptr);
                }
                
                // Apply dark mode to window
                if (gDarkMode) {
                    BOOL dark = TRUE;
                    DwmSetWindowAttribute(hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &dark, sizeof(dark));
                } else {
                    BOOL dark = FALSE;
                    DwmSetWindowAttribute(hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &dark, sizeof(dark));
                }
                return 0;
            }
            case WM_DRAWITEM: {
                DRAWITEMSTRUCT* dis = (DRAWITEMSTRUCT*)l;
                if (dis->CtlType == ODT_BUTTON) {
                    DrawDarkButton(dis);
                    return TRUE;
                }
                break;
            }
            case WM_CTLCOLOREDIT:
            case WM_CTLCOLORSTATIC:
                if (gDarkMode) {
                    HDC hdc = (HDC)w;
                    SetTextColor(hdc, RGB(255, 255, 255));
                    SetBkColor(hdc, RGB(32, 32, 32));
                    if (!gDarkBgBrush) gDarkBgBrush = CreateSolidBrush(RGB(32, 32, 32));
                    return (LRESULT)gDarkBgBrush;
                } else {
                    HDC hdc = (HDC)w;
                    SetTextColor(hdc, RGB(0, 0, 0));
                    SetBkColor(hdc, GetSysColor(COLOR_WINDOW));
                    return (LRESULT)GetSysColorBrush(COLOR_WINDOW);
                }
                break;
            case WM_CTLCOLORBTN:
                if (gDarkMode) {
                    if (!gDarkBgBrush) gDarkBgBrush = CreateSolidBrush(RGB(32, 32, 32));
                    return (LRESULT)gDarkBgBrush;
                } else {
                    return (LRESULT)GetSysColorBrush(COLOR_BTNFACE);
                }
                break;
            case WM_ERASEBKGND:
                {
                    HDC hdc = (HDC)w;
                    RECT rc;
                    GetClientRect(hWnd, &rc);
                    if (gDarkMode) {
                        if (!gDarkBgBrush) gDarkBgBrush = CreateSolidBrush(RGB(32, 32, 32));
                        FillRect(hdc, &rc, gDarkBgBrush);
                    } else {
                        FillRect(hdc, &rc, GetSysColorBrush(COLOR_WINDOW));
                    }
                    return 1;
                }
                break;
            case WM_COMMAND:
                if (HIWORD(w) == BN_CLICKED) {
                    int id = LOWORD(w);
                    if (id == IDOK || id == IDYES || id == IDNO || id == IDCANCEL) {
                        if (pData) pData->result = id;
                        DestroyWindow(hWnd);
                        return 0;
                    }
                }
                break;
            case WM_CLOSE:
                if (pData) pData->result = IDCANCEL;
                DestroyWindow(hWnd);
                return 0;
            }
            return DefWindowProcW(hWnd, msg, w, l);
        };
        wc.hInstance = GetModuleHandleW(nullptr);
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hbrBackground = gDarkMode ? gDarkBgBrush : (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = MSGBOX_WND_CLASS;
        RegisterClassExW(&wc);
        msgBoxClassRegistered = true;
    }
    
    // Calculate size based on message length
    int dlgWidth = 400;
    int dlgHeight = 160;
    
    RECT workArea;
    SystemParametersInfo(SPI_GETWORKAREA, 0, &workArea, 0);
    int x = workArea.left + ((workArea.right - workArea.left) - dlgWidth) / 2;
    int y = workArea.top + ((workArea.bottom - workArea.top) - dlgHeight) / 2;
    
    HWND hMsgBox = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_TOPMOST,
        MSGBOX_WND_CLASS, title.c_str(),
        WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        x, y, dlgWidth, dlgHeight,
        hParent, nullptr, GetModuleHandleW(nullptr), pData);
    
    if (!hMsgBox) {
        delete pData;
        return IDCANCEL;
    }
    
    // Disable parent window
    if (hParent) EnableWindow(hParent, FALSE);
    
    // Message loop
    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        if (!IsWindow(hMsgBox)) break;
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    
    // Re-enable parent
    if (hParent) {
        EnableWindow(hParent, TRUE);
        SetForegroundWindow(hParent);
    }
    
    int result = pData->result;
    delete pData;
    return result;
}

static std::wstring Utf8ToWide(const std::string& s) {
    if (s.empty()) return L"";
    int len = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), nullptr, 0);
    if (len <= 0) return L"";
    std::wstring w(len, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), w.data(), len);
    return w;
}

static std::string WideToUtf8(const std::wstring& w) {
    if (w.empty()) return "";
    int len = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), nullptr, 0, nullptr, nullptr);
    if (len <= 0) return "";
    std::string s(len, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), s.data(), len, nullptr, nullptr);
    return s;
}

// Decode JSON Unicode escape sequences like \u0027 -> ' or u0027 -> '
static std::string DecodeJsonUnicodeEscapes(const std::string& s) {
    std::string out;
    out.reserve(s.size());
    for (size_t i = 0; i < s.size(); ++i) {
        // Check for \uXXXX format
        if (s[i] == '\\' && i + 5 < s.size() && s[i + 1] == 'u') {
            // Parse 4 hex digits
            std::string hex = s.substr(i + 2, 4);
            bool valid = true;
            for (char c : hex) {
                if (!isxdigit((unsigned char)c)) { valid = false; break; }
            }
            if (valid) {
                unsigned int codepoint = std::stoul(hex, nullptr, 16);
                if (codepoint < 0x80) {
                    out += (char)codepoint;
                } else if (codepoint < 0x800) {
                    out += (char)(0xC0 | (codepoint >> 6));
                    out += (char)(0x80 | (codepoint & 0x3F));
                } else {
                    out += (char)(0xE0 | (codepoint >> 12));
                    out += (char)(0x80 | ((codepoint >> 6) & 0x3F));
                    out += (char)(0x80 | (codepoint & 0x3F));
                }
                i += 5; // skip \uXXXX
                continue;
            }
        }
        // Check for standalone uXXXX format (without backslash) - sometimes seen in GOG data
        if (s[i] == 'u' && i + 4 < s.size()) {
            std::string hex = s.substr(i + 1, 4);
            bool valid = true;
            for (char c : hex) {
                if (!isxdigit((unsigned char)c)) { valid = false; break; }
            }
            // Only decode if it looks like a Unicode escape (all 4 chars are hex)
            // and the codepoint is in a reasonable range for special chars
            if (valid) {
                unsigned int codepoint = std::stoul(hex, nullptr, 16);
                // Only decode if it's a printable special character range
                if (codepoint >= 0x20 && codepoint < 0x80) {
                    out += (char)codepoint;
                    i += 4; // skip uXXXX
                    continue;
                }
            }
        }
        out += s[i];
    }
    return out;
}

// Decode HTML entities like &#x27; -> ' and &#39; -> ' and &amp; -> &
static std::string DecodeHtmlEntities(const std::string& s) {
    std::string out;
    out.reserve(s.size());
    for (size_t i = 0; i < s.size(); ++i) {
        if (s[i] == '&') {
            // Check for numeric entities &#xNN; or &#NN;
            if (i + 2 < s.size() && s[i + 1] == '#') {
                size_t start = i + 2;
                bool isHex = (start < s.size() && (s[start] == 'x' || s[start] == 'X'));
                if (isHex) start++;
                
                size_t end = start;
                while (end < s.size() && s[end] != ';' && end - start < 8) end++;
                
                if (end < s.size() && s[end] == ';' && end > start) {
                    std::string num = s.substr(start, end - start);
                    bool valid = true;
                    for (char c : num) {
                        if (isHex ? !isxdigit((unsigned char)c) : !isdigit((unsigned char)c)) {
                            valid = false; break;
                        }
                    }
                    if (valid) {
                        unsigned int codepoint = std::stoul(num, nullptr, isHex ? 16 : 10);
                        if (codepoint < 0x80) {
                            out += (char)codepoint;
                        } else if (codepoint < 0x800) {
                            out += (char)(0xC0 | (codepoint >> 6));
                            out += (char)(0x80 | (codepoint & 0x3F));
                        } else {
                            out += (char)(0xE0 | (codepoint >> 12));
                            out += (char)(0x80 | ((codepoint >> 6) & 0x3F));
                            out += (char)(0x80 | (codepoint & 0x3F));
                        }
                        i = end; // skip past ;
                        continue;
                    }
                }
            }
            // Check for named entities
            else if (s.substr(i, 5) == "&amp;") { out += '&'; i += 4; continue; }
            else if (s.substr(i, 6) == "&apos;") { out += '\''; i += 5; continue; }
            else if (s.substr(i, 6) == "&quot;") { out += '"'; i += 5; continue; }
            else if (s.substr(i, 4) == "&lt;") { out += '<'; i += 3; continue; }
            else if (s.substr(i, 4) == "&gt;") { out += '>'; i += 3; continue; }
            else if (s.substr(i, 6) == "&nbsp;") { out += ' '; i += 5; continue; }
        }
        out += s[i];
    }
    return out;
}

// Decode both JSON unicode escapes and HTML entities
static std::string DecodeAllEntities(const std::string& s) {
    return DecodeHtmlEntities(DecodeJsonUnicodeEscapes(s));
}

static std::string TrimA(std::string s) {
    auto issp = [](unsigned char c){ return c==' '||c=='\t'||c=='\r'||c=='\n'; };
    while (!s.empty() && issp((unsigned char)s.front())) s.erase(s.begin());
    while (!s.empty() && issp((unsigned char)s.back())) s.pop_back();
    return s;
}

static std::string JsEscape(const std::string& s) {
    std::string o;
    o.reserve(s.size() + 16);
    for (char c : s) {
        switch (c) {
        case '\\': o += "\\\\"; break;
        case '"':  o += "\\\""; break;
        case '\r': o += "\\r"; break;
        case '\n': o += "\\n"; break;
        case '\t': o += "\\t"; break;
        default:   o += c; break;
        }
    }
    return o;
}

static std::string JsUnescape(const std::string& s) {
    std::string o;
    o.reserve(s.size());
    for (size_t i = 0; i < s.size(); ++i) {
        char c = s[i];
        if (c == '\\' && i + 1 < s.size()) {
            char n = s[++i];
            if (n == 'n') o.push_back('\n');
            else if (n == 'r') o.push_back('\r');
            else if (n == 't') o.push_back('\t');
            else o.push_back(n);
        } else {
            o.push_back(c);
        }
    }
    return o;
}

static std::wstring GetCacheBasePath() {
    wchar_t tmp[MAX_PATH]{};
    DWORD n = GetTempPathW(_countof(tmp), tmp);
    std::wstring base = (n > 0) ? std::wstring(tmp) : L".\\";
    return base + L"XboxAppUpdaterArtCache\\";
}

static std::wstring GetImageCachePath() {
    return GetCacheBasePath() + L"imageCache\\";
}

static std::wstring GetCacheFilePath() {
    return GetCacheBasePath() + L"config.json";
}

// Ensure cache directories exist
static void EnsureCacheDirectories() {
    std::wstring basePath = GetCacheBasePath();
    std::wstring imagePath = GetImageCachePath();
    
    // Create base directory
    if (!fs::exists(basePath)) {
        fs::create_directories(basePath);
    }
    // Create image cache directory
    if (!fs::exists(imagePath)) {
        fs::create_directories(imagePath);
    }
}

// Simple hash function for URL to filename
static std::wstring HashUrlToFilename(const std::wstring& url) {
    // Use a simple hash - FNV-1a
    uint64_t hash = 14695981039346656037ULL;
    for (wchar_t c : url) {
        hash ^= static_cast<uint64_t>(c);
        hash *= 1099511628211ULL;
    }
    
    // Convert to hex string
    wchar_t buf[32];
    swprintf_s(buf, L"%016llx", hash);
    
    // Determine extension from URL
    std::wstring ext = L".png";
    size_t dotPos = url.rfind(L'.');
    if (dotPos != std::wstring::npos) {
        std::wstring urlExt = url.substr(dotPos);
        // Only use common image extensions
        if (urlExt == L".png" || urlExt == L".jpg" || urlExt == L".jpeg" || 
            urlExt == L".webp" || urlExt == L".gif" || urlExt == L".ico") {
            ext = urlExt;
        }
    }
    
    return std::wstring(buf) + ext;
}

// Get cached image path for a URL
static std::wstring GetCachedImagePath(const std::wstring& url) {
    return GetImageCachePath() + HashUrlToFilename(url);
}

// Check if image is cached
static bool IsImageCached(const std::wstring& url) {
    std::wstring cachePath = GetCachedImagePath(url);
    return fs::exists(cachePath);
}

// Load image from cache
static bool LoadImageFromCache(const std::wstring& url, std::vector<unsigned char>& outData) {
    std::wstring cachePath = GetCachedImagePath(url);
    
    if (!fs::exists(cachePath)) return false;
    
    std::ifstream file(cachePath, std::ios::binary);
    if (!file) return false;
    
    file.seekg(0, std::ios::end);
    size_t size = static_cast<size_t>(file.tellg());
    file.seekg(0, std::ios::beg);
    
    outData.resize(size);
    file.read(reinterpret_cast<char*>(outData.data()), size);
    
    return file.good();
}

// Save image to cache
static bool SaveImageToCache(const std::wstring& url, const std::vector<unsigned char>& data) {
    EnsureCacheDirectories();
    
    std::wstring cachePath = GetCachedImagePath(url);
    
    std::ofstream file(cachePath, std::ios::binary);
    if (!file) return false;
    
    file.write(reinterpret_cast<const char*>(data.data()), data.size());
    return file.good();
}

// Get path to per-game cache tracking directory
static std::wstring GetGameCacheDbPath() {
    return GetCacheBasePath() + L"cachedImagesDB\\";
}

// Get path to a specific game's cache file
static std::wstring GetGameCacheFilePath(const std::wstring& gameId) {
    return GetGameCacheDbPath() + gameId + L".txt";
}

// Register a cached file for a specific game
static void RegisterGameCacheFile(const std::wstring& gameId, const std::wstring& cacheFileName) {
    if (gameId.empty() || cacheFileName.empty()) return;
    
    std::wstring dbPath = GetGameCacheDbPath();
    if (!fs::exists(dbPath)) {
        fs::create_directories(dbPath);
    }
    
    std::wstring filePath = GetGameCacheFilePath(gameId);
    
    // Read existing entries to avoid duplicates
    std::vector<std::wstring> entries;
    if (fs::exists(filePath)) {
        std::wifstream inFile(filePath);
        std::wstring line;
        while (std::getline(inFile, line)) {
            if (!line.empty() && line != cacheFileName) {
                entries.push_back(line);
            }
        }
    }
    
    // Add the new entry
    entries.push_back(cacheFileName);
    
    // Write back
    std::wofstream outFile(filePath);
    for (const auto& entry : entries) {
        outFile << entry << L"\n";
    }
}

// Clear all cached images for a specific game
static int ClearGameCache(const std::wstring& gameId) {
    if (gameId.empty()) return 0;
    
    std::wstring cacheFilePath = GetGameCacheFilePath(gameId);
    if (!fs::exists(cacheFilePath)) return 0;
    
    std::wstring imageCachePath = GetImageCachePath();
    int deletedCount = 0;
    
    // Read the list of cached files
    std::wifstream inFile(cacheFilePath);
    std::wstring line;
    while (std::getline(inFile, line)) {
        if (line.empty()) continue;
        
        std::wstring fullPath = imageCachePath + line;
        if (fs::exists(fullPath)) {
            try {
                fs::remove(fullPath);
                deletedCount++;
            } catch (...) {}
        }
    }
    inFile.close();
    
    // Delete the cache tracking file
    try {
        fs::remove(cacheFilePath);
    } catch (...) {}
    
    return deletedCount;
}

// Clear all cached images
static int ClearAllImageCache() {
    std::wstring imageCachePath = GetImageCachePath();
    std::wstring gameDbPath = GetGameCacheDbPath();
    int deletedCount = 0;
    
    // Delete all files in imageCache folder
    if (fs::exists(imageCachePath)) {
        for (const auto& entry : fs::directory_iterator(imageCachePath)) {
            if (entry.is_regular_file()) {
                try {
                    fs::remove(entry.path());
                    deletedCount++;
                } catch (...) {}
            }
        }
    }
    
    // Delete all per-game tracking files
    if (fs::exists(gameDbPath)) {
        for (const auto& entry : fs::directory_iterator(gameDbPath)) {
            if (entry.is_regular_file()) {
                try {
                    fs::remove(entry.path());
                } catch (...) {}
            }
        }
    }
    
    return deletedCount;
}

// Folder picker
static std::wstring BrowseForFolder(HWND owner, const std::wstring& title) {
    BROWSEINFOW bi{};
    bi.hwndOwner = owner;
    bi.lpszTitle = title.c_str();
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE | BIF_EDITBOX;

    PIDLIST_ABSOLUTE pidl = SHBrowseForFolderW(&bi);
    if (!pidl) return L"";

    wchar_t path[MAX_PATH]{};
    std::wstring out;
    if (SHGetPathFromIDListW(pidl, path)) out = path;

    CoTaskMemFree(pidl);
    return out;
}

// ------------------------------------------------------------
// WinHTTP
// ------------------------------------------------------------

static bool HttpGetUtf8(const std::wstring& url, const std::wstring& bearer, std::string& outBody) {
    outBody.clear();

    URL_COMPONENTS uc{};
    uc.dwStructSize = sizeof(uc);
    wchar_t host[256]{};
    wchar_t path[4096]{};
    uc.lpszHostName = host; uc.dwHostNameLength = _countof(host);
    uc.lpszUrlPath  = path; uc.dwUrlPathLength  = _countof(path);

    if (!WinHttpCrackUrl(url.c_str(), 0, 0, &uc)) return false;

    HINTERNET hSession = WinHttpOpen(
        L"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
        WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS, 0);

    if (!hSession) return false;

    DWORD redirectPolicy = WINHTTP_OPTION_REDIRECT_POLICY_ALWAYS;
    WinHttpSetOption(hSession, WINHTTP_OPTION_REDIRECT_POLICY, &redirectPolicy, sizeof(redirectPolicy));

    HINTERNET hConnect = WinHttpConnect(hSession, uc.lpszHostName, uc.nPort, 0);
    if (!hConnect) { WinHttpCloseHandle(hSession); return false; }

    DWORD flags = (uc.nScheme == INTERNET_SCHEME_HTTPS) ? WINHTTP_FLAG_SECURE : 0;
    HINTERNET hRequest = WinHttpOpenRequest(
        hConnect, L"GET", uc.lpszUrlPath,
        nullptr, WINHTTP_NO_REFERER,
        WINHTTP_DEFAULT_ACCEPT_TYPES, flags);

    if (!hRequest) {
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return false;
    }

    std::wstring headers;
    headers += L"Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8\r\n";
    headers += L"Accept-Language: en-US,en;q=0.9\r\n";
    headers += L"Accept-Encoding: identity\r\n"; // avoid gzip/br issues
    headers += L"Connection: close\r\n";
    if (!bearer.empty()) headers += L"Authorization: Bearer " + bearer + L"\r\n";

    BOOL ok = WinHttpSendRequest(
        hRequest,
        headers.c_str(),
        (DWORD)headers.size(),
        WINHTTP_NO_REQUEST_DATA, 0, 0, 0);

    if (ok) ok = WinHttpReceiveResponse(hRequest, nullptr);

    if (!ok) {
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return false;
    }

    DWORD status = 0;
    DWORD statusSize = sizeof(status);
    WinHttpQueryHeaders(hRequest,
        WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
        WINHTTP_HEADER_NAME_BY_INDEX,
        &status, &statusSize, WINHTTP_NO_HEADER_INDEX);

    std::string body;
    for (;;) {
        DWORD avail = 0;
        if (!WinHttpQueryDataAvailable(hRequest, &avail) || avail == 0) break;
        size_t old = body.size();
        body.resize(old + avail);
        DWORD read = 0;
        if (!WinHttpReadData(hRequest, body.data() + old, avail, &read)) break;
        body.resize(old + read);
    }

    WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);

    outBody = std::move(body);
    return (status >= 200 && status < 300);
}

// ------------------------------------------------------------
// Manifest parsing
// ------------------------------------------------------------

static std::vector<std::wstring> ExtractIdsFromManifest(const std::wstring& manifestPath, const std::wstring& prefix) {
    std::vector<std::wstring> ids;
    std::string content = ReadAllTextA(manifestPath);
    if (content.empty()) return ids;

    std::wstring w = Utf8ToWide(content);
    if (w.empty()) return ids;

    // IMPORTANT: use a custom raw-string delimiter here because the pattern contains )"
    std::wregex re(LR"ID("id"\s*:\s*"([^"]+)")ID", std::regex::icase);
    std::wsregex_iterator it(w.begin(), w.end(), re), end;

    for (; it != end; ++it) {
        std::wstring id = (*it)[1].str();
        if (id.rfind(prefix, 0) == 0) ids.push_back(id);
    }
    return ids;
}

static std::wstring ExpectedPngFromManifestId(const std::wstring& store, const std::wstring& manifestId) {
    std::wstring s = manifestId;
    size_t p = s.find(L':');
    if (p != std::wstring::npos) s = s.substr(p + 1);
    for (auto& c : s) if (c == L':') c = L'_';
    return store + L"_" + s + L".png";
}

static bool TrySplitStoreAndRest(const std::wstring& idStr, std::wstring& outStore, std::wstring& outRest) {
    outStore.clear();
    outRest.clear();
    size_t p = idStr.find(L':');
    if (p == std::wstring::npos) return false;
    outStore = idStr.substr(0, p);
    outRest = idStr.substr(p + 1);
    return !outStore.empty() && !outRest.empty();
}

static bool TryGetFirstTokenBeforeColon(const std::wstring& rest, std::wstring& first) {
    first.clear();
    size_t p = rest.find(L':');
    first = (p == std::wstring::npos) ? rest : rest.substr(0, p);
    return !first.empty();
}

// ------------------------------------------------------------
// Title resolvers
// ------------------------------------------------------------

static bool JsonExtractStringValueAfter(const std::string& s, size_t startPos, std::string& out) {
    out.clear();
    size_t p = s.find('"', startPos);
    if (p == std::string::npos) return false;
    ++p;

    while (p < s.size()) {
        char c = s[p++];
        if (c == '"') break;
        if (c == '\\' && p < s.size()) {
            char n = s[p++];
            if (n == 'n') out.push_back('\n');
            else if (n == 't') out.push_back('\t');
            else out.push_back(n);
        } else {
            out.push_back(c);
        }
    }
    return !out.empty();
}

// Steam: Steam Store API appdetails -> data.name
static std::wstring ResolveSteamTitle_SteamStoreApi(const std::wstring& steamAppId) {
    std::wstring cacheKey = L"steam:" + steamAppId;
    auto it = gTitleCache.find(cacheKey);
    if (it != gTitleCache.end()) return it->second;

    std::wstring url = L"https://store.steampowered.com/api/appdetails?appids=" + steamAppId + L"&l=english";
    std::string body;
    if (!HttpGetUtf8(url, L"", body)) return L"";

    size_t namePos = body.find("\"name\"");
    if (namePos == std::string::npos) return L"";

    size_t colon = body.find(':', namePos);
    if (colon == std::string::npos) return L"";

    std::string titleUtf8;
    if (!JsonExtractStringValueAfter(body, colon + 1, titleUtf8)) return L"";

    titleUtf8 = TrimA(titleUtf8);
    if (titleUtf8.empty()) return L"";

    std::wstring title = Utf8ToWide(titleUtf8);
    if (!title.empty()) gTitleCache[cacheKey] = title;
    return title;
}

static bool JsonExtractStringField(const std::string& json, const char* key, std::string& out) {
    out.clear();
    std::string needle = std::string("\"") + key + "\":\"";
    size_t p = json.find(needle);
    if (p == std::string::npos) return false;
    p += needle.size();

    while (p < json.size()) {
        char c = json[p++];
        if (c == '"') break;
        if (c == '\\' && p < json.size()) {
            char n = json[p++];
            if (n == 'n') out.push_back('\n');
            else if (n == 't') out.push_back('\t');
            else out.push_back(n);
        } else {
            out.push_back(c);
        }
    }
    return !out.empty();
}

static std::wstring ResolveGogTitle_GogApi(const std::wstring& productId) {
    std::wstring cacheKey = L"gog:" + productId;
    auto it = gTitleCache.find(cacheKey);
    if (it != gTitleCache.end()) return it->second;

    std::wstring url = L"https://api.gog.com/products/" + productId;
    std::string body;
    if (!HttpGetUtf8(url, L"", body)) return L"";

    std::string titleUtf8;
    if (!JsonExtractStringField(body, "title", titleUtf8)) return L"";

    // Decode JSON unicode escapes like \u0027 -> '
    titleUtf8 = DecodeJsonUnicodeEscapes(titleUtf8);

    std::wstring title = Utf8ToWide(titleUtf8);
    if (!title.empty()) gTitleCache[cacheKey] = title;
    return title;
}

static void EnsureUbiMapLoaded() {
    if (gUbiMapLoaded) return;
    gUbiMapLoaded = true;

    const std::wstring url = L"https://raw.githubusercontent.com/Haoose/UPLAY_GAME_ID/master/README.md";
    std::string md;
    if (!HttpGetUtf8(url, L"", md)) return;

    std::regex lineRe(R"(^\s*(\d+)\s*-\s*(.+?)\s*$)");
    size_t start = 0;
    while (start < md.size()) {
        size_t end = md.find('\n', start);
        if (end == std::string::npos) end = md.size();
        std::string line = md.substr(start, end - start);
        start = end + 1;

        std::smatch m;
        if (std::regex_match(line, m, lineRe) && m.size() >= 3) {
            std::wstring wid = Utf8ToWide(m[1].str());
            std::wstring wname = Utf8ToWide(m[2].str());
            if (!wid.empty() && !wname.empty()) gUbiIdToName[wid] = wname;
        }
    }
}

static std::wstring ResolveUbiTitle_GitHub(const std::wstring& uplayId) {
    EnsureUbiMapLoaded();
    auto it = gUbiIdToName.find(uplayId);
    if (it != gUbiIdToName.end()) return it->second;
    return L"";
}

// Epic: first Base game row, title is 2nd column (after offer name)
static std::wstring ResolveEpicTitle_EGData_FirstBaseGameTitleColumn(const std::wstring& sandboxId) {
    std::wstring cacheKey = L"epic:sandbox:basegame:titlecol:" + sandboxId;
    auto it = gTitleCache.find(cacheKey);
    if (it != gTitleCache.end()) return it->second;

    std::wstring url = L"https://egdata.app/sandboxes/" + sandboxId + L"/offers";
    std::string html;
    if (!HttpGetUtf8(url, L"", html)) return L"";

    size_t basePos = html.find("Base game");
    if (basePos == std::string::npos) return L"";

    size_t trStart = html.rfind("<tr", basePos);
    if (trStart == std::string::npos) return L"";
    size_t trEnd = html.find("</tr>", basePos);
    if (trEnd == std::string::npos) return L"";
    std::string row = html.substr(trStart, trEnd - trStart);

    std::vector<std::string> tds;
    tds.reserve(4);

    size_t p = 0;
    while (tds.size() < 3) {
        size_t tdOpen = row.find("<td", p);
        if (tdOpen == std::string::npos) break;
        size_t gt = row.find('>', tdOpen);
        if (gt == std::string::npos) break;
        size_t tdClose = row.find("</td>", gt);
        if (tdClose == std::string::npos) break;

        std::string tdInner = row.substr(gt + 1, tdClose - (gt + 1));

        // strip tags (simple)
        std::string out;
        out.reserve(tdInner.size());
        bool inTag = false;
        for (char c : tdInner) {
            if (c == '<') { inTag = true; continue; }
            if (c == '>') { inTag = false; continue; }
            if (!inTag) out.push_back(c);
        }

        tds.push_back(TrimA(out));
        p = tdClose + 5;
    }

    if (tds.size() < 2) return L"";
    std::string title = TrimA(tds[1]);
    if (title.empty()) return L"";

    // Decode HTML entities like &#x27; -> '
    title = DecodeHtmlEntities(title);

    std::wstring wtitle = Utf8ToWide(title);
    if (!wtitle.empty()) gTitleCache[cacheKey] = wtitle;
    return wtitle;
}

// ------------------------------------------------------------
// Cache file (JS) read/write
// ------------------------------------------------------------

struct CacheState {
    int iconSize = 64;
    std::wstring layout = L"list"; // "list" or "grid"
    AppConfig config;
    std::vector<GameEntry> items;
};

static std::string BoolToJs(bool b) { return b ? "true" : "false"; }

static std::string WToJsString(const std::wstring& w) {
    return "\"" + JsEscape(WideToUtf8(w)) + "\"";
}

static std::wstring JsStringToW(const std::string& jsStrNoQuotes) {
    return Utf8ToWide(JsUnescape(jsStrNoQuotes));
}

static bool SaveCacheFile(const CacheState& st) {
    // Ensure cache directories exist
    EnsureCacheDirectories();
    
    std::wstring path = GetCacheFilePath();

    std::string out;
    out += "// XboxAppArtUpdater cache v1\n";
    out += "const cache = {\n";
    out += "  version: 1,\n";
    out += "  iconSize: " + std::to_string(st.iconSize) + ",\n";
    out += "  layout: " + WToJsString(st.layout) + ",\n";
    out += "  config: {\n";
    out += "    steamGridDbKey: " + WToJsString(st.config.steamGridDbKey) + ",\n";
    out += "    thirdPartyBasePath: " + WToJsString(st.config.thirdPartyBasePath) + ",\n";
    out += "    darkMode: " + BoolToJs(st.config.darkMode) + ",\n";
    out += "    parallelDownloads: " + std::to_string(st.config.parallelDownloads) + ",\n";
    out += "    minImageSize: " + std::to_string(st.config.minImageSize) + "\n";
    out += "  },\n";
    out += "  games: [\n";

    for (size_t i = 0; i < st.items.size(); ++i) {
        const auto& e = st.items[i];
        out += "    {";
        out += "store:" + WToJsString(e.store) + ",";
        out += "id:" + WToJsString(e.idStr) + ",";
        out += "expected:" + WToJsString(e.expectedFileName) + ",";
        out += "hasArt:" + BoolToJs(e.hasArt) + ",";
        out += "fileName:" + WToJsString(e.fileName) + ",";
        out += "filePath:" + WToJsString(e.filePath) + ",";
        out += "title:" + WToJsString(e.title);
        if (!e.customTitle.empty()) {
            out += ",customTitle:" + WToJsString(e.customTitle);
        }
        out += "}";
        if (i + 1 < st.items.size()) out += ",";
        out += "\n";
    }

    out += "  ]\n";
    out += "};\n";
    out += "module.exports = cache;\n";

    return WriteAllTextA(path, out);
}

// Parser for our cache format.
// IMPORTANT: regex strings here MUST be raw strings to avoid C++ escape issues.
// ALSO IMPORTANT: use custom delimiters because patterns contain )"
static bool LoadCacheFile(CacheState& st) {
    st = CacheState{};

    std::wstring path = GetCacheFilePath();
    if (!fs::exists(path)) return false;

    std::string text = ReadAllTextA(path);
    if (text.empty()) return false;

    // iconSize
    {
        const std::regex re(R"IS(iconSize\s*:\s*(\d+))IS");
        std::smatch m;
        if (std::regex_search(text, m, re) && m.size() >= 2) {
            st.iconSize = std::max(16, std::min(512, atoi(m[1].str().c_str())));
        }
    }

    // layout
    {
        const std::regex re(R"LAY(layout\s*:\s*"([^"]*)")LAY");
        std::smatch m;
        if (std::regex_search(text, m, re) && m.size() >= 2) {
            st.layout = JsStringToW(m[1].str());
            if (st.layout != L"grid" && st.layout != L"list") st.layout = L"list";
        }
    }

    // config.steamGridDbKey
    {
        const std::regex re(R"KEY(steamGridDbKey\s*:\s*"([^"]*)")KEY");
        std::smatch m;
        if (std::regex_search(text, m, re) && m.size() >= 2) {
            st.config.steamGridDbKey = JsStringToW(m[1].str());
        }
    }

    // config.thirdPartyBasePath
    {
        const std::regex re(R"PATH(thirdPartyBasePath\s*:\s*"([^"]*)")PATH");
        std::smatch m;
        if (std::regex_search(text, m, re) && m.size() >= 2) {
            st.config.thirdPartyBasePath = JsStringToW(m[1].str());
        }
    }

    // config.darkMode
    {
        const std::regex re(R"DARK(darkMode\s*:\s*(true|false))DARK");
        std::smatch m;
        if (std::regex_search(text, m, re) && m.size() >= 2) {
            st.config.darkMode = (m[1].str() == "true");
        }
    }

    // config.parallelDownloads
    {
        const std::regex re(R"PARA(parallelDownloads\s*:\s*(\d+))PARA");
        std::smatch m;
        if (std::regex_search(text, m, re) && m.size() >= 2) {
            st.config.parallelDownloads = atoi(m[1].str().c_str());
        }
    }

    // config.minImageSize
    {
        const std::regex re(R"PARA(minImageSize\s*:\s*(\d+))PARA");
        std::smatch m;
        if (std::regex_search(text, m, re) && m.size() >= 2) {
            st.config.minImageSize = atoi(m[1].str().c_str());
            if (st.config.minImageSize < 0) st.config.minImageSize = 0;
        }
    }

    // games objects
    // Matches lines like:
    // {store:"steam",id:"steam:123",expected:"steam_123.png",hasArt:true,fileName:"...",filePath:"...",title:"...",...}
    // Accepts extra fields (like customTitle) after title
    const std::regex objRe(
        R"REGEX(\{\s*store:"([^"]*)",\s*id:"([^"]*)",\s*expected:"([^"]*)",\s*hasArt:(true|false),\s*fileName:"([^"]*)",\s*filePath:"([^"]*)",\s*title:"([^"]*)"((?:,[^}]*)?)\})REGEX"
    );

    auto beginIt = std::sregex_iterator(text.begin(), text.end(), objRe);
    auto endIt   = std::sregex_iterator();

    for (auto it = beginIt; it != endIt; ++it) {
        std::smatch m = *it;
        if (m.size() < 9) continue;

        GameEntry e{};
        e.store = JsStringToW(m[1].str());
        e.idStr = JsStringToW(m[2].str());
        e.expectedFileName = JsStringToW(m[3].str());
        e.hasArt = (m[4].str() == "true");
        e.fileName = JsStringToW(m[5].str());
        e.filePath = JsStringToW(m[6].str());
        e.title = JsStringToW(m[7].str());
        // Parse customTitle if present in the extra fields
        std::string extras = m[8].str();
        std::smatch m2;
        std::regex customTitleRe(R"(customTitle:\"([^\"]*)\")");
        if (std::regex_search(extras, m2, customTitleRe) && m2.size() >= 2) {
            e.customTitle = JsStringToW(m2[1].str());
        }
        st.items.push_back(std::move(e));
    }

    // Consider it "loaded" even if no games; config can still be present.
    return true;
}

// ------------------------------------------------------------
// Thumbnail helpers (GDI+)
// ------------------------------------------------------------

static HBITMAP CreatePlaceholderBitmap(int S) {
    Bitmap bmp(S, S, PixelFormat32bppARGB);
    Graphics g(&bmp);
    g.SetSmoothingMode(SmoothingModeHighQuality);
    g.SetInterpolationMode(InterpolationModeHighQualityBicubic);

    // Use dark or light colors based on mode
    if (gDarkMode) {
        g.Clear(Color(255, 45, 45, 45));
        Pen pen1(Color(255, 70, 70, 70), (REAL)std::max(2, S / 40));
        g.DrawRectangle(&pen1, 2, 2, S - 4, S - 4);
        Pen pen2(Color(255, 80, 80, 80), (REAL)std::max(2, S / 35));
        g.DrawLine(&pen2, S / 6, S / 6, S - S / 6, S - S / 6);
        g.DrawLine(&pen2, S - S / 6, S / 6, S / 6, S - S / 6);
    } else {
        g.Clear(Color(255, 240, 240, 240));
        Pen pen1(Color(255, 200, 200, 200), (REAL)std::max(2, S / 40));
        g.DrawRectangle(&pen1, 2, 2, S - 4, S - 4);
        Pen pen2(Color(255, 190, 190, 190), (REAL)std::max(2, S / 35));
        g.DrawLine(&pen2, S / 6, S / 6, S - S / 6, S - S / 6);
        g.DrawLine(&pen2, S - S / 6, S / 6, S / 6, S - S / 6);
    }

    HBITMAP hbm = nullptr;
    bmp.GetHBITMAP(Color(0, 0, 0, 0), &hbm);
    return hbm;
}

static HBITMAP LoadPngToHBITMAP_AspectFit(const std::wstring& pngPath, int S) {
    Bitmap src(pngPath.c_str());
    if (src.GetLastStatus() != Ok || src.GetWidth() == 0 || src.GetHeight() == 0) return nullptr;

    Bitmap dst(S, S, PixelFormat32bppARGB);
    Graphics g(&dst);
    g.SetSmoothingMode(SmoothingModeHighQuality);
    g.SetInterpolationMode(InterpolationModeHighQualityBicubic);
    g.SetPixelOffsetMode(PixelOffsetModeHighQuality);
    g.Clear(gDarkMode ? Color(255, 45, 45, 45) : Color(255, 245, 245, 245));

    const double sw = (double)src.GetWidth();
    const double sh = (double)src.GetHeight();
    const double scale = (std::min)((double)S / sw, (double)S / sh);

    const int dw = (int)(sw * scale);
    const int dh = (int)(sh * scale);
    const int dx = (S - dw) / 2;
    const int dy = (S - dh) / 2;

    Rect destRect(dx, dy, dw, dh);
    g.DrawImage(&src, destRect);

    HBITMAP hbm = nullptr;
    dst.GetHBITMAP(Color(0, 0, 0, 0), &hbm);
    return hbm;
}

static void DestroyImageList() {
    if (gImgList) {
        ImageList_Destroy(gImgList);
        gImgList = nullptr;
    }
    gPlaceholderImgIndex = -1;
    gImageIndexCache.clear();
}

static void EnsureImageList() {
    if (gImgList) return;

    const int S = gIconSize;
    gImgList = ImageList_Create(S, S, ILC_COLOR32 | ILC_MASK, 64, 64);

    HBITMAP ph = CreatePlaceholderBitmap(S);
    if (ph) {
        gPlaceholderImgIndex = ImageList_Add(gImgList, ph, nullptr);
        DeleteObject(ph);
    }

    if (gList && gImgList) {
        ListView_SetImageList(gList, gImgList, LVSIL_NORMAL);
        ListView_SetImageList(gList, gImgList, LVSIL_SMALL);
    }
}

static int GetImageIndexForPathOrPlaceholder(const std::wstring& pngPathOrEmpty) {
    EnsureImageList();

    if (pngPathOrEmpty.empty() || !fs::exists(pngPathOrEmpty)) {
        return (gPlaceholderImgIndex >= 0) ? gPlaceholderImgIndex : -1;
    }

    auto it = gImageIndexCache.find(pngPathOrEmpty);
    if (it != gImageIndexCache.end()) return it->second;

    HBITMAP hbm = LoadPngToHBITMAP_AspectFit(pngPathOrEmpty, gIconSize);
    if (!hbm) {
        int idx = (gPlaceholderImgIndex >= 0) ? gPlaceholderImgIndex : -1;
        gImageIndexCache[pngPathOrEmpty] = idx;
        return idx;
    }

    int idx = ImageList_Add(gImgList, hbm, nullptr);
    DeleteObject(hbm);

    if (idx < 0) idx = (gPlaceholderImgIndex >= 0) ? gPlaceholderImgIndex : -1;
    gImageIndexCache[pngPathOrEmpty] = idx;
    return idx;
}

// ------------------------------------------------------------
// ListView columns (list mode only)
// ------------------------------------------------------------

static void ClearColumns(HWND lv) {
    HWND hdr = ListView_GetHeader(lv);
    if (!hdr) return;
    int count = Header_GetItemCount(hdr);
    for (int i = count - 1; i >= 0; --i) {
        ListView_DeleteColumn(lv, i);
    }
}

static void SetupListColumns(HWND lv) {
    ClearColumns(lv);

    LVCOLUMN c{};
    c.mask = LVCF_TEXT | LVCF_WIDTH;

    c.cx = gIconSize + 80; c.pszText = (LPWSTR)L"Art";         ListView_InsertColumn(lv, 0, &c);
    c.cx = 80;             c.pszText = (LPWSTR)L"Store";       ListView_InsertColumn(lv, 1, &c);
    c.cx = 360;            c.pszText = (LPWSTR)L"Game ID";     ListView_InsertColumn(lv, 2, &c);
    c.cx = 280;            c.pszText = (LPWSTR)L"Title";       ListView_InsertColumn(lv, 3, &c);
    c.cx = 110;            c.pszText = (LPWSTR)L"Missing art"; ListView_InsertColumn(lv, 4, &c);
    c.cx = 260;            c.pszText = (LPWSTR)L"File";        ListView_InsertColumn(lv, 5, &c);
}

// ------------------------------------------------------------
// Header subclass for dark mode custom draw
// ------------------------------------------------------------

static LRESULT CALLBACK HeaderSubclassProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        
        RECT clientRc;
        GetClientRect(hWnd, &clientRc);
        
        // Colors based on mode
        COLORREF bgColor = gDarkMode ? RGB(45, 45, 45) : RGB(255, 255, 255);
        COLORREF textColor = gDarkMode ? RGB(255, 255, 255) : RGB(0, 0, 0);
        COLORREF borderColor = gDarkMode ? RGB(70, 70, 70) : RGB(200, 200, 200);
        
        // Fill entire header background
        HBRUSH bgBrush = CreateSolidBrush(bgColor);
        FillRect(hdc, &clientRc, bgBrush);
        DeleteObject(bgBrush);
        
        // Get header info
        int itemCount = Header_GetItemCount(hWnd);
        
        // Create bold font
        HFONT hFont = (HFONT)SendMessageW(hWnd, WM_GETFONT, 0, 0);
        LOGFONTW lf{};
        if (hFont) {
            GetObjectW(hFont, sizeof(lf), &lf);
        } else {
            SystemParametersInfoW(SPI_GETICONTITLELOGFONT, sizeof(lf), &lf, 0);
        }
        lf.lfWeight = FW_BOLD;
        HFONT boldFont = CreateFontIndirectW(&lf);
        HFONT oldFont = (HFONT)SelectObject(hdc, boldFont);
        
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, textColor);
        
        for (int i = 0; i < itemCount; i++) {
            RECT itemRc;
            Header_GetItemRect(hWnd, i, &itemRc);
            
            // Draw separator line
            HPEN borderPen = CreatePen(PS_SOLID, 1, borderColor);
            HPEN oldPen = (HPEN)SelectObject(hdc, borderPen);
            MoveToEx(hdc, itemRc.right - 1, itemRc.top + 4, nullptr);
            LineTo(hdc, itemRc.right - 1, itemRc.bottom - 4);
            SelectObject(hdc, oldPen);
            DeleteObject(borderPen);
            
            // Get item text
            wchar_t text[256]{};
            HDITEMW hdi{};
            hdi.mask = HDI_TEXT;
            hdi.pszText = text;
            hdi.cchTextMax = _countof(text);
            Header_GetItem(hWnd, i, &hdi);
            
            // Draw text
            RECT textRc = itemRc;
            textRc.left += 8;
            textRc.right -= 8;
            DrawTextW(hdc, text, -1, &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
        }
        
        // Draw bottom border
        HPEN borderPen = CreatePen(PS_SOLID, 1, borderColor);
        HPEN oldPen = (HPEN)SelectObject(hdc, borderPen);
        MoveToEx(hdc, clientRc.left, clientRc.bottom - 1, nullptr);
        LineTo(hdc, clientRc.right, clientRc.bottom - 1);
        SelectObject(hdc, oldPen);
        DeleteObject(borderPen);
        
        SelectObject(hdc, oldFont);
        DeleteObject(boldFont);
        
        EndPaint(hWnd, &ps);
        return 0;
    }
    case WM_ERASEBKGND: {
        HDC hdc = (HDC)wParam;
        RECT rc;
        GetClientRect(hWnd, &rc);
        HBRUSH bgBrush = CreateSolidBrush(gDarkMode ? RGB(45, 45, 45) : RGB(255, 255, 255));
        FillRect(hdc, &rc, bgBrush);
        DeleteObject(bgBrush);
        return 1;
    }
    }
    return CallWindowProcW(gOriginalHeaderProc, hWnd, msg, wParam, lParam);
}

// Status bar subclass procedure for dark mode
static LRESULT CALLBACK StatusBarSubclassProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_PAINT: {
        if (gDarkMode) {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            
            RECT clientRc;
            GetClientRect(hWnd, &clientRc);
            
            // Fill background
            HBRUSH bgBrush = CreateSolidBrush(RGB(32, 32, 32));
            FillRect(hdc, &clientRc, bgBrush);
            DeleteObject(bgBrush);
            
            // Get the text
            int textLen = (int)SendMessageW(hWnd, WM_GETTEXTLENGTH, 0, 0);
            std::wstring text(textLen + 1, L'\0');
            SendMessageW(hWnd, WM_GETTEXT, text.size(), (LPARAM)text.data());
            
            // Set up font
            HFONT hFont = (HFONT)SendMessageW(hWnd, WM_GETFONT, 0, 0);
            HFONT oldFont = nullptr;
            if (hFont) {
                oldFont = (HFONT)SelectObject(hdc, hFont);
            }
            
            // Draw text
            SetBkMode(hdc, TRANSPARENT);
            SetTextColor(hdc, RGB(255, 255, 255));
            
            RECT textRc = clientRc;
            textRc.left += 6;
            textRc.right -= 6;
            DrawTextW(hdc, text.c_str(), -1, &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
            
            if (oldFont) {
                SelectObject(hdc, oldFont);
            }
            
            // Draw top border
            HPEN borderPen = CreatePen(PS_SOLID, 1, RGB(70, 70, 70));
            HPEN oldPen = (HPEN)SelectObject(hdc, borderPen);
            MoveToEx(hdc, clientRc.left, clientRc.top, nullptr);
            LineTo(hdc, clientRc.right, clientRc.top);
            SelectObject(hdc, oldPen);
            DeleteObject(borderPen);
            
            EndPaint(hWnd, &ps);
            return 0;
        }
        break;
    }
    case WM_ERASEBKGND: {
        if (gDarkMode) {
            return 1;  // We'll paint in WM_PAINT
        }
        break;
    }
    }
    return CallWindowProcW(gOriginalStatusBarProc, hWnd, msg, wParam, lParam);
}

// Tab control subclass procedure for custom background
static LRESULT CALLBACK TabSubclassProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_ERASEBKGND: {
        HDC hdc = (HDC)wParam;
        RECT rc;
        GetClientRect(hWnd, &rc);
        COLORREF bgColor = gDarkMode ? RGB(32, 32, 32) : RGB(255, 255, 255);
        HBRUSH bgBrush = CreateSolidBrush(bgColor);
        FillRect(hdc, &rc, bgBrush);
        DeleteObject(bgBrush);
        return 1;
    }
    case WM_PAINT: {
        // Custom paint to avoid focus rectangle and bevel
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        
        RECT rc;
        GetClientRect(hWnd, &rc);
        
        // Fill entire background
        COLORREF bgColor = gDarkMode ? RGB(32, 32, 32) : RGB(255, 255, 255);
        HBRUSH bgBrush = CreateSolidBrush(bgColor);
        FillRect(hdc, &rc, bgBrush);
        DeleteObject(bgBrush);
        
        // Let the owner-draw handle the tabs
        EndPaint(hWnd, &ps);
        
        // Now send WM_DRAWITEM for each tab
        int tabCount = TabCtrl_GetItemCount(hWnd);
        HWND parent = GetParent(hWnd);
        for (int i = 0; i < tabCount; i++) {
            RECT tabRc;
            TabCtrl_GetItemRect(hWnd, i, &tabRc);
            
            DRAWITEMSTRUCT dis = {};
            dis.CtlType = ODT_TAB;
            dis.CtlID = GetDlgCtrlID(hWnd);
            dis.itemID = i;
            dis.itemAction = ODA_DRAWENTIRE;
            dis.itemState = (i == TabCtrl_GetCurSel(hWnd)) ? ODS_SELECTED : 0;
            dis.hwndItem = hWnd;
            dis.hDC = GetDC(hWnd);
            dis.rcItem = tabRc;
            
            SendMessageW(parent, WM_DRAWITEM, dis.CtlID, (LPARAM)&dis);
            ReleaseDC(hWnd, dis.hDC);
        }
        return 0;
    }
    case WM_SETFOCUS:
    case WM_KILLFOCUS:
        // Suppress focus changes to avoid dotted rectangle
        InvalidateRect(hWnd, nullptr, TRUE);
        return 0;
    }
    return CallWindowProcW(gOriginalTabProc, hWnd, msg, wParam, lParam);
}

// ------------------------------------------------------------
// ListView recreation (reliable view switching)
// ------------------------------------------------------------

static void RecreateListView(HWND parent) {
    RECT rc{};
    GetClientRect(parent, &rc);

    int wdt = rc.right - rc.left;
    int hgt = rc.bottom - rc.top;

    if (gList) DestroyWindow(gList);
    gList = nullptr;

    DWORD baseStyle = WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_SHAREIMAGELISTS;
    DWORD viewStyle = gGridLayout ? LVS_ICON : LVS_REPORT;

    gList = CreateWindowW(WC_LISTVIEWW, L"",
        baseStyle | viewStyle,
        10, 115, wdt - 20, hgt - 125,
        parent, (HMENU)(INT_PTR)IDC_LIST,
        GetModuleHandleW(nullptr), nullptr);

    if (!gList) return;

    if (gGridLayout) {
        ListView_SetExtendedListViewStyle(gList, LVS_EX_DOUBLEBUFFER);
    } else {
        ListView_SetExtendedListViewStyle(
            gList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_DOUBLEBUFFER);
        SetupListColumns(gList);
        
        // Subclass the header control for custom styling (both dark and light mode)
        HWND hHeader = ListView_GetHeader(gList);
        if (hHeader) {
            gOriginalHeaderProc = (WNDPROC)SetWindowLongPtrW(hHeader, GWLP_WNDPROC, (LONG_PTR)HeaderSubclassProc);
        }
    }

    // Apply mode-specific styling to the ListView
    if (gDarkMode) {
        SetWindowTheme(gList, L"DarkMode_Explorer", nullptr);
        ListView_SetBkColor(gList, RGB(32, 32, 32));
        ListView_SetTextBkColor(gList, RGB(32, 32, 32));
        ListView_SetTextColor(gList, RGB(255, 255, 255));
    } else {
        SetWindowTheme(gList, L"Explorer", nullptr);
        ListView_SetBkColor(gList, RGB(255, 255, 255));
        ListView_SetTextBkColor(gList, RGB(255, 255, 255));
        ListView_SetTextColor(gList, RGB(0, 0, 0));
    }

    DestroyImageList();
    EnsureImageList();

    ListView_SetImageList(gList, gImgList, LVSIL_NORMAL);
    ListView_SetImageList(gList, gImgList, LVSIL_SMALL);

    if (gGridLayout) {
        ListView_SetView(gList, LV_VIEW_ICON);
        int pad = std::max(16, gIconSize / 2);
        int spacingX = gIconSize + pad + 60;
        int spacingY = gIconSize + pad + 40;
        SendMessageW(gList, LVM_SETICONSPACING, 0, MAKELPARAM(spacingX, spacingY));
    } else {
        ListView_SetView(gList, LV_VIEW_DETAILS);
    }
}

// ------------------------------------------------------------
// Scan logic
// ------------------------------------------------------------

static std::vector<GameEntry> ScanCache() {
    std::vector<GameEntry> out;
    std::wstring base = GetThirdPartyLibrariesPath();

    struct StoreInfo { const wchar_t* store; const wchar_t* manifest; const wchar_t* prefix; };
    StoreInfo stores[] = {
        {L"steam", L"steam.manifest", L"steam:"},
        {L"epic",  L"epic.manifest",  L"epic:"},
        {L"gog",   L"gog.manifest",   L"gog:"},
        {L"ubi",   L"ubi.manifest",   L"ubi:"},
        {L"bnet",  L"bnet.manifest",  L"bnet:"},
    };

    for (auto& s : stores) {
        fs::path dir = fs::path(base) / s.store;
        if (!fs::exists(dir) || !fs::is_directory(dir)) continue;

        fs::path mf = dir / s.manifest;
        auto ids = fs::exists(mf)
            ? ExtractIdsFromManifest(mf.wstring(), s.prefix)
            : std::vector<std::wstring>();

        for (auto& id : ids) {
            GameEntry e{};
            e.store = s.store;
            e.idStr = id;
            e.expectedFileName = ExpectedPngFromManifestId(s.store, id);

            fs::path expected = dir / e.expectedFileName;
            e.hasArt = fs::exists(expected) && fs::is_regular_file(expected);
            if (e.hasArt) {
                e.filePath = expected.wstring();
                e.fileName = e.expectedFileName;
            }
            out.push_back(std::move(e));
        }
    }
    return out;
}

// ------------------------------------------------------------
// Title resolving dispatcher
// ------------------------------------------------------------

static std::wstring ResolveTitleForEntry(const GameEntry& e) {
    std::wstring store, rest;
    if (!TrySplitStoreAndRest(e.idStr, store, rest)) return L"";

    if (store == L"steam" && IsAllDigits(rest)) {
        return ResolveSteamTitle_SteamStoreApi(rest);
    }
    if (store == L"gog" && IsAllDigits(rest)) {
        return ResolveGogTitle_GogApi(rest);
    }
    if (store == L"ubi" && IsAllDigits(rest)) {
        return ResolveUbiTitle_GitHub(rest);
    }
    if (store == L"epic") {
        std::wstring sandboxId;
        if (TryGetFirstTokenBeforeColon(rest, sandboxId)) {
            return ResolveEpicTitle_EGData_FirstBaseGameTitleColumn(sandboxId);
        }
    }
    return L"";
}

// ------------------------------------------------------------
// Rendering
// ------------------------------------------------------------

// Apply filters and return filtered list
static std::vector<GameEntry> ApplyFilters(const std::vector<GameEntry>& items) {
    std::vector<GameEntry> filtered;
    filtered.reserve(items.size());
    
    for (const auto& e : items) {
        // Store filter
        if (!gFilterStore.empty()) {
            if (gFilterStore == L"steam" && e.store != L"steam") continue;
            if (gFilterStore == L"epic" && e.store != L"epic") continue;
            if (gFilterStore == L"gog" && e.store != L"gog") continue;
            if (gFilterStore == L"ubi" && e.store != L"ubi") continue;
        }
        
        // Art filter: 0=all, 1=has art, 2=missing art
        if (gFilterArt == 1 && !e.hasArt) continue;
        if (gFilterArt == 2 && e.hasArt) continue;
        
        // Unresolved filter
        if (gFilterUnresolved && e.title.empty()) {
            // Show only unresolved
        } else if (gFilterUnresolved && !e.title.empty()) {
            continue;
        }

        // Search filter (case-insensitive)
        if (!gFilterSearch.empty()) {
            std::wstring titleLower = e.title;
            std::wstring searchLower = gFilterSearch;
            std::transform(titleLower.begin(), titleLower.end(), titleLower.begin(), ::towlower);
            std::transform(searchLower.begin(), searchLower.end(), searchLower.begin(), ::towlower);
            if (titleLower.find(searchLower) == std::wstring::npos) continue;
        }
        
        filtered.push_back(e);
    }
    
    return filtered;
}

static void PopulateListFromItems(const std::vector<GameEntry>& items) {
    if (!gList) return;

    ListView_DeleteAllItems(gList);
    gImageIndexCache.clear();
    
    // Show/hide the "no scan" message overlay
    if (gNoScanMessage) {
        bool showMessage = items.empty() && !gHasScanned;
        ShowWindow(gNoScanMessage, showMessage ? SW_SHOW : SW_HIDE);
        // If showing message, bring it to top
        if (showMessage) {
            SetWindowPos(gNoScanMessage, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE);
        }
    }
    
    // Apply filters and save for double-click handling
    gFilteredItems = ApplyFilters(items);

    for (size_t idx = 0; idx < gFilteredItems.size(); ++idx) {
        const auto& e = gFilteredItems[idx];
        int imgIndex = GetImageIndexForPathOrPlaceholder(e.hasArt ? e.filePath : L"");

        std::wstring gridLabel = !e.DisplayTitle().empty() ? e.DisplayTitle() : e.idStr;
        std::wstring listArtLabel = e.hasArt ? L"" : L"(missing)";

        LVITEM it{};
        it.mask = LVIF_TEXT | LVIF_IMAGE | LVIF_PARAM;
        it.iItem = ListView_GetItemCount(gList);
        it.iImage = imgIndex;
        it.lParam = (LPARAM)idx;  // Store index into gFilteredItems
        it.pszText = (LPWSTR)(gGridLayout ? gridLabel.c_str() : listArtLabel.c_str());

        int row = ListView_InsertItem(gList, &it);

        if (!gGridLayout) {
            ListView_SetItemText(gList, row, 1, (LPWSTR)e.store.c_str());
            ListView_SetItemText(gList, row, 2, (LPWSTR)e.idStr.c_str());
            ListView_SetItemText(gList, row, 3, (LPWSTR)(e.DisplayTitle().empty() ? L"(unresolved)" : e.DisplayTitle().c_str()));
            ListView_SetItemText(gList, row, 4, (LPWSTR)(e.hasArt ? L"No" : L"Yes"));

            std::wstring fileCell = e.hasArt
                ? e.fileName
                : (!e.expectedFileName.empty() ? e.expectedFileName : L"(none)");
            ListView_SetItemText(gList, row, 5, (LPWSTR)fileCell.c_str());
        }
    }

    if (gGridLayout) {
        ListView_SetView(gList, LV_VIEW_ICON);
    } else {
        ListView_SetView(gList, LV_VIEW_DETAILS);
    }
    
    // Show Quick Apply button if list has items, hide if empty
    if (gBtnQuickApply) {
        ShowWindow(gBtnQuickApply, gFilteredItems.empty() ? SW_HIDE : SW_SHOW);
    }
}

// ------------------------------------------------------------
// Checkbox handler: update unresolved filter and refresh list
static void OnUnresolvedCheckChanged() {
    gFilterUnresolved = (gChkUnresolved && SendMessageW(gChkUnresolved, BM_GETCHECK, 0, 0) == BST_CHECKED);
    PopulateListFromItems(gLastItems);
}

// Scan action (fast + cache-aware)
// ------------------------------------------------------------

// Helper to update status bar
static void SetStatusText(const std::wstring& text) {
    if (gStatusBar) {
        SendMessageW(gStatusBar, SB_SETTEXTW, 0, (LPARAM)text.c_str());
        UpdateWindow(gStatusBar);
    }
}

// Helper to process messages during long operations
static void ProcessPendingMessages() {
    MSG msg;
    while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
}

// Forward declarations for SteamGridDB functions used by Quick Apply
static std::wstring SteamGridDbGetGameIdBySteamAppId(const std::wstring& steamAppId, const std::wstring& apiKey);
static std::wstring SteamGridDbSearchGame(const std::wstring& gameName, const std::wstring& apiKey);
static std::vector<SteamGridDbImage> SteamGridDbGetGrids(const std::wstring& gameId, const std::wstring& apiKey);
static std::vector<SteamGridDbImage> SteamGridDbGetIcons(const std::wstring& gameId, const std::wstring& apiKey);
static std::vector<SteamGridDbImage> SteamGridDbGetHeroes(const std::wstring& gameId, const std::wstring& apiKey);
static std::vector<SteamGridDbImage> SteamGridDbGetLogos(const std::wstring& gameId, const std::wstring& apiKey);
static std::vector<SteamGridDbImage> WebSearchGameImages(const std::wstring& gameTitle, const std::wstring& resolution, int pageIndex);
static void DebugLog(const std::wstring& msg);
static bool HttpDownloadBinary(const std::wstring& url, std::vector<unsigned char>& out);
static bool HttpDownloadBinaryCached(const std::wstring& url, std::vector<unsigned char>& out, const std::wstring& gameId);
static void SaveUiStateToCacheFile();

// Quick apply random art to a single game
// Returns true if art was applied successfully
static bool QuickApplyArtToGame(GameEntry& entry, std::wstring& outStatus, std::wstring& debugInfo) {
    debugInfo.clear();
    debugInfo += L"=== Quick Apply Debug Info ===\r\n";
    debugInfo += L"Game: " + entry.title + L"\r\n";
    debugInfo += L"ID: " + entry.idStr + L"\r\n\r\n";
    
    if (gConfig.steamGridDbKey.empty() && !gQuickApplyConfig.useWeb) {
        outStatus = L"No API key";
        debugInfo += L"ERROR: No API key and Web disabled\r\n";
        return false;
    }

    // Collect all available images from enabled sources
    std::vector<SteamGridDbImage> allImages;
    std::vector<SteamGridDbImage> webSearchResults;  // Keep web results separate for index selection
    
    // Web search (if enabled)
    if (gQuickApplyConfig.useWeb && !entry.title.empty()) {
        debugInfo += L"Fetching Web results...\r\n";
        debugInfo += L"Search query: " + entry.title + L" game cover imagesize:" + gQuickApplyConfig.webResolution + L"\r\n";
        webSearchResults = WebSearchGameImages(entry.title, gQuickApplyConfig.webResolution, 0);
        debugInfo += L"Found " + std::to_wstring(webSearchResults.size()) + L" web images\r\n";
        for (size_t i = 0; i < webSearchResults.size() && i < 10; i++) {
            debugInfo += L"  [" + std::to_wstring(i + 1) + L"] " + webSearchResults[i].url + L"\r\n";
        }
        if (webSearchResults.size() > 10) {
            debugInfo += L"  ... and " + std::to_wstring(webSearchResults.size() - 10) + L" more\r\n";
        }
        debugInfo += L"\r\n";
        // Don't add to allImages yet - we need them separate for specific index selection
    }
    
    // SteamGridDB sources (if API key available)
    if (!gConfig.steamGridDbKey.empty()) {
        debugInfo += L"SteamGridDB API key available\r\n";
        // Find the game on SteamGridDB
        std::wstring gameId;
        std::wstring store, rest;
        if (TrySplitStoreAndRest(entry.idStr, store, rest)) {
            if (store == L"steam" && IsAllDigits(rest)) {
                gameId = SteamGridDbGetGameIdBySteamAppId(rest, gConfig.steamGridDbKey);
            }
        }
        if (gameId.empty() && !entry.title.empty()) {
            gameId = SteamGridDbSearchGame(entry.title, gConfig.steamGridDbKey);
        }
        
        if (!gameId.empty()) {
            debugInfo += L"SteamGridDB Game ID: " + gameId + L"\r\n";
            // Get images from enabled sources
            if (gQuickApplyConfig.useGrids) {
                std::vector<SteamGridDbImage> grids = SteamGridDbGetGrids(gameId, gConfig.steamGridDbKey);
                debugInfo += L"Grids: " + std::to_wstring(grids.size()) + L" images\r\n";
                allImages.insert(allImages.end(), grids.begin(), grids.end());
            }
            if (gQuickApplyConfig.useHeroes) {
                std::vector<SteamGridDbImage> heroes = SteamGridDbGetHeroes(gameId, gConfig.steamGridDbKey);
                debugInfo += L"Heroes: " + std::to_wstring(heroes.size()) + L" images\r\n";
                allImages.insert(allImages.end(), heroes.begin(), heroes.end());
            }
            if (gQuickApplyConfig.useLogos) {
                std::vector<SteamGridDbImage> logos = SteamGridDbGetLogos(gameId, gConfig.steamGridDbKey);
                debugInfo += L"Logos: " + std::to_wstring(logos.size()) + L" images\r\n";
                allImages.insert(allImages.end(), logos.begin(), logos.end());
            }
            if (gQuickApplyConfig.useIcons) {
                std::vector<SteamGridDbImage> icons = SteamGridDbGetIcons(gameId, gConfig.steamGridDbKey);
                debugInfo += L"Icons: " + std::to_wstring(icons.size()) + L" images\r\n";
                allImages.insert(allImages.end(), icons.begin(), icons.end());
            }
        } else {
            debugInfo += L"Game not found on SteamGridDB\r\n";
        }
        debugInfo += L"\r\n";
    }
    
    // Filter out small/blurry images from SteamGridDB (web images are already sized correctly)
    std::vector<SteamGridDbImage> filteredImages;
    
    for (const auto& img : allImages) {
        if (img.width >= gConfig.minImageSize && img.height >= gConfig.minImageSize) {
            filteredImages.push_back(img);
        }
    }
    
    // Determine which image to use
    const SteamGridDbImage* selectedImg = nullptr;
    int selectedIdx = 0;
    
    DebugLog(L"[QuickApply] useWeb=" + std::to_wstring(gQuickApplyConfig.useWeb) + 
             L" webRandom=" + std::to_wstring(gQuickApplyConfig.webRandom) + 
             L" webResultIndex=" + std::to_wstring(gQuickApplyConfig.webResultIndex) +
             L" webSearchResults.size=" + std::to_wstring(webSearchResults.size()));
    
    debugInfo += L"\r\n[Selection Logic Check]\r\n";
    debugInfo += L"useWeb=" + std::to_wstring(gQuickApplyConfig.useWeb) + L"\r\n";
    debugInfo += L"webRandom=" + std::to_wstring(gQuickApplyConfig.webRandom) + L"\r\n";
    debugInfo += L"!webRandom=" + std::to_wstring(!gQuickApplyConfig.webRandom) + L"\r\n";
    debugInfo += L"webSearchResults.empty()=" + std::to_wstring(webSearchResults.empty()) + L"\r\n";
    debugInfo += L"!webSearchResults.empty()=" + std::to_wstring(!webSearchResults.empty()) + L"\r\n";
    debugInfo += L"Condition result=" + std::to_wstring(gQuickApplyConfig.useWeb && !gQuickApplyConfig.webRandom && !webSearchResults.empty()) + L"\r\n\r\n";
    
    if (gQuickApplyConfig.useWeb && !gQuickApplyConfig.webRandom && !webSearchResults.empty()) {
        // Use specific web result index
        int idx = gQuickApplyConfig.webResultIndex - 1;  // Convert to 0-based
        if (idx < 0) idx = 0;
        if (idx >= (int)webSearchResults.size()) idx = (int)webSearchResults.size() - 1;
        selectedImg = &webSearchResults[idx];
        selectedIdx = idx;
        DebugLog(L"[QuickApply] Using web result #" + std::to_wstring(gQuickApplyConfig.webResultIndex) + 
                 L" (index " + std::to_wstring(idx) + L") URL: " + selectedImg->url);
        
        debugInfo += L"Selection Mode: Specific web result #" + std::to_wstring(gQuickApplyConfig.webResultIndex) + L"\r\n";
        debugInfo += L"Selected: " + selectedImg->url + L"\r\n";
    } else {
        // Random selection from all filtered images (including web)
        filteredImages.insert(filteredImages.end(), webSearchResults.begin(), webSearchResults.end());
        
        if (filteredImages.empty()) {
            outStatus = L"No art found";
            debugInfo += L"\r\nResult: No art found\r\n";
            return false;
        }
        
        // Pick a random image using thread-safe random with atomic counter
        static std::atomic<unsigned int> sRandomCounter(0);
        unsigned int counter = sRandomCounter.fetch_add(1);
        unsigned int seed = (unsigned int)time(nullptr) ^ 
                            (unsigned int)GetCurrentThreadId() ^ 
                            (unsigned int)std::hash<std::wstring>{}(entry.idStr) ^
                            (counter * 2654435761u);  // Multiply by golden ratio prime for better distribution
        selectedIdx = seed % (int)filteredImages.size();
        selectedImg = &filteredImages[selectedIdx];
        
        debugInfo += L"Selection Mode: Random (seed=" + std::to_wstring(seed) + 
                     L", total=" + std::to_wstring(filteredImages.size()) + 
                     L", selected index=" + std::to_wstring(selectedIdx) + L")\r\n";
        debugInfo += L"Selected: " + selectedImg->url + L"\r\n";
    }
    
    if (!selectedImg) {
        outStatus = L"No art found";
        debugInfo += L"\r\nResult: No art found\r\n";
        return false;
    }
    
    const SteamGridDbImage& img = *selectedImg;

    // Download the image - use direct download (not cached) to ensure fresh image
    std::vector<unsigned char> data;
    if (!HttpDownloadBinary(img.url, data) || data.empty()) {
        outStatus = L"Download failed";
        debugInfo += L"\r\nResult: Download failed for " + img.url + L"\r\n";
        return false;
    }

    // Determine save path - use existing path if available
    fs::path savePath;
    std::wstring fileName;
    
    if (entry.hasArt && !entry.filePath.empty() && fs::exists(entry.filePath)) {
        // Use the existing file path
        savePath = entry.filePath;
        fileName = fs::path(entry.filePath).filename().wstring();
    } else {
        // Construct new path
        std::wstring basePath = GetThirdPartyLibrariesPath();
        fs::path storePath = fs::path(basePath) / entry.store;
        fileName = entry.expectedFileName;
        if (fileName.empty()) {
            fileName = ExpectedPngFromManifestId(entry.store, entry.idStr);
        }
        savePath = storePath / fileName;
        
        // Create directory if needed
        if (!fs::exists(storePath)) {
            fs::create_directories(storePath);
        }
    }

    // Backup original art if it exists (only first time)
    if (entry.hasArt && !entry.filePath.empty() && fs::exists(entry.filePath)) {
        BackupOriginalArt(entry.store, entry.filePath, fileName);
    }

    // Remove read-only attribute before writing
    std::wstring savePathStr = savePath.wstring();
    if (fs::exists(savePath)) {
        DWORD attrs = GetFileAttributesW(savePathStr.c_str());
        if (attrs != INVALID_FILE_ATTRIBUTES && (attrs & FILE_ATTRIBUTE_READONLY)) {
            SetFileAttributesW(savePathStr.c_str(), attrs & ~FILE_ATTRIBUTE_READONLY);
        }
    }

    // Save the image
    HANDLE hFile = CreateFileW(savePathStr.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) {
        DWORD err = GetLastError();
        outStatus = L"Save failed (err:" + std::to_wstring(err) + L")";
        debugInfo += L"\r\nResult: Save failed (error " + std::to_wstring(err) + L")\r\n";
        return false;
    }

    DWORD written = 0;
    BOOL writeOk = WriteFile(hFile, data.data(), (DWORD)data.size(), &written, nullptr);
    CloseHandle(hFile);
    
    if (!writeOk || written != data.size()) {
        outStatus = L"Write failed";
        debugInfo += L"\r\nResult: Write failed\r\n";
        return false;
    }

    // Set file as read-only
    SetFileReadOnly(savePath.wstring());

    // Update the entry
    entry.hasArt = true;
    entry.filePath = savePath.wstring();
    entry.fileName = fileName;

    // Include image index in status for debugging
    outStatus = L"Applied (#" + std::to_wstring(selectedIdx + 1) + L")";
    debugInfo += L"\r\nResult: Success! Saved to " + savePath.wstring() + L"\r\n";
    return true;
}

// Show debug info dialog for Quick Apply results
static void ShowQuickApplyDebugInfo(HWND hParent, const std::wstring& gameTitle, const std::wstring& debugText) {
    static bool debugDlgClassRegistered = false;
    static const wchar_t* DEBUG_DLG_CLASS = L"QuickApplyDebugClass";
    
    if (!debugDlgClassRegistered) {
        WNDCLASSEXW wc{};
        wc.cbSize = sizeof(wc);
        wc.lpfnWndProc = [](HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) -> LRESULT {
            static HWND gDebugEdit = nullptr;
            
            if (msg == WM_CREATE) {
                CREATESTRUCTW* cs = (CREATESTRUCTW*)lParam;
                std::wstring* pText = (std::wstring*)cs->lpCreateParams;
                
                // Create edit control
                gDebugEdit = CreateWindowExW(0, L"EDIT", pText->c_str(),
                    WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
                    10, 10, 760, 530, hWnd, nullptr, GetModuleHandle(nullptr), nullptr);
                if (gDebugEdit) {
                    SendMessage(gDebugEdit, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);
                    SendMessage(gDebugEdit, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, MAKELPARAM(5, 5));
                }
                
                // Create OK button
                HWND btnOk = CreateWindowExW(0, L"BUTTON", L"OK",
                    WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
                    350, 550, 80, 28, hWnd, (HMENU)IDOK, GetModuleHandle(nullptr), nullptr);
                if (btnOk) {
                    SendMessage(btnOk, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);
                }
                
                return 0;
            }
            
            if (msg == WM_CTLCOLORSTATIC || msg == WM_CTLCOLOREDIT) {
                if (gDarkMode) {
                    HDC hdc = (HDC)wParam;
                    SetTextColor(hdc, RGB(255, 255, 255));
                    SetBkColor(hdc, RGB(32, 32, 32));
                    return (LRESULT)GetStockObject(BLACK_BRUSH);
                }
            }
            
            if (msg == WM_COMMAND) {
                if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) {
                    PostQuitMessage(0);  // Exit the modal loop
                    return 0;
                }
            }
            
            if (msg == WM_CLOSE) {
                PostQuitMessage(0);  // Exit the modal loop
                return 0;
            }
            
            if (msg == WM_DESTROY) {
                return 0;
            }
            
            return DefWindowProcW(hWnd, msg, wParam, lParam);
        };
        wc.hInstance = GetModuleHandle(nullptr);
        wc.lpszClassName = DEBUG_DLG_CLASS;
        wc.hbrBackground = gDarkMode ? CreateSolidBrush(RGB(32, 32, 32)) : (HBRUSH)(COLOR_WINDOW + 1);
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        RegisterClassExW(&wc);
        debugDlgClassRegistered = true;
    }
    
    std::wstring title = L"Quick Apply Debug: " + gameTitle;
    
    HWND hDlg = CreateWindowExW(0, DEBUG_DLG_CLASS, title.c_str(),
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        CW_USEDEFAULT, CW_USEDEFAULT, 800, 630,
        hParent, nullptr, GetModuleHandle(nullptr), (LPVOID)&debugText);
    
    if (hDlg) {
        // Center on parent
        RECT rcParent, rcDlg;
        GetWindowRect(hParent, &rcParent);
        GetWindowRect(hDlg, &rcDlg);
        int x = rcParent.left + (rcParent.right - rcParent.left - (rcDlg.right - rcDlg.left)) / 2;
        int y = rcParent.top + (rcParent.bottom - rcParent.top - (rcDlg.bottom - rcDlg.top)) / 2;
        SetWindowPos(hDlg, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
        
        // Modal message loop
        MSG msg;
        while (GetMessageW(&msg, nullptr, 0, 0)) {
            if (msg.message == WM_QUIT) {
                // Don't propagate WM_QUIT - just exit our modal loop
                break;
            }
            if (!IsDialogMessage(hDlg, &msg)) {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
            if (!IsWindow(hDlg)) break;
        }
        
        // Clean up window if still exists
        if (IsWindow(hDlg)) {
            DestroyWindow(hDlg);
        }
    }
}

// Show Quick Apply configuration dialog
static bool ShowQuickApplyConfigDialog(HWND hParent) {
    static bool configDlgClassRegistered = false;
    static const wchar_t* CONFIG_DLG_CLASS = L"QuickApplyConfigClass";
    
    static HWND gQACfgWebCheck = nullptr;
    static HWND gQACfgWebResCombo = nullptr;
    static HWND gQACfgWebRandomRadio = nullptr;
    static HWND gQACfgWebIndexRadio = nullptr;
    static HWND gQACfgWebIndexEdit = nullptr;
    static HWND gQACfgGridsCheck = nullptr;
    static HWND gQACfgHeroesCheck = nullptr;
    static HWND gQACfgLogosCheck = nullptr;
    static HWND gQACfgIconsCheck = nullptr;
    static HWND gQACfgApiWarning = nullptr;
    static bool dialogResult = false;
    
    if (!configDlgClassRegistered) {
        WNDCLASSEXW wc = { sizeof(WNDCLASSEXW) };
        wc.lpfnWndProc = [](HWND hWnd, UINT msg, WPARAM w, LPARAM l) -> LRESULT {
            switch (msg) {
            case WM_CREATE: {
                // Title
                CreateWindowW(L"STATIC", L"Select art sources for Quick Apply:",
                    WS_CHILD | WS_VISIBLE,
                    20, 20, 340, 20,
                    hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                
                // Web checkbox
                gQACfgWebCheck = CreateWindowW(L"BUTTON", L"Web Search",
                    WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                    40, 50, 120, 20,
                    hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                SendMessageW(gQACfgWebCheck, BM_SETCHECK, gQuickApplyConfig.useWeb ? BST_CHECKED : BST_UNCHECKED, 0);
                
                // Web resolution label and combo
                CreateWindowW(L"STATIC", L"Size:",
                    WS_CHILD | WS_VISIBLE,
                    180, 52, 35, 20,
                    hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                
                gQACfgWebResCombo = CreateWindowW(WC_COMBOBOXW, L"",
                    WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
                    220, 48, 100, 200,
                    hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                
                // Apply same styling as main window comboboxes
                LONG_PTR exStyle = GetWindowLongPtrW(gQACfgWebResCombo, GWL_EXSTYLE);
                if (exStyle & WS_EX_CLIENTEDGE) {
                    SetWindowLongPtrW(gQACfgWebResCombo, GWL_EXSTYLE, exStyle & ~WS_EX_CLIENTEDGE);
                    SetWindowPos(gQACfgWebResCombo, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
                }
                SetWindowTheme(gQACfgWebResCombo, L"", L"");
                WNDPROC oldProc = (WNDPROC)GetWindowLongPtrW(gQACfgWebResCombo, GWLP_WNDPROC);
                if (oldProc != ComboSubclassProc) {
                    if (!gOriginalComboProc) gOriginalComboProc = oldProc;
                    SetWindowLongPtrW(gQACfgWebResCombo, GWLP_WNDPROC, (LONG_PTR)ComboSubclassProc);
                }
                
                SendMessageW(gQACfgWebResCombo, CB_ADDSTRING, 0, (LPARAM)L"265x265");
                SendMessageW(gQACfgWebResCombo, CB_ADDSTRING, 0, (LPARAM)L"512x512");
                SendMessageW(gQACfgWebResCombo, CB_ADDSTRING, 0, (LPARAM)L"1024x1024");
                SendMessageW(gQACfgWebResCombo, CB_ADDSTRING, 0, (LPARAM)L"2160x2160");
                SendMessageW(gQACfgWebResCombo, CB_ADDSTRING, 0, (LPARAM)L"4096x4096");
                // Select current resolution
                int selIdx = 2; // default 1024
                if (gQuickApplyConfig.webResolution == L"265x265") selIdx = 0;
                else if (gQuickApplyConfig.webResolution == L"512x512") selIdx = 1;
                else if (gQuickApplyConfig.webResolution == L"2160x2160") selIdx = 3;
                else if (gQuickApplyConfig.webResolution == L"4096x4096") selIdx = 4;
                SendMessageW(gQACfgWebResCombo, CB_SETCURSEL, selIdx, 0);
                
                // Web result selection - Radio buttons (without text)
                // "Choose the result #" option first
                gQACfgWebIndexRadio = CreateWindowW(L"BUTTON", L"",
                    WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON | WS_GROUP,
                    60, 85, 16, 16,
                    hWnd, (HMENU)1002, GetModuleHandleW(nullptr), nullptr);
                
                // Separate label for "Choose the result # (Recommended)"
                CreateWindowW(L"STATIC", L"Choose the result # (Recommended):",
                    WS_CHILD | WS_VISIBLE,
                    80, 85, 260, 20,
                    hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                
                // Random option below
                gQACfgWebRandomRadio = CreateWindowW(L"BUTTON", L"",
                    WS_CHILD | WS_VISIBLE | BS_AUTORADIOBUTTON,
                    60, 110, 16, 16,
                    hWnd, (HMENU)1001, GetModuleHandleW(nullptr), nullptr);
                
                // Separate label for Random
                CreateWindowW(L"STATIC", L"Random",
                    WS_CHILD | WS_VISIBLE,
                    80, 110, 60, 20,
                    hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                
                // Set the checked state AFTER creating both buttons
                if (gQuickApplyConfig.webRandom) {
                    SendMessageW(gQACfgWebRandomRadio, BM_SETCHECK, BST_CHECKED, 0);
                } else {
                    SendMessageW(gQACfgWebIndexRadio, BM_SETCHECK, BST_CHECKED, 0);
                }
                
                gQACfgWebIndexEdit = CreateWindowW(L"EDIT", L"",
                    WS_CHILD | WS_VISIBLE | WS_BORDER | ES_NUMBER,
                    350, 83, 40, 20,
                    hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                wchar_t indexStr[10];
                swprintf(indexStr, 10, L"%d", gQuickApplyConfig.webResultIndex);
                SetWindowTextW(gQACfgWebIndexEdit, indexStr);
                
                // Enable/disable the edit box based on which radio is selected
                EnableWindow(gQACfgWebIndexEdit, !gQuickApplyConfig.webRandom);
                
                // Web warning note
                CreateWindowW(L"STATIC", L"Note: Web results are not curated and may be wrong",
                    WS_CHILD | WS_VISIBLE,
                    40, 135, 380, 35,
                    hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                
                // API key warning (shown when no API key is set)
                bool hasApiKey = !gConfig.steamGridDbKey.empty();
                gQACfgApiWarning = CreateWindowW(L"STATIC", L"To use steamgrid go to 'Config' menu and put the API key there",
                    WS_CHILD | (hasApiKey ? 0 : WS_VISIBLE) | SS_LEFT,
                    40, 170, 380, 35,
                    hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                
                // Grids checkbox
                gQACfgGridsCheck = CreateWindowW(L"BUTTON", L"Grids (SteamGridDB)",
                    WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                    40, 210, 180, 20,
                    hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                SendMessageW(gQACfgGridsCheck, BM_SETCHECK, gQuickApplyConfig.useGrids ? BST_CHECKED : BST_UNCHECKED, 0);
                EnableWindow(gQACfgGridsCheck, hasApiKey);
                
                // Heroes checkbox
                gQACfgHeroesCheck = CreateWindowW(L"BUTTON", L"Heroes (SteamGridDB)",
                    WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                    40, 240, 180, 20,
                    hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                SendMessageW(gQACfgHeroesCheck, BM_SETCHECK, gQuickApplyConfig.useHeroes ? BST_CHECKED : BST_UNCHECKED, 0);
                EnableWindow(gQACfgHeroesCheck, hasApiKey);
                
                // Logos checkbox
                gQACfgLogosCheck = CreateWindowW(L"BUTTON", L"Logos (SteamGridDB)",
                    WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                    40, 270, 180, 20,
                    hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                SendMessageW(gQACfgLogosCheck, BM_SETCHECK, gQuickApplyConfig.useLogos ? BST_CHECKED : BST_UNCHECKED, 0);
                EnableWindow(gQACfgLogosCheck, hasApiKey);
                
                // Icons checkbox
                gQACfgIconsCheck = CreateWindowW(L"BUTTON", L"Icons (SteamGridDB)",
                    WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
                    40, 300, 180, 20,
                    hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                SendMessageW(gQACfgIconsCheck, BM_SETCHECK, gQuickApplyConfig.useIcons ? BST_CHECKED : BST_UNCHECKED, 0);
                EnableWindow(gQACfgIconsCheck, hasApiKey);
                
                // OK and Cancel buttons
                CreateWindowW(L"BUTTON", L"OK",
                    WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
                    120, 340, 80, 30,
                    hWnd, (HMENU)IDOK, GetModuleHandleW(nullptr), nullptr);
                    
                CreateWindowW(L"BUTTON", L"Cancel",
                    WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
                    210, 340, 80, 30,
                    hWnd, (HMENU)IDCANCEL, GetModuleHandleW(nullptr), nullptr);
                
                // Apply dark mode theming to checkboxes and radio buttons
                if (gDarkMode) {
                    SetWindowTheme(gQACfgWebCheck, L"DarkMode_Explorer", nullptr);
                    SetWindowTheme(gQACfgWebRandomRadio, L"DarkMode_Explorer", nullptr);
                    SetWindowTheme(gQACfgWebIndexRadio, L"DarkMode_Explorer", nullptr);
                    SetWindowTheme(gQACfgGridsCheck, L"DarkMode_Explorer", nullptr);
                    SetWindowTheme(gQACfgHeroesCheck, L"DarkMode_Explorer", nullptr);
                    SetWindowTheme(gQACfgLogosCheck, L"DarkMode_Explorer", nullptr);
                    SetWindowTheme(gQACfgIconsCheck, L"DarkMode_Explorer", nullptr);
                }
                
                // Apply dark mode to window
                if (gDarkMode) {
                    BOOL dark = TRUE;
                    DwmSetWindowAttribute(hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &dark, sizeof(dark));
                } else {
                    BOOL dark = FALSE;
                    DwmSetWindowAttribute(hWnd, DWMWA_USE_IMMERSIVE_DARK_MODE, &dark, sizeof(dark));
                }
                return 0;
            }
            case WM_DRAWITEM: {
                DRAWITEMSTRUCT* dis = (DRAWITEMSTRUCT*)l;
                if (dis->CtlType == ODT_BUTTON) {
                    DrawDarkButton(dis);
                    return TRUE;
                }
                if (dis->CtlType == ODT_COMBOBOX) {
                    DebugLog(L"[QADialog] WM_DRAWITEM for combobox, itemID=" + std::to_wstring(dis->itemID) + 
                             L" itemState=" + std::to_wstring(dis->itemState) + 
                             L" hwndItem=" + std::to_wstring((UINT_PTR)dis->hwndItem));
                    DrawDarkComboItem(dis);
                    return TRUE;
                }
                break;
            }
            case WM_CTLCOLORSTATIC:
            case WM_CTLCOLOREDIT:
                if (gDarkMode) {
                    HDC hdc = (HDC)w;
                    HWND hControl = (HWND)l;
                    
                    // Check if this is the API warning label - paint it red
                    if (hControl == gQACfgApiWarning) {
                        SetTextColor(hdc, RGB(255, 80, 80)); // Red text
                        SetBkColor(hdc, RGB(32, 32, 32));
                        SetBkMode(hdc, TRANSPARENT);
                        if (!gDarkBgBrush) gDarkBgBrush = CreateSolidBrush(RGB(32, 32, 32));
                        return (LRESULT)gDarkBgBrush;
                    }
                    
                    SetTextColor(hdc, RGB(255, 255, 255));
                    SetBkColor(hdc, RGB(32, 32, 32));
                    SetBkMode(hdc, TRANSPARENT);
                    if (!gDarkBgBrush) gDarkBgBrush = CreateSolidBrush(RGB(32, 32, 32));
                    return (LRESULT)gDarkBgBrush;
                } else {
                    HDC hdc = (HDC)w;
                    HWND hControl = (HWND)l;
                    
                    // Check if this is the API warning label - paint it red
                    if (hControl == gQACfgApiWarning) {
                        SetTextColor(hdc, RGB(220, 0, 0)); // Dark red text
                        SetBkColor(hdc, GetSysColor(COLOR_WINDOW));
                        return (LRESULT)GetSysColorBrush(COLOR_WINDOW);
                    }
                    
                    SetTextColor(hdc, RGB(0, 0, 0));
                    SetBkColor(hdc, GetSysColor(COLOR_WINDOW));
                    return (LRESULT)GetSysColorBrush(COLOR_WINDOW);
                }
                break;
            case WM_CTLCOLORBTN:
                if (gDarkMode) {
                    HDC hdc = (HDC)w;
                    SetTextColor(hdc, RGB(255, 255, 255));
                    SetBkMode(hdc, TRANSPARENT);
                    if (!gDarkBgBrush) gDarkBgBrush = CreateSolidBrush(RGB(32, 32, 32));
                    return (LRESULT)gDarkBgBrush;
                } else {
                    return (LRESULT)GetSysColorBrush(COLOR_BTNFACE);
                }
                break;
            case WM_CTLCOLORLISTBOX:
                // For combobox dropdown list
                DebugLog(L"[QADialog] WM_CTLCOLORLISTBOX called, hdc=" + std::to_wstring((UINT_PTR)w) + 
                         L" hwnd=" + std::to_wstring((UINT_PTR)l));
                if (gDarkMode) {
                    HDC hdc = (HDC)w;
                    SetTextColor(hdc, RGB(255, 255, 255));
                    SetBkColor(hdc, RGB(45, 45, 45));
                    if (!gDarkEditBrush) gDarkEditBrush = CreateSolidBrush(RGB(45, 45, 45));
                    return (LRESULT)gDarkEditBrush;
                } else {
                    HDC hdc = (HDC)w;
                    SetTextColor(hdc, RGB(0, 0, 0));
                    SetBkColor(hdc, RGB(255, 255, 255));
                    if (!gLightBgBrush) gLightBgBrush = CreateSolidBrush(RGB(255, 255, 255));
                    return (LRESULT)gLightBgBrush;
                }
            case WM_ERASEBKGND: {
                HDC hdc = (HDC)w;
                RECT rc;
                GetClientRect(hWnd, &rc);
                if (gDarkMode) {
                    if (!gDarkBgBrush) gDarkBgBrush = CreateSolidBrush(RGB(32, 32, 32));
                    FillRect(hdc, &rc, gDarkBgBrush);
                } else {
                    FillRect(hdc, &rc, GetSysColorBrush(COLOR_WINDOW));
                }
                return 1;
            }
            case WM_COMMAND: {
                int id = LOWORD(w);
                int notif = HIWORD(w);
                
                // Handle radio button clicks
                if (id == 1001 && notif == BN_CLICKED) {  // Random radio
                    SendMessageW(gQACfgWebRandomRadio, BM_SETCHECK, BST_CHECKED, 0);
                    SendMessageW(gQACfgWebIndexRadio, BM_SETCHECK, BST_UNCHECKED, 0);
                    EnableWindow(gQACfgWebIndexEdit, FALSE);  // Disable edit box
                }
                else if (id == 1002 && notif == BN_CLICKED) {  // Result # radio
                    SendMessageW(gQACfgWebIndexRadio, BM_SETCHECK, BST_CHECKED, 0);
                    SendMessageW(gQACfgWebRandomRadio, BM_SETCHECK, BST_UNCHECKED, 0);
                    EnableWindow(gQACfgWebIndexEdit, TRUE);  // Enable edit box
                }
                // Auto-select "Result #:" radio when user clicks or types in the edit box
                else if ((HWND)l == gQACfgWebIndexEdit && (notif == EN_SETFOCUS || notif == EN_CHANGE)) {
                    SendMessageW(gQACfgWebIndexRadio, BM_SETCHECK, BST_CHECKED, 0);
                    SendMessageW(gQACfgWebRandomRadio, BM_SETCHECK, BST_UNCHECKED, 0);
                    EnableWindow(gQACfgWebIndexEdit, TRUE);  // Ensure it's enabled
                }
                
                if (id == IDOK) {
                    // Save configuration
                    gQuickApplyConfig.useWeb = (SendMessageW(gQACfgWebCheck, BM_GETCHECK, 0, 0) == BST_CHECKED);
                    int resIdx = (int)SendMessageW(gQACfgWebResCombo, CB_GETCURSEL, 0, 0);
                    const wchar_t* resolutions[] = { L"265x265", L"512x512", L"1024x1024", L"2160x2160", L"4096x4096" };
                    if (resIdx >= 0 && resIdx < 5) {
                        gQuickApplyConfig.webResolution = resolutions[resIdx];
                    }
                    
                    DebugLog(L"[ConfigDialog] RandomRadio HWND: " + std::to_wstring((UINT_PTR)gQACfgWebRandomRadio));
                    DebugLog(L"[ConfigDialog] IndexRadio HWND: " + std::to_wstring((UINT_PTR)gQACfgWebIndexRadio));
                    DebugLog(L"[ConfigDialog] Are they valid? Random=" + std::to_wstring(IsWindow(gQACfgWebRandomRadio)) + 
                             L" Index=" + std::to_wstring(IsWindow(gQACfgWebIndexRadio)));
                    
                    LRESULT randomChecked = SendMessageW(gQACfgWebRandomRadio, BM_GETCHECK, 0, 0);
                    LRESULT indexChecked = SendMessageW(gQACfgWebIndexRadio, BM_GETCHECK, 0, 0);
                    DebugLog(L"[ConfigDialog] Random radio state: " + std::to_wstring(randomChecked) + L" (BST_CHECKED=" + std::to_wstring(BST_CHECKED) + L")");
                    DebugLog(L"[ConfigDialog] Index radio state: " + std::to_wstring(indexChecked) + L" (BST_CHECKED=" + std::to_wstring(BST_CHECKED) + L")");
                    DebugLog(L"[ConfigDialog] Setting webRandom = " + std::to_wstring(randomChecked == BST_CHECKED));
                    gQuickApplyConfig.webRandom = (randomChecked == BST_CHECKED);
                    
                    // Get web result index
                    wchar_t indexText[10];
                    GetWindowTextW(gQACfgWebIndexEdit, indexText, 10);
                    int index = _wtoi(indexText);
                    if (index < 1) index = 1;
                    gQuickApplyConfig.webResultIndex = index;
                    
                    gQuickApplyConfig.useGrids = (SendMessageW(gQACfgGridsCheck, BM_GETCHECK, 0, 0) == BST_CHECKED);
                    gQuickApplyConfig.useHeroes = (SendMessageW(gQACfgHeroesCheck, BM_GETCHECK, 0, 0) == BST_CHECKED);
                    gQuickApplyConfig.useLogos = (SendMessageW(gQACfgLogosCheck, BM_GETCHECK, 0, 0) == BST_CHECKED);
                    gQuickApplyConfig.useIcons = (SendMessageW(gQACfgIconsCheck, BM_GETCHECK, 0, 0) == BST_CHECKED);
                    
                    DebugLog(L"[ConfigDialog] FINAL CONFIG: webRandom=" + std::to_wstring(gQuickApplyConfig.webRandom) + 
                             L" webResultIndex=" + std::to_wstring(gQuickApplyConfig.webResultIndex));
                    
                    dialogResult = true;
                    DestroyWindow(hWnd);
                    return 0;
                }
                if (id == IDCANCEL) {
                    dialogResult = false;
                    DestroyWindow(hWnd);
                    return 0;
                }
                break;
            }
            case WM_CLOSE:
                dialogResult = false;
                DestroyWindow(hWnd);
                return 0;
            }
            return DefWindowProcW(hWnd, msg, w, l);
        };
        wc.hInstance = GetModuleHandleW(nullptr);
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hbrBackground = gDarkMode ? gDarkBgBrush : (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = CONFIG_DLG_CLASS;
        RegisterClassExW(&wc);
        configDlgClassRegistered = true;
    }
    
    int dlgWidth = 460;
    int dlgHeight = 420;
    RECT workArea;
    SystemParametersInfo(SPI_GETWORKAREA, 0, &workArea, 0);
    int x = workArea.left + ((workArea.right - workArea.left) - dlgWidth) / 2;
    int y = workArea.top + ((workArea.bottom - workArea.top) - dlgHeight) / 2;
    
    HWND hConfigDlg = CreateWindowExW(
        WS_EX_DLGMODALFRAME | WS_EX_TOPMOST,
        CONFIG_DLG_CLASS, L"Quick Apply Settings",
        WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        x, y, dlgWidth, dlgHeight,
        hParent, nullptr, GetModuleHandleW(nullptr), nullptr);
    
    if (!hConfigDlg) return false;
    
    // Disable parent window
    if (hParent) EnableWindow(hParent, FALSE);
    
    // Message loop
    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        if (!IsWindow(hConfigDlg)) break;
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    
    // Re-enable parent
    if (hParent) {
        EnableWindow(hParent, TRUE);
        SetForegroundWindow(hParent);
    }
    
    return dialogResult;
}

// Quick apply art to multiple selected games (multi-threaded)
static void DoQuickApply(HWND hWnd) {
    if (!gList) return;
    if (gQuickApplyRunning) return; // Already running

    // Get selected items
    std::vector<int> selectedIndices;
    int idx = -1;
    while ((idx = ListView_GetNextItem(gList, idx, LVNI_SELECTED)) != -1) {
        selectedIndices.push_back(idx);
    }

    if (selectedIndices.empty()) {
        DarkMessageBox(hWnd, 
            L"No games selected.\n\nPlease select one or more games in the list, then click Quick Apply.",
            L"No Selection", MB_OK | MB_ICONINFORMATION);
        return;
    }

    // Show confirmation only the first time
    if (!gQuickApplyConfig.hasShownQuickApplyIntro) {
        wchar_t confirmMsg[512];
        swprintf(confirmMsg, 512, L"Apply art to %d selected game(s)?\n\nYou can choose which sources to use (Web, Grids, Heroes, Logos, Icons) in the next window.", 
            (int)selectedIndices.size());
        if (DarkMessageBox(hWnd, confirmMsg, L"Quick Apply", MB_YESNO | MB_ICONQUESTION) != IDYES) {
            return;
        }
        gQuickApplyConfig.hasShownQuickApplyIntro = true;
    }
    
    // Show configuration dialog
    if (!ShowQuickApplyConfigDialog(hWnd)) {
        return; // User cancelled
    }
    
    // Check that at least one source is enabled
    if (!gQuickApplyConfig.useWeb && !gQuickApplyConfig.useGrids && !gQuickApplyConfig.useHeroes && 
        !gQuickApplyConfig.useLogos && !gQuickApplyConfig.useIcons) {
        DarkMessageBox(hWnd, L"Please select at least one art source.", L"No Sources Selected", MB_OK | MB_ICONWARNING);
        return;
    }

    // Mark as running - disables filter controls
    gQuickApplyRunning = true;

    // Disable filter controls during processing
    if (gCmbStore) EnableWindow(gCmbStore, FALSE);
    if (gCmbArtFilter) EnableWindow(gCmbArtFilter, FALSE);
    if (gSearchEdit) EnableWindow(gSearchEdit, FALSE);
    if (gBtnQuickApply) EnableWindow(gBtnQuickApply, FALSE);
    if (gBtnScan) EnableWindow(gBtnScan, FALSE);

    // Seed random number generator
    srand((unsigned int)time(nullptr));

    // Build list of games to process
    struct QuickApplyTask {
        int listIdx;
        int dataIdx;
        GameEntry* originalEntry;
        GameEntry* filteredEntry;
        std::wstring title;
    };
    std::vector<QuickApplyTask> tasks;
    
    for (int listIdx : selectedIndices) {
        LVITEM lvi{};
        lvi.mask = LVIF_PARAM;
        lvi.iItem = listIdx;
        ListView_GetItem(gList, &lvi);
        int dataIdx = (int)lvi.lParam;
        
        if (dataIdx < 0 || dataIdx >= (int)gFilteredItems.size()) continue;
        
        GameEntry& filteredEntry = gFilteredItems[dataIdx];
        GameEntry* originalEntry = nullptr;
        for (auto& e : gLastItems) {
            if (e.store == filteredEntry.store && e.idStr == filteredEntry.idStr) {
                originalEntry = &e;
                break;
            }
        }
        if (!originalEntry) continue;
        
        QuickApplyTask task;
        task.listIdx = listIdx;
        task.dataIdx = dataIdx;
        task.originalEntry = originalEntry;
        task.filteredEntry = &filteredEntry;
        task.title = originalEntry->title.empty() ? originalEntry->idStr : originalEntry->title;
        tasks.push_back(task);
    }
    
    if (tasks.empty()) {
        DarkMessageBox(hWnd, L"No valid games to process.", L"Quick Apply", MB_OK | MB_ICONWARNING);
        return;
    }

    // Results tracking
    struct QuickApplyResult {
        size_t taskIdx;
        bool success;
        std::wstring title;
        std::wstring status;
        std::wstring debugInfo;
    };
    std::mutex resultsMutex;
    std::vector<QuickApplyResult> results;
    std::atomic<size_t> nextTask(0);
    std::atomic<size_t> completedCount(0);
    
    // Determine parallel count
    int parallelCount = gConfig.parallelDownloads;
    if (parallelCount <= 0) {
        parallelCount = (int)std::thread::hardware_concurrency();
        if (parallelCount < 2) parallelCount = 4;
    }
    // Limit to reasonable number for API calls
    if (parallelCount > 8) parallelCount = 8;
    
    // Copy config for thread safety
    std::wstring apiKey = gConfig.steamGridDbKey;
    
    // Worker function
    auto workerFunc = [&]() {
        while (true) {
            size_t taskIdx = nextTask.fetch_add(1);
            if (taskIdx >= tasks.size()) break;
            
            QuickApplyTask& task = tasks[taskIdx];
            GameEntry entryCopy = *task.originalEntry; // Work on a copy for thread safety
            
            std::wstring status;
            std::wstring debugInfo;
            bool success = QuickApplyArtToGame(entryCopy, status, debugInfo);
            
            // Store result
            {
                std::lock_guard<std::mutex> lock(resultsMutex);
                QuickApplyResult res;
                res.taskIdx = taskIdx;
                res.success = success;
                res.title = task.title;
                res.status = status;
                res.debugInfo = debugInfo;
                results.push_back(res);
                
                // Update original entry if successful (protected by mutex)
                if (success) {
                    task.originalEntry->hasArt = entryCopy.hasArt;
                    task.originalEntry->filePath = entryCopy.filePath;
                    task.originalEntry->fileName = entryCopy.fileName;
                    task.filteredEntry->hasArt = entryCopy.hasArt;
                    task.filteredEntry->filePath = entryCopy.filePath;
                    task.filteredEntry->fileName = entryCopy.fileName;
                }
            }
            completedCount.fetch_add(1);
        }
    };
    
    // Start worker threads
    std::vector<std::thread> workers;
    int numThreads = (tasks.size() < (size_t)parallelCount) ? (int)tasks.size() : parallelCount;
    for (int t = 0; t < numThreads; ++t) {
        workers.emplace_back(workerFunc);
    }
    
    // Process messages on main thread while workers are running
    size_t totalTasks = tasks.size();
    while (completedCount.load() < totalTasks) {
        MSG msg;
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
        
        size_t done = completedCount.load();
        SetStatusText(L"Quick Apply: Processing " + std::to_wstring(done) + L" of " + std::to_wstring(totalTasks) + L"...");
        Sleep(50);
    }
    
    // Wait for all threads to complete
    for (auto& t : workers) {
        if (t.joinable()) t.join();
    }

    // Save cache
    SaveUiStateToCacheFile();

    // Refresh the list view
    DestroyImageList();
    EnsureImageList();
    PopulateListFromItems(gLastItems);

    // Build results lists
    std::vector<std::wstring> successTitles;
    std::vector<std::wstring> failedTitles;
    for (const auto& res : results) {
        if (res.success) {
            successTitles.push_back(res.title + L" " + res.status);  // Include status like "Applied (#3 of 15)"
        } else {
            failedTitles.push_back(res.title + L" (" + res.status + L")");
        }
    }
    
    // Build result message
    std::wstring resultText;
    resultText += L"Quick Apply Complete!\r\n";
    resultText += L"======================================\r\n\r\n";
    
    resultText += L"[OK] Successfully Applied: " + std::to_wstring(successTitles.size()) + L"\r\n";
    if (!successTitles.empty()) {
        resultText += L"--------------------------------------\r\n";
        for (const auto& title : successTitles) {
            resultText += L"  + " + title + L"\r\n";
        }
    }
    resultText += L"\r\n";
    
    resultText += L"[X] Failed: " + std::to_wstring(failedTitles.size()) + L"\r\n";
    if (!failedTitles.empty()) {
        resultText += L"--------------------------------------\r\n";
        for (const auto& title : failedTitles) {
            resultText += L"  - " + title + L"\r\n";
        }
    }

    // Show results in a dialog window using a proper window class
    static bool resultWndClassRegistered = false;
    static const wchar_t* RESULT_WND_CLASS = L"QuickApplyResultWnd";
    
    if (!resultWndClassRegistered) {
        WNDCLASSEXW wc{};
        wc.cbSize = sizeof(wc);
        wc.style = CS_HREDRAW | CS_VREDRAW;
        wc.lpfnWndProc = [](HWND hWnd, UINT msg, WPARAM w, LPARAM l) -> LRESULT {
            static HWND sEdit = nullptr;
            static HWND sOkBtn = nullptr;
            
            switch (msg) {
            case WM_CREATE: {
                CREATESTRUCTW* cs = (CREATESTRUCTW*)l;
                std::wstring* pText = (std::wstring*)cs->lpCreateParams;
                
                RECT rc;
                GetClientRect(hWnd, &rc);
                int cw = rc.right - rc.left;
                int ch = rc.bottom - rc.top;
                
                // Create multiline edit for results
                sEdit = CreateWindowExW(
                    0,
                    L"EDIT", pText->c_str(),
                    WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
                    10, 10, cw - 20, ch - 60,
                    hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
                
                // Create OK button with owner-draw for dark mode
                sOkBtn = CreateWindowW(
                    L"BUTTON", L"OK",
                    WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
                    (cw - 80) / 2, ch - 45, 80, 30,
                    hWnd, (HMENU)IDOK, GetModuleHandleW(nullptr), nullptr);
                
                // Set font
                HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
                SendMessageW(sEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
                SendMessageW(sOkBtn, WM_SETFONT, (WPARAM)hFont, TRUE);
                
                // Apply dark mode to edit
                if (gDarkMode) {
                    SetWindowTheme(sEdit, L"DarkMode_Explorer", nullptr);
                }
                return 0;
            }
            case WM_DRAWITEM: {
                DRAWITEMSTRUCT* dis = (DRAWITEMSTRUCT*)l;
                if (dis->CtlType == ODT_BUTTON) {
                    DrawDarkButton(dis);
                    return TRUE;
                }
                break;
            }
            case WM_CTLCOLOREDIT:
            case WM_CTLCOLORSTATIC:
                if (gDarkMode) {
                    HDC hdc = (HDC)w;
                    SetTextColor(hdc, RGB(255, 255, 255));
                    SetBkColor(hdc, RGB(45, 45, 45));
                    if (!gDarkEditBrush) gDarkEditBrush = CreateSolidBrush(RGB(45, 45, 45));
                    return (LRESULT)gDarkEditBrush;
                }
                break;
            case WM_CTLCOLORBTN:
                if (gDarkMode) {
                    if (!gDarkBgBrush) gDarkBgBrush = CreateSolidBrush(RGB(32, 32, 32));
                    return (LRESULT)gDarkBgBrush;
                }
                break;
            case WM_ERASEBKGND:
                if (gDarkMode) {
                    HDC hdc = (HDC)w;
                    RECT rc;
                    GetClientRect(hWnd, &rc);
                    if (!gDarkBgBrush) gDarkBgBrush = CreateSolidBrush(RGB(32, 32, 32));
                    FillRect(hdc, &rc, gDarkBgBrush);
                    return 1;
                }
                break;
            case WM_COMMAND:
                if (LOWORD(w) == IDOK || HIWORD(w) == BN_CLICKED) {
                    PostMessageW(hWnd, WM_CLOSE, 0, 0);
                    return 0;
                }
                break;
            case WM_CLOSE:
                DestroyWindow(hWnd);
                return 0;
            case WM_DESTROY:
                PostQuitMessage(0);
                return 0;
            }
            return DefWindowProcW(hWnd, msg, w, l);
        };
        wc.hInstance = GetModuleHandleW(nullptr);
        wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
        wc.hbrBackground = gDarkMode ? gDarkBgBrush : (HBRUSH)(COLOR_WINDOW + 1);
        wc.lpszClassName = RESULT_WND_CLASS;
        RegisterClassExW(&wc);
        resultWndClassRegistered = true;
    }
    
    int dlgWidth = 500;
    int dlgHeight = 400;
    RECT workArea;
    SystemParametersInfo(SPI_GETWORKAREA, 0, &workArea, 0);
    int x = workArea.left + ((workArea.right - workArea.left) - dlgWidth) / 2;
    int y = workArea.top + ((workArea.bottom - workArea.top) - dlgHeight) / 2;
    
    HWND hResultWnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        RESULT_WND_CLASS, L"Quick Apply Results",
        WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        x, y, dlgWidth, dlgHeight,
        hWnd, nullptr, GetModuleHandleW(nullptr), (LPVOID)&resultText);
    
    if (hResultWnd) {
        // Apply dark mode to title bar
        ApplyDarkModeToWindow(hResultWnd);
        
        // Modal message loop
        EnableWindow(hWnd, FALSE);
        MSG msg;
        while (GetMessageW(&msg, nullptr, 0, 0)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
        EnableWindow(hWnd, TRUE);
        SetForegroundWindow(hWnd);
    }
    
    // Show debug info for single-game selections (disabled - enable for debugging)
    // if (results.size() == 1 && !results[0].debugInfo.empty()) {
    //     ShowQuickApplyDebugInfo(hWnd, results[0].title, results[0].debugInfo);
    // }

    // Re-enable controls
    gQuickApplyRunning = false;
    if (gCmbStore) EnableWindow(gCmbStore, TRUE);
    if (gCmbArtFilter) EnableWindow(gCmbArtFilter, TRUE);
    if (gSearchEdit) EnableWindow(gSearchEdit, TRUE);
    if (gBtnQuickApply) EnableWindow(gBtnQuickApply, TRUE);
    if (gBtnScan) EnableWindow(gBtnScan, TRUE);

    SetStatusText(L"Quick Apply complete - " + std::to_wstring(successTitles.size()) + L" succeeded, " + std::to_wstring(failedTitles.size()) + L" failed");
}

static void DoScanAndCache(HWND hWnd) {
    SetStatusText(L"Scanning for games...");
    ProcessPendingMessages();
    
    auto scanned = ScanCache();
    
    SetStatusText(L"Found " + std::to_wstring(scanned.size()) + L" games. Processing...");
    ProcessPendingMessages();

    std::unordered_map<std::wstring, std::wstring> localTitleMap;
    localTitleMap.reserve(gLastItems.size());
    for (auto& old : gLastItems) {
        if (!old.title.empty()) localTitleMap[old.idStr] = old.title;
    }

    for (auto& e : scanned) {
        auto it = localTitleMap.find(e.idStr);
        if (it != localTitleMap.end()) e.title = it->second;
    }

    // Count how many need resolution
    int needsResolution = 0;
    for (auto& e : scanned) {
        if (e.title.empty()) needsResolution++;
    }

    int resolvedNow = 0;
    
    if (needsResolution > 0) {
        SetStatusText(L"Resolving " + std::to_wstring(needsResolution) + L" titles in parallel...");
        ProcessPendingMessages();
        
        // Collect entries that need resolution
        std::vector<size_t> indicesToResolve;
        for (size_t i = 0; i < scanned.size(); i++) {
            if (scanned[i].title.empty()) {
                indicesToResolve.push_back(i);
            }
        }
        
        // Use parallel processing with thread pool
        std::atomic<int> processed(0);
        std::atomic<int> resolved(0);
        
        auto resolveTask = [&](size_t idx) {
            scanned[idx].title = ResolveTitleForEntry(scanned[idx]);
            if (!scanned[idx].title.empty()) {
                resolved++;
            }
            processed++;
        };
        
        // Process in parallel using thread pool
        std::vector<std::thread> threads;
        int threadCount = std::thread::hardware_concurrency();
        if (threadCount < 2) threadCount = 2;
        if (threadCount > 8) threadCount = 8;
        
        size_t itemsPerThread = (indicesToResolve.size() + threadCount - 1) / threadCount;
        
        for (int t = 0; t < threadCount; t++) {
            size_t start = t * itemsPerThread;
            if (start >= indicesToResolve.size()) break;
            
            size_t end = std::min(start + itemsPerThread, indicesToResolve.size());
            
            threads.emplace_back([&, start, end]() {
                for (size_t i = start; i < end; i++) {
                    resolveTask(indicesToResolve[i]);
                }
            });
        }
        
        // Monitor progress from main thread
        while (processed.load() < needsResolution) {
            int current = processed.load();
            SetStatusText(L"Resolving titles " + std::to_wstring(current) + L" of " + std::to_wstring(needsResolution) + L"...");
            ProcessPendingMessages();
            std::this_thread::sleep_for(std::chrono::milliseconds(100));
        }
        
        for (auto& thread : threads) {
            thread.join();
        }
        
        resolvedNow = resolved.load();
    }

    SetStatusText(L"Populating list...");
    ProcessPendingMessages();

    gLastItems = scanned;
    PopulateListFromItems(gLastItems);

    SetStatusText(L"Saving cache...");
    ProcessPendingMessages();

    CacheState st;
    st.iconSize = gIconSize;
    st.layout = gGridLayout ? L"grid" : L"list";
    st.config = gConfig;
    st.items = gLastItems;
    SaveCacheFile(st);


    // Count unresolved and missing art
    int unresolvedCount = 0, missingArtCount = 0;
    for (const auto& e : gLastItems) {
        if (e.title.empty()) unresolvedCount++;
        if (!e.hasArt) missingArtCount++;
    }
    std::wstring statusFinal = std::to_wstring(gLastItems.size()) + L" games";
    if (unresolvedCount > 0)
        statusFinal += L" | " + std::to_wstring(unresolvedCount) + L" unresolved";
    if (missingArtCount > 0)
        statusFinal += L" | " + std::to_wstring(missingArtCount) + L" missing art";
    statusFinal += L" | " + std::to_wstring(resolvedNow) + L" titles resolved";
    SetStatusText(statusFinal);


    std::wstring msg = L"Scan complete and cached.\n\n"
                       L"Games found: " + std::to_wstring(gLastItems.size()) + L"\n"
                       L"Titles resolved this run: " + std::to_wstring(resolvedNow) + L"\n\n"
                       L"Cache file:\n" + GetCacheFilePath();
    DarkMessageBox(hWnd, msg.c_str(), APP_TITLE, MB_OK | MB_ICONINFORMATION);
    
    // Update scan button to "Refresh" after first scan
    if (!gHasScanned) {
        gHasScanned = true;
        if (gBtnScan) {
            SetWindowTextW(gBtnScan, L"Refresh");
            InvalidateRect(gBtnScan, nullptr, TRUE);
        }
    }
}

// ------------------------------------------------------------
// UI: Combo helpers
// ------------------------------------------------------------

static void InitCombosAndDefaults() {
    SendMessageW(gCmbSize, CB_ADDSTRING, 0, (LPARAM)L"32");
    SendMessageW(gCmbSize, CB_ADDSTRING, 0, (LPARAM)L"64");
    SendMessageW(gCmbSize, CB_ADDSTRING, 0, (LPARAM)L"128");

    SendMessageW(gCmbLayout, CB_ADDSTRING, 0, (LPARAM)L"List");
    SendMessageW(gCmbLayout, CB_ADDSTRING, 0, (LPARAM)L"Grid");

    SendMessageW(gCmbSize, CB_SETCURSEL, 1, 0);
    SendMessageW(gCmbLayout, CB_SETCURSEL, 0, 0);
    gIconSize = 64;
    gGridLayout = false;
}

static void ApplyCombosFromState(const CacheState& st) {
    gIconSize = st.iconSize;
    int selSize = 1;
    if (gIconSize <= 32) selSize = 0;
    else if (gIconSize >= 128) selSize = 2;
    else selSize = 1;
    SendMessageW(gCmbSize, CB_SETCURSEL, selSize, 0);
    gIconSize = (selSize == 0 ? 32 : selSize == 2 ? 128 : 64);

    gGridLayout = (st.layout == L"grid");
    SendMessageW(gCmbLayout, CB_SETCURSEL, gGridLayout ? 1 : 0, 0);
}

static void SaveUiStateToCacheFile() {
    CacheState st;
    st.iconSize = gIconSize;
    st.layout = gGridLayout ? L"grid" : L"list";
    st.config = gConfig;
    st.items = gLastItems;
    SaveCacheFile(st);
}

// ------------------------------------------------------------
// SteamGridDB API functions
// ------------------------------------------------------------

static bool HttpDownloadBinary(const std::wstring& url, std::vector<unsigned char>& outData) {
    outData.clear();

    URL_COMPONENTS uc{};
    uc.dwStructSize = sizeof(uc);
    wchar_t host[256]{};
    wchar_t path[4096]{};
    uc.lpszHostName = host; uc.dwHostNameLength = _countof(host);
    uc.lpszUrlPath = path; uc.dwUrlPathLength = _countof(path);

    if (!WinHttpCrackUrl(url.c_str(), 0, 0, &uc)) return false;

    // Use default port if not specified
    INTERNET_PORT port = uc.nPort;
    if (port == 0) {
        port = (uc.nScheme == INTERNET_SCHEME_HTTPS) ? INTERNET_DEFAULT_HTTPS_PORT : INTERNET_DEFAULT_HTTP_PORT;
    }

    HINTERNET hSession = WinHttpOpen(
        L"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
        WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS, 0);

    if (!hSession) return false;

    DWORD redirectPolicy = WINHTTP_OPTION_REDIRECT_POLICY_ALWAYS;
    WinHttpSetOption(hSession, WINHTTP_OPTION_REDIRECT_POLICY, &redirectPolicy, sizeof(redirectPolicy));

    HINTERNET hConnect = WinHttpConnect(hSession, uc.lpszHostName, port, 0);
    if (!hConnect) { WinHttpCloseHandle(hSession); return false; }

    DWORD flags = (uc.nScheme == INTERNET_SCHEME_HTTPS) ? WINHTTP_FLAG_SECURE : 0;
    HINTERNET hRequest = WinHttpOpenRequest(
        hConnect, L"GET", uc.lpszUrlPath,
        nullptr, WINHTTP_NO_REFERER,
        WINHTTP_DEFAULT_ACCEPT_TYPES, flags);

    if (!hRequest) {
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return false;
    }

    // Add headers for image request
    std::wstring headers = L"Accept: image/*,*/*;q=0.8\r\nAccept-Encoding: identity\r\n";

    BOOL ok = WinHttpSendRequest(hRequest, headers.c_str(), (DWORD)headers.size(),
        WINHTTP_NO_REQUEST_DATA, 0, 0, 0);

    if (ok) ok = WinHttpReceiveResponse(hRequest, nullptr);

    if (!ok) {
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return false;
    }

    DWORD status = 0;
    DWORD statusSize = sizeof(status);
    WinHttpQueryHeaders(hRequest,
        WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
        WINHTTP_HEADER_NAME_BY_INDEX,
        &status, &statusSize, WINHTTP_NO_HEADER_INDEX);

    for (;;) {
        DWORD avail = 0;
        if (!WinHttpQueryDataAvailable(hRequest, &avail) || avail == 0) break;
        size_t old = outData.size();
        outData.resize(old + avail);
        DWORD read = 0;
        if (!WinHttpReadData(hRequest, outData.data() + old, avail, &read)) break;
        outData.resize(old + read);
    }

    WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);

    return (status >= 200 && status < 300);
}

// Cached version of HttpDownloadBinary - checks cache first, saves to cache after download
// If gameId is provided, registers the cached file for that game
static bool HttpDownloadBinaryCached(const std::wstring& url, std::vector<unsigned char>& outData, const std::wstring& gameId = L"") {
    // Check cache first
    if (LoadImageFromCache(url, outData)) {
        // Still register for the game if we have a gameId (for cache tracking)
        if (!gameId.empty()) {
            RegisterGameCacheFile(gameId, HashUrlToFilename(url));
        }
        return true;
    }
    
    // Not in cache, download
    if (!HttpDownloadBinary(url, outData)) {
        return false;
    }
    
    // Save to cache for next time
    if (!outData.empty()) {
        SaveImageToCache(url, outData);
        // Register for the game if we have a gameId
        if (!gameId.empty()) {
            RegisterGameCacheFile(gameId, HashUrlToFilename(url));
        }
    }
    
    return true;
}

static std::wstring SteamGridDbSearchGame(const std::wstring& gameName, const std::wstring& apiKey) {
    // Search for a game by name and return the game ID
    std::wstring searchTerm = gameName;
    // URL encode the search term
    std::wstring encodedSearch;
    for (wchar_t c : searchTerm) {
        if ((c >= L'A' && c <= L'Z') || (c >= L'a' && c <= L'z') ||
            (c >= L'0' && c <= L'9') || c == L'-' || c == L'_' || c == L'.') {
            encodedSearch += c;
        } else if (c == L' ') {
            encodedSearch += L"%20";
        } else {
            // Simple percent encoding for basic characters
            char buf[16];
            sprintf_s(buf, "%%%02X", (unsigned char)c);
            encodedSearch += Utf8ToWide(buf);
        }
    }

    std::wstring url = L"https://www.steamgriddb.com/api/v2/search/autocomplete/" + encodedSearch;
    std::string body;
    if (!HttpGetUtf8(url, apiKey, body)) return L"";

    // Parse JSON response to get game ID
    // Look for "id": <number>
    std::regex idRe(R"("id"\s*:\s*(\d+))");
    std::smatch m;
    if (std::regex_search(body, m, idRe)) {
        return Utf8ToWide(m[1].str());
    }
    return L"";
}

static std::wstring SteamGridDbGetGameIdBySteamAppId(const std::wstring& steamAppId, const std::wstring& apiKey) {
    std::wstring url = L"https://www.steamgriddb.com/api/v2/games/steam/" + steamAppId;
    std::string body;
    if (!HttpGetUtf8(url, apiKey, body)) return L"";

    // Parse JSON response to get game ID
    std::regex idRe(R"("id"\s*:\s*(\d+))");
    std::smatch m;
    if (std::regex_search(body, m, idRe)) {
        return Utf8ToWide(m[1].str());
    }
    return L"";
}

static std::vector<SteamGridDbImage> SteamGridDbGetGrids(const std::wstring& gameId, const std::wstring& apiKey) {
    std::vector<SteamGridDbImage> results;

    // Get grids (box art) for the game - filter for 600x900 style (2:3 ratio)
    std::wstring url = L"https://www.steamgriddb.com/api/v2/grids/game/" + gameId + L"?dimensions=600x900";
    std::string body;
    if (!HttpGetUtf8(url, apiKey, body)) return results;

    // Unescape JSON escaped slashes
    std::string unescaped;
    unescaped.reserve(body.size());
    for (size_t i = 0; i < body.size(); ++i) {
        if (body[i] == '\\' && i + 1 < body.size() && body[i + 1] == '/') {
            unescaped += '/';
            ++i;
        } else {
            unescaped += body[i];
        }
    }
    body = unescaped;

    // Parse JSON - simpler approach: find each url/thumb pair
    std::regex urlRe("\"url\"\\s*:\\s*\"([^\"]+)\"");
    std::regex thumbRe("\"thumb\"\\s*:\\s*\"([^\"]+)\"");
    std::regex idRe("\"id\"\\s*:\\s*(\\d+)");
    std::regex widthRe("\"width\"\\s*:\\s*(\\d+)");
    std::regex heightRe("\"height\"\\s*:\\s*(\\d+)");

    // Split by objects (look for url fields as markers)
    std::string::const_iterator it(body.cbegin());
    std::smatch m;
    while (std::regex_search(it, body.cend(), m, urlRe)) {
        SteamGridDbImage img;
        img.url = Utf8ToWide(m[1].str());

        // Look backwards for id (within reasonable range)
        std::string prefix(body.cbegin(), m[0].first);
        std::smatch pm;
        if (std::regex_search(prefix, pm, idRe)) {
            // Find the last match
            std::string::const_iterator searchIt = prefix.cbegin();
            while (std::regex_search(searchIt, prefix.cend(), pm, idRe)) {
                img.id = Utf8ToWide(pm[1].str());
                searchIt = pm.suffix().first;
            }
        }

        // Look forward for thumb
        std::string remaining(m.suffix().first, body.cend());
        std::smatch tm;
        if (std::regex_search(remaining, tm, thumbRe)) {
            img.thumb = Utf8ToWide(tm[1].str());
        }

        // Look for width/height
        if (std::regex_search(remaining, tm, widthRe)) {
            img.width = std::stoi(tm[1].str());
        }
        if (std::regex_search(remaining, tm, heightRe)) {
            img.height = std::stoi(tm[1].str());
        }

        if (!img.url.empty()) {
            results.push_back(img);
        }
        it = m.suffix().first;
    }

    return results;
}

// Generic function to fetch images from SteamGridDB for different art types
static std::vector<SteamGridDbImage> SteamGridDbGetImages(const std::wstring& gameId, const std::wstring& apiKey, const std::wstring& artType, const std::wstring& dimensions = L"") {
    std::vector<SteamGridDbImage> results;

    // Build URL: /api/v2/{type}/game/{id}
    std::wstring url = L"https://www.steamgriddb.com/api/v2/" + artType + L"/game/" + gameId;
    if (!dimensions.empty()) {
        url += L"?dimensions=" + dimensions;
    }
    
    std::string body;
    if (!HttpGetUtf8(url, apiKey, body)) return results;

    // Unescape JSON escaped slashes
    std::string unescaped;
    unescaped.reserve(body.size());
    for (size_t i = 0; i < body.size(); ++i) {
        if (body[i] == '\\' && i + 1 < body.size() && body[i + 1] == '/') {
            unescaped += '/';
            ++i;
        } else {
            unescaped += body[i];
        }
    }
    body = unescaped;

    // Parse JSON
    std::regex urlRe("\"url\"\\s*:\\s*\"([^\"]+)\"");
    std::regex thumbRe("\"thumb\"\\s*:\\s*\"([^\"]+)\"");
    std::regex idRe("\"id\"\\s*:\\s*(\\d+)");
    std::regex widthRe("\"width\"\\s*:\\s*(\\d+)");
    std::regex heightRe("\"height\"\\s*:\\s*(\\d+)");

    std::string::const_iterator it(body.cbegin());
    std::smatch m;
    while (std::regex_search(it, body.cend(), m, urlRe)) {
        SteamGridDbImage img;
        img.url = Utf8ToWide(m[1].str());

        std::string prefix(body.cbegin(), m[0].first);
        std::smatch pm;
        if (std::regex_search(prefix, pm, idRe)) {
            std::string::const_iterator searchIt = prefix.cbegin();
            while (std::regex_search(searchIt, prefix.cend(), pm, idRe)) {
                img.id = Utf8ToWide(pm[1].str());
                searchIt = pm.suffix().first;
            }
        }

        std::string remaining(m.suffix().first, body.cend());
        std::smatch tm;
        if (std::regex_search(remaining, tm, thumbRe)) {
            img.thumb = Utf8ToWide(tm[1].str());
        }
        if (std::regex_search(remaining, tm, widthRe)) {
            img.width = std::stoi(tm[1].str());
        }
        if (std::regex_search(remaining, tm, heightRe)) {
            img.height = std::stoi(tm[1].str());
        }

        if (!img.url.empty()) {
            results.push_back(img);
        }
        it = m.suffix().first;
    }

    return results;
}

static std::vector<SteamGridDbImage> SteamGridDbGetHeroes(const std::wstring& gameId, const std::wstring& apiKey) {
    return SteamGridDbGetImages(gameId, apiKey, L"heroes");
}

static std::vector<SteamGridDbImage> SteamGridDbGetLogos(const std::wstring& gameId, const std::wstring& apiKey) {
    return SteamGridDbGetImages(gameId, apiKey, L"logos");
}

static std::vector<SteamGridDbImage> SteamGridDbGetIcons(const std::wstring& gameId, const std::wstring& apiKey) {
    return SteamGridDbGetImages(gameId, apiKey, L"icons");
}

// Forward declaration for debug logging
static void DebugLog(const std::wstring& msg);

// Web image search using Google Images
// Searches for game cover art with specified resolution
static std::vector<SteamGridDbImage> WebSearchGameImages(const std::wstring& gameTitle, const std::wstring& resolution, int pageIndex = 0) {
    std::vector<SteamGridDbImage> results;
    
    // Build search query: "GAME TITLE game cover imagesize:WIDTHxHEIGHT"
    std::wstring searchQuery = gameTitle + L" game cover imagesize:" + resolution;
    
    DebugLog(L"[Web] Search query: " + searchQuery);
    DebugLog(L"[Web] Page: " + std::to_wstring(pageIndex));
    
    // URL encode the search query
    std::wstring encodedQuery;
    for (wchar_t c : searchQuery) {
        if ((c >= L'A' && c <= L'Z') || (c >= L'a' && c <= L'z') ||
            (c >= L'0' && c <= L'9') || c == L'-' || c == L'_' || c == L'.' || c == L':') {
            encodedQuery += c;
        } else if (c == L' ') {
            encodedQuery += L"+";
        } else {
            // Percent encode
            char narrowChar = (char)c;
            wchar_t buf[8];
            swprintf(buf, 8, L"%%%02X", (unsigned char)narrowChar);
            encodedQuery += buf;
        }
    }
    
    // Use Google Images search
    // tbm=isch for image search
    // start= for pagination (Google uses ~20 results per page)
    int startIndex = pageIndex * 20;
    std::wstring url = L"https://www.google.com/search?q=" + encodedQuery + 
                       L"&tbm=isch&start=" + std::to_wstring(startIndex);
    
    DebugLog(L"[Web] URL: " + url);
    
    // Use WinHTTP to get the page
    HINTERNET hSession = WinHttpOpen(L"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
        WINHTTP_NO_PROXY_NAME,
        WINHTTP_NO_PROXY_BYPASS, 0);
    if (!hSession) {
        DebugLog(L"[Web] ERROR: Failed to open WinHTTP session");
        return results;
    }
    
    URL_COMPONENTSW urlComp{};
    urlComp.dwStructSize = sizeof(urlComp);
    wchar_t hostName[256]{};
    wchar_t urlPath[2048]{};
    urlComp.lpszHostName = hostName;
    urlComp.dwHostNameLength = _countof(hostName);
    urlComp.lpszUrlPath = urlPath;
    urlComp.dwUrlPathLength = _countof(urlPath);
    
    if (!WinHttpCrackUrl(url.c_str(), 0, 0, &urlComp)) {
        DebugLog(L"[Web] ERROR: Failed to parse URL");
        WinHttpCloseHandle(hSession);
        return results;
    }
    
    HINTERNET hConnect = WinHttpConnect(hSession, hostName, urlComp.nPort, 0);
    if (!hConnect) {
        DebugLog(L"[Web] ERROR: Failed to connect to server");
        WinHttpCloseHandle(hSession);
        return results;
    }
    
    DWORD flags = WINHTTP_FLAG_SECURE;
    HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"GET", urlPath,
        nullptr, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
    if (!hRequest) {
        DebugLog(L"[Web] ERROR: Failed to open request");
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return results;
    }
    
    // Set headers to look like a real browser
    WinHttpAddRequestHeaders(hRequest, 
        L"Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8", 
        -1, WINHTTP_ADDREQ_FLAG_ADD);
    WinHttpAddRequestHeaders(hRequest, 
        L"Accept-Language: en-US,en;q=0.5", 
        -1, WINHTTP_ADDREQ_FLAG_ADD);
    
    if (!WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0, nullptr, 0, 0, 0)) {
        DebugLog(L"[Web] ERROR: Failed to send request");
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return results;
    }
    
    if (!WinHttpReceiveResponse(hRequest, nullptr)) {
        DebugLog(L"[Web] ERROR: Failed to receive response");
        WinHttpCloseHandle(hRequest);
        WinHttpCloseHandle(hConnect);
        WinHttpCloseHandle(hSession);
        return results;
    }
    
    // Read response
    std::string htmlContent;
    DWORD bytesRead = 0;
    char buffer[8192];
    while (WinHttpReadData(hRequest, buffer, sizeof(buffer), &bytesRead) && bytesRead > 0) {
        htmlContent.append(buffer, bytesRead);
    }
    
    WinHttpCloseHandle(hRequest);
    WinHttpCloseHandle(hConnect);
    WinHttpCloseHandle(hSession);
    
    DebugLog(L"[Web] Response size: " + std::to_wstring(htmlContent.size()) + L" bytes");
    
    // Google embeds image data in the page in various formats
    // Look for image URLs in data attributes or JSON-like structures
    std::vector<std::string> imageUrls;
    
    // Pattern 1: Look for ["URL",width,height] patterns (Google's image data format)
    {
        std::regex imgRegex("\\[\"(https?://[^\"]+\\.(?:jpg|jpeg|png|webp)[^\"]*)\",[0-9]+,[0-9]+\\]");
        auto begin = std::sregex_iterator(htmlContent.begin(), htmlContent.end(), imgRegex);
        auto end = std::sregex_iterator();
        for (auto it = begin; it != end; ++it) {
            std::string url = (*it)[1].str();
            // Skip Google's own thumbnails
            if (url.find("gstatic.com") == std::string::npos && 
                url.find("google.com") == std::string::npos) {
                imageUrls.push_back(url);
            }
        }
        DebugLog(L"[Web] Pattern 1 found: " + std::to_wstring(imageUrls.size()));
    }
    
    // Pattern 2: Look for ou":"URL" (original URL in JSON)
    {
        std::regex imgRegex("\"ou\":\"(https?://[^\"]+)\"");
        auto begin = std::sregex_iterator(htmlContent.begin(), htmlContent.end(), imgRegex);
        auto end = std::sregex_iterator();
        for (auto it = begin; it != end; ++it) {
            std::string url = (*it)[1].str();
            if (url.find("gstatic.com") == std::string::npos && 
                url.find("google.com") == std::string::npos) {
                // Check if not already added
                bool found = false;
                for (const auto& existing : imageUrls) {
                    if (existing == url) { found = true; break; }
                }
                if (!found) imageUrls.push_back(url);
            }
        }
        DebugLog(L"[Web] After pattern 2: " + std::to_wstring(imageUrls.size()));
    }
    
    // Pattern 3: Look for data-src with image URLs
    {
        std::regex imgRegex("data-src=\"(https?://[^\"]+\\.(?:jpg|jpeg|png|webp)[^\"]*)\"");
        auto begin = std::sregex_iterator(htmlContent.begin(), htmlContent.end(), imgRegex);
        auto end = std::sregex_iterator();
        for (auto it = begin; it != end; ++it) {
            std::string url = (*it)[1].str();
            if (url.find("gstatic.com") == std::string::npos && 
                url.find("google.com") == std::string::npos) {
                bool found = false;
                for (const auto& existing : imageUrls) {
                    if (existing == url) { found = true; break; }
                }
                if (!found) imageUrls.push_back(url);
            }
        }
        DebugLog(L"[Web] After pattern 3: " + std::to_wstring(imageUrls.size()));
    }
    
    int count = 0;
    for (const auto& imgUrl : imageUrls) {
        
        // Decode any escaped characters
        std::string decodedUrl;
        for (size_t i = 0; i < imgUrl.size(); ++i) {
            if (imgUrl[i] == '\\' && i + 1 < imgUrl.size() && imgUrl[i+1] == '/') {
                decodedUrl += '/';
                ++i;
            } else if (imgUrl[i] == '\\' && i + 1 < imgUrl.size() && imgUrl[i+1] == 'u') {
                // Unicode escape \uXXXX - try to decode
                if (i + 5 < imgUrl.size()) {
                    try {
                        std::string hex = imgUrl.substr(i + 2, 4);
                        int codepoint = std::stoi(hex, nullptr, 16);
                        if (codepoint < 128) {
                            decodedUrl += (char)codepoint;
                        }
                        i += 5;
                    } catch (...) {
                        decodedUrl += imgUrl[i];
                    }
                } else {
                    decodedUrl += imgUrl[i];
                }
            } else {
                decodedUrl += imgUrl[i];
            }
        }
        
        // Skip if URL looks invalid
        if (decodedUrl.find("http") != 0) continue;
        
        SteamGridDbImage img;
        img.id = std::to_wstring(startIndex + count);
        img.url = Utf8ToWide(decodedUrl);
        img.thumb = img.url;  // Use same URL for thumbnail
        img.author = L"Google Images";
        
        // Parse resolution from the search query
        size_t xPos = resolution.find(L'x');
        if (xPos != std::wstring::npos) {
            img.width = _wtoi(resolution.substr(0, xPos).c_str());
            img.height = _wtoi(resolution.substr(xPos + 1).c_str());
        }
        
        results.push_back(img);
        ++count;
    }
    
    DebugLog(L"[Web] Found " + std::to_wstring(results.size()) + L" images");
    
    return results;
}

// ------------------------------------------------------------
// Art Update window
// ------------------------------------------------------------

static void DestroyArtImageList() {
    if (gArtImgList) {
        ImageList_Destroy(gArtImgList);
        gArtImgList = nullptr;
    }
}

static void CloseArtWindow(bool applyChanges);
static void OpenConfigWindow(HWND parent);

static HBITMAP CreateThumbnailFromUrl(const std::wstring& url, int size, const std::wstring& gameId = L"") {
    std::vector<unsigned char> data;
    if (!HttpDownloadBinaryCached(url, data, gameId) || data.empty()) {
        return nullptr;
    }

    // Create GDI+ image from memory
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, data.size());
    if (!hMem) return nullptr;

    void* pMem = GlobalLock(hMem);
    if (pMem) {
        memcpy(pMem, data.data(), data.size());
        GlobalUnlock(hMem);

        IStream* pStream = nullptr;
        if (SUCCEEDED(CreateStreamOnHGlobal(hMem, FALSE, &pStream))) {
            Gdiplus::Image* srcImg = Gdiplus::Image::FromStream(pStream);
            pStream->Release();

            if (srcImg && srcImg->GetLastStatus() == Gdiplus::Ok) {
                // Create thumbnail bitmap
                Gdiplus::Bitmap* bmp = new Gdiplus::Bitmap(size, size, PixelFormat32bppARGB);
                Gdiplus::Graphics g(bmp);
                g.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
                g.Clear(gDarkMode ? Gdiplus::Color(255, 45, 45, 45) : Gdiplus::Color(255, 240, 240, 240));

                // Scale to fit
                int srcW = srcImg->GetWidth();
                int srcH = srcImg->GetHeight();
                float scale = (float)size / std::max(srcW, srcH);
                int dstW = (int)(srcW * scale);
                int dstH = (int)(srcH * scale);
                int x = (size - dstW) / 2;
                int y = (size - dstH) / 2;

                g.DrawImage(srcImg, x, y, dstW, dstH);
                delete srcImg;

                HBITMAP hBmp = nullptr;
                bmp->GetHBITMAP(Gdiplus::Color(255, 255, 255, 255), &hBmp);
                delete bmp;

                GlobalFree(hMem);
                return hBmp;
            }
            if (srcImg) delete srcImg;
        }
    }
    GlobalFree(hMem);
    return nullptr;
}

// Helper to append debug message to the debug text box
static void DebugLog(const std::wstring& msg) {
    if (!gArtDebug) return;
    int len = GetWindowTextLengthW(gArtDebug);
    SendMessageW(gArtDebug, EM_SETSEL, len, len);
    std::wstring line = msg + L"\r\n";
    SendMessageW(gArtDebug, EM_REPLACESEL, FALSE, (LPARAM)line.c_str());
    // Scroll to bottom
    SendMessageW(gArtDebug, EM_SCROLLCARET, 0, 0);
}

static void PopulateArtList() {
    if (!gArtList) return;

    // Cancel any previous loading
    gArtCancelLoading = true;
    
    // Increment generation to invalidate any in-progress loads
    ++gArtLoadGeneration;
    int myGeneration = gArtLoadGeneration;
    
    // Wait a moment for previous loading to notice cancellation
    MSG msg;
    for (int i = 0; i < 10 && gArtIsLoading; ++i) {
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
        Sleep(10);
    }

    gArtCancelLoading = false;
    gArtIsLoading = true;

    // Disable redraw during bulk operations
    SendMessageW(gArtList, WM_SETREDRAW, FALSE, 0);

    ListView_DeleteAllItems(gArtList);
    DestroyArtImageList();

    const int thumbSize = 128;
    gArtImgList = ImageList_Create(thumbSize, thumbSize, ILC_COLOR32, (int)gArtImages.size() + 1, 4);
    ListView_SetImageList(gArtList, gArtImgList, LVSIL_NORMAL);

    // Add placeholder for loading
    Gdiplus::Bitmap* placeholder = new Gdiplus::Bitmap(thumbSize, thumbSize, PixelFormat32bppARGB);
    {
        Gdiplus::Graphics g(placeholder);
        g.Clear(gDarkMode ? Gdiplus::Color(255, 50, 50, 50) : Gdiplus::Color(255, 200, 200, 200));
        Gdiplus::Font font(L"Segoe UI", 10);
        Gdiplus::SolidBrush brush(gDarkMode ? Gdiplus::Color(255, 150, 150, 150) : Gdiplus::Color(255, 100, 100, 100));
        Gdiplus::StringFormat sf;
        sf.SetAlignment(Gdiplus::StringAlignmentCenter);
        sf.SetLineAlignment(Gdiplus::StringAlignmentCenter);
        Gdiplus::RectF rect(0, 0, (float)thumbSize, (float)thumbSize);
        g.DrawString(L"Loading...", -1, &font, rect, &sf, &brush);
    }
    HBITMAP hPlaceholder = nullptr;
    placeholder->GetHBITMAP(Gdiplus::Color(255, 255, 255, 255), &hPlaceholder);
    delete placeholder;
    int placeholderIdx = ImageList_Add(gArtImgList, hPlaceholder, nullptr);
    DeleteObject(hPlaceholder);

    // Add items with placeholder initially
    for (size_t i = 0; i < gArtImages.size(); ++i) {
        LVITEM it{};
        it.mask = LVIF_TEXT | LVIF_IMAGE | LVIF_PARAM;
        it.iItem = (int)i;
        it.iImage = placeholderIdx;
        it.lParam = (LPARAM)i;

        std::wstring label = L"Image " + std::to_wstring(i + 1);
        if (gArtImages[i].width > 0) {
            label += L" (" + std::to_wstring(gArtImages[i].width) + L"x" + std::to_wstring(gArtImages[i].height) + L")";
        }
        it.pszText = (LPWSTR)label.c_str();

        ListView_InsertItem(gArtList, &it);
    }

    ListView_SetView(gArtList, LV_VIEW_ICON);
    SendMessageW(gArtList, LVM_SETICONSPACING, 0, MAKELPARAM(thumbSize + 40, thumbSize + 40));

    // Re-enable redraw
    SendMessageW(gArtList, WM_SETREDRAW, TRUE, 0);
    InvalidateRect(gArtList, nullptr, TRUE);

    // Update UI before loading
    UpdateWindow(gArtWnd);

    DebugLog(L"Starting to load " + std::to_wstring(gArtImages.size()) + L" images with parallel downloads...");

    // Copy images to download - gArtImages may change if user switches tabs
    struct DownloadTask {
        size_t index;
        std::wstring url;
    };
    std::vector<DownloadTask> tasks;
    for (size_t i = 0; i < gArtImages.size(); ++i) {
        std::wstring thumbUrl = gArtImages[i].thumb.empty() ? gArtImages[i].url : gArtImages[i].thumb;
        tasks.push_back({ i, thumbUrl });
    }
    
    size_t imageCount = tasks.size();
    
    // Determine parallel download count from config (0 = auto = CPU threads)
    int parallelCount = gConfig.parallelDownloads;
    if (parallelCount <= 0) {
        parallelCount = (int)std::thread::hardware_concurrency();
        if (parallelCount < 2) parallelCount = 8; // Fallback
    }
    
    // Results structure
    struct DownloadResult {
        size_t index;
        HBITMAP hBitmap;
    };
    std::mutex resultsMutex;
    std::vector<DownloadResult> completedResults;
    std::atomic<size_t> nextTask(0);
    std::atomic<size_t> completedCount(0);
    std::atomic<bool> cancelFlag(false);
    
    // Worker function - capture gameId for cache tracking
    std::wstring gameIdForCache = gArtGameId;
    auto workerFunc = [&, gameIdForCache]() {
        while (!cancelFlag.load()) {
            size_t taskIdx = nextTask.fetch_add(1);
            if (taskIdx >= tasks.size()) break;
            
            const auto& task = tasks[taskIdx];
            HBITMAP hThumb = CreateThumbnailFromUrl(task.url, thumbSize, gameIdForCache);
            
            if (cancelFlag.load()) {
                if (hThumb) DeleteObject(hThumb);
                break;
            }
            
            {
                std::lock_guard<std::mutex> lock(resultsMutex);
                completedResults.push_back({ task.index, hThumb });
            }
            completedCount.fetch_add(1);
        }
    };
    
    // Start worker threads
    std::vector<std::thread> workers;
    int numThreads = (imageCount < (size_t)parallelCount) ? (int)imageCount : parallelCount;
    for (int t = 0; t < numThreads; ++t) {
        workers.emplace_back(workerFunc);
    }
    
    // Process results on main thread while workers are downloading
    size_t processedCount = 0;
    while (processedCount < imageCount) {
        // Check for cancellation
        if (gArtCancelLoading || !gArtWnd || myGeneration != gArtLoadGeneration) {
            DebugLog(L"Loading cancelled.");
            cancelFlag.store(true);
            break;
        }
        
        // Process pending Windows messages
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
            if (gArtCancelLoading || !gArtWnd || myGeneration != gArtLoadGeneration) {
                cancelFlag.store(true);
                break;
            }
        }
        
        if (cancelFlag.load()) break;
        
        // Process any completed results
        std::vector<DownloadResult> toProcess;
        {
            std::lock_guard<std::mutex> lock(resultsMutex);
            toProcess.swap(completedResults);
        }
        
        for (const auto& result : toProcess) {
            if (!gArtWnd || !gArtList || !gArtImgList || myGeneration != gArtLoadGeneration) {
                if (result.hBitmap) DeleteObject(result.hBitmap);
                continue;
            }
            
            if (result.hBitmap) {
                int imgIdx = ImageList_Add(gArtImgList, result.hBitmap, nullptr);
                DeleteObject(result.hBitmap);
                
                LVITEM it{};
                it.mask = LVIF_IMAGE;
                it.iItem = (int)result.index;
                it.iImage = imgIdx;
                ListView_SetItem(gArtList, &it);
                
                DebugLog(L"[" + std::to_wstring(result.index + 1) + L"] OK - loaded successfully");
            } else {
                DebugLog(L"[" + std::to_wstring(result.index + 1) + L"] FAILED - could not load thumbnail");
            }
            processedCount++;
        }
        
        // Update status
        size_t currentCompleted = completedCount.load();
        if (currentCompleted > 0) {
            std::wstring statusMsg = L"Loading thumbnails: " + std::to_wstring(currentCompleted) + L" of " + std::to_wstring(imageCount) + L"...";
            SetWindowTextW(gArtStatus, statusMsg.c_str());
        }
        
        // Redraw periodically
        if (processedCount % 4 == 0 || processedCount == imageCount) {
            ListView_RedrawItems(gArtList, 0, (int)imageCount - 1);
            UpdateWindow(gArtList);
        }
        
        // Small sleep to avoid busy-waiting
        Sleep(10);
    }
    
    // Wait for all worker threads to finish
    for (auto& worker : workers) {
        if (worker.joinable()) {
            worker.join();
        }
    }
    
    // Clean up any remaining results if cancelled
    {
        std::lock_guard<std::mutex> lock(resultsMutex);
        for (const auto& result : completedResults) {
            if (result.hBitmap) DeleteObject(result.hBitmap);
        }
        completedResults.clear();
    }

    gArtIsLoading = false;

    // Only update final status if this is still the current generation
    if (myGeneration == gArtLoadGeneration && gArtStatus && !cancelFlag.load()) {
        std::wstring finalStatus = L"Found " + std::to_wstring(imageCount) + L" images. Select one and click Apply.";
        SetWindowTextW(gArtStatus, finalStatus.c_str());
        DebugLog(L"Thumbnail loading complete.");
        
        // Final redraw
        if (gArtList) {
            ListView_RedrawItems(gArtList, 0, (int)imageCount - 1);
            UpdateWindow(gArtList);
        }
    }
}

static void ApplySelectedArt() {
    if (!gArtList) return;

    int sel = ListView_GetNextItem(gArtList, -1, LVNI_SELECTED);
    if (sel < 0 || sel >= (int)gArtImages.size()) {
        DarkMessageBox(gArtWnd, L"Please select an image first.", L"No Selection", MB_OK | MB_ICONINFORMATION);
        return;
    }

    const SteamGridDbImage& img = gArtImages[sel];

    // Download full image
    // Pass gameId for cache tracking - will re-download if cache was cleared
    SetWindowTextW(gArtStatus, L"Downloading image...");

    std::vector<unsigned char> data;
    if (!HttpDownloadBinaryCached(img.url, data, gArtGameId) || data.empty()) {
        SetWindowTextW(gArtStatus, L"Failed to download image.");
        DarkMessageBox(gArtWnd, L"Failed to download the selected image.", L"Download Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Determine save path
    std::wstring basePath = GetThirdPartyLibrariesPath();
    fs::path storePath = fs::path(basePath) / gArtCurrentEntry.store;
    std::wstring fileName = gArtCurrentEntry.expectedFileName;
    if (fileName.empty()) {
        // Generate filename from ID
        fileName = ExpectedPngFromManifestId(gArtCurrentEntry.store, gArtCurrentEntry.idStr);
    }

    fs::path savePath = storePath / fileName;

    // Create directory if needed
    if (!fs::exists(storePath)) {
        fs::create_directories(storePath);
    }

    // Backup original art if it exists and we don't have a backup yet
    if (gArtCurrentEntry.hasArt && !gArtCurrentEntry.filePath.empty()) {
        BackupOriginalArt(gArtCurrentEntry.store, gArtCurrentEntry.filePath, fileName);
    }

    // Remove read-only attribute before writing (if file exists)
    RemoveFileReadOnly(savePath.wstring());

    // Save the image
    HANDLE hFile = CreateFileW(savePath.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) {
        SetWindowTextW(gArtStatus, L"Failed to save image.");
        DarkMessageBox(gArtWnd, L"Failed to save the image file.", L"Save Error", MB_OK | MB_ICONERROR);
        return;
    }

    DWORD written = 0;
    WriteFile(hFile, data.data(), (DWORD)data.size(), &written, nullptr);
    CloseHandle(hFile);

    // Set file as read-only to prevent Xbox app from modifying/degrading the image
    SetFileReadOnly(savePath.wstring());

    // Update the entry in gLastItems
    for (auto& e : gLastItems) {
        if (e.store == gArtCurrentEntry.store && e.idStr == gArtCurrentEntry.idStr) {
            e.hasArt = true;
            e.filePath = savePath.wstring();
            e.fileName = fileName;
            break;
        }
    }

    // Always update customTitle from the field before saving
    HWND hEdit = GetDlgItem(gArtWnd, 4100); // IDC_ART_TITLE_EDIT
    if (hEdit) {
        wchar_t titleBuf[256] = {};
        GetWindowTextW(hEdit, titleBuf, _countof(titleBuf));
        std::wstring newTitle = titleBuf;
        for (auto& e : gLastItems) {
            if (e.store == gArtCurrentEntry.store && e.idStr == gArtCurrentEntry.idStr) {
                e.customTitle = newTitle;
                break;
            }
        }
    }
    // Save cache and refresh
    SaveUiStateToCacheFile();

    SetWindowTextW(gArtStatus, L"Art saved successfully!");
    DarkMessageBox(gArtWnd, L"Art has been saved successfully!\n\nThe file has been set to read-only to prevent the Xbox app from modifying it.", L"Success", MB_OK | MB_ICONINFORMATION);

    CloseArtWindow(false);

    // Refresh main list
    DestroyImageList();
    EnsureImageList();
    PopulateListFromItems(gLastItems);
}

// Restore original art
static void RestoreOriginalArt() {
    std::wstring fileName = gArtCurrentEntry.expectedFileName;
    if (fileName.empty()) {
        fileName = ExpectedPngFromManifestId(gArtCurrentEntry.store, gArtCurrentEntry.idStr);
    }

    if (!HasBackup(gArtCurrentEntry.store, fileName)) {
        DarkMessageBox(gArtWnd, L"No backup found for this game's artwork.", L"No Backup", MB_OK | MB_ICONINFORMATION);
        return;
    }

    if (RestoreArtFromBackup(gArtCurrentEntry.store, fileName)) {
        // Update the entry in gLastItems
        std::wstring basePath = GetThirdPartyLibrariesPath();
        fs::path restoredPath = fs::path(basePath) / gArtCurrentEntry.store / fileName;

        for (auto& e : gLastItems) {
            if (e.store == gArtCurrentEntry.store && e.idStr == gArtCurrentEntry.idStr) {
                e.hasArt = true;
                e.filePath = restoredPath.wstring();
                e.fileName = fileName;
                break;
            }
        }

        SaveUiStateToCacheFile();
        DarkMessageBox(gArtWnd, L"Original artwork has been restored!", L"Success", MB_OK | MB_ICONINFORMATION);

        CloseArtWindow(false);

        // Refresh main list
        DestroyImageList();
        EnsureImageList();
        PopulateListFromItems(gLastItems);
    } else {
        DarkMessageBox(gArtWnd, L"Failed to restore the original artwork.", L"Error", MB_OK | MB_ICONERROR);
    }
}

// Custom draw for art preview (current or selected)
static void DrawArtPreview(HDC hdc, RECT& rect, HBITMAP bitmap, const wchar_t* placeholder) {
    // Fill background first
    HBRUSH bgBrush = CreateSolidBrush(gDarkMode ? RGB(32, 32, 32) : RGB(220, 220, 220));
    FillRect(hdc, &rect, bgBrush);
    DeleteObject(bgBrush);
    
    if (bitmap) {
        HDC memDC = CreateCompatibleDC(hdc);
        HGDIOBJ oldBmp = SelectObject(memDC, bitmap);

        BITMAP bm;
        GetObject(bitmap, sizeof(bm), &bm);

        // Calculate aspect-fit dimensions
        int rectW = rect.right - rect.left;
        int rectH = rect.bottom - rect.top;
        float scale = (std::min)((float)rectW / bm.bmWidth, (float)rectH / bm.bmHeight);
        int dstW = (int)(bm.bmWidth * scale);
        int dstH = (int)(bm.bmHeight * scale);
        int x = rect.left + (rectW - dstW) / 2;
        int y = rect.top + (rectH - dstH) / 2;

        SetStretchBltMode(hdc, HALFTONE);
        StretchBlt(hdc, x, y, dstW, dstH, memDC, 0, 0, bm.bmWidth, bm.bmHeight, SRCCOPY);

        SelectObject(memDC, oldBmp);
        DeleteDC(memDC);
    } else {
        // Draw placeholder text
        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, gDarkMode ? RGB(160, 160, 160) : RGB(128, 128, 128));
        DrawTextW(hdc, placeholder, -1, &rect, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }
}

// Load preview for selected image from SteamGridDB
static void LoadSelectedPreview(int index) {
    // Clear existing preview first
    if (gArtPreviewBitmap) {
        DeleteObject(gArtPreviewBitmap);
        gArtPreviewBitmap = nullptr;
    }
    gArtSelectedIndex = index;
    
    // Force immediate repaint to clear old image
    if (gArtPreviewImg) {
        InvalidateRect(gArtPreviewImg, nullptr, TRUE);
        UpdateWindow(gArtPreviewImg);
    }

    if (index < 0 || index >= (int)gArtImages.size()) {
        return;
    }

    // Use thumb URL for preview (faster) or full URL
    std::wstring previewUrl = gArtImages[index].thumb.empty() ? gArtImages[index].url : gArtImages[index].thumb;
    
    DebugLog(L"Loading preview for selection " + std::to_wstring(index + 1) + L": " + previewUrl);

    // Download and create bitmap (larger size for preview)
    // Pass gameId for cache tracking - will re-download if cache was cleared
    std::vector<unsigned char> data;
    if (HttpDownloadBinaryCached(previewUrl, data, gArtGameId) && !data.empty()) {
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, data.size());
        if (hMem) {
            void* pMem = GlobalLock(hMem);
            if (pMem) {
                memcpy(pMem, data.data(), data.size());
                GlobalUnlock(hMem);

                IStream* pStream = nullptr;
                if (SUCCEEDED(CreateStreamOnHGlobal(hMem, FALSE, &pStream))) {
                    Gdiplus::Image* srcImg = Gdiplus::Image::FromStream(pStream);
                    pStream->Release();

                    if (srcImg && srcImg->GetLastStatus() == Gdiplus::Ok) {
                        // Create a larger bitmap for the preview (keep original aspect ratio)
                        int srcW = srcImg->GetWidth();
                        int srcH = srcImg->GetHeight();
                        int maxSize = 300;
                        float scale = (float)maxSize / std::max(srcW, srcH);
                        int dstW = (int)(srcW * scale);
                        int dstH = (int)(srcH * scale);

                        Gdiplus::Bitmap* bmp = new Gdiplus::Bitmap(dstW, dstH, PixelFormat32bppARGB);
                        Gdiplus::Graphics g(bmp);
                        g.SetInterpolationMode(Gdiplus::InterpolationModeHighQualityBicubic);
                        g.Clear(gDarkMode ? Gdiplus::Color(255, 45, 45, 45) : Gdiplus::Color(255, 240, 240, 240));
                        g.DrawImage(srcImg, 0, 0, dstW, dstH);

                        bmp->GetHBITMAP(Gdiplus::Color(255, 255, 255, 255), &gArtPreviewBitmap);
                        delete bmp;
                    }
                    if (srcImg) delete srcImg;
                }
            }
            GlobalFree(hMem);
        }
    }

    // Redraw preview
    if (gArtPreviewImg) InvalidateRect(gArtPreviewImg, nullptr, TRUE);
}

static LRESULT CALLBACK ArtWndProc(HWND hWnd, UINT msg, WPARAM w, LPARAM l) {

        // Handle Save and Search button for title
        static const int IDC_ART_TITLE_EDIT = 4100;
        static const int IDC_ART_TITLE_SAVE_BTN = 4101;

        if (msg == WM_COMMAND) {
            int wmId = LOWORD(w);
            if (wmId == IDC_ART_TITLE_SAVE_BTN) {
                // Get text from the edit control
                wchar_t titleBuf[256] = {};
                HWND hEdit = GetDlgItem(hWnd, IDC_ART_TITLE_EDIT);
                if (hEdit) {
                    GetWindowTextW(hEdit, titleBuf, _countof(titleBuf));
                    std::wstring newTitle = titleBuf;
                    if (!newTitle.empty() && newTitle != gArtCurrentEntry.customTitle) {
                        sCustomTitleChanged = true;
                        // Save to config file in %TEMP%\XboxAppUpdaterArtCache\config.json
                        wchar_t tempPath[MAX_PATH];
                        GetTempPathW(MAX_PATH, tempPath);
                        std::wstring configPath = tempPath;
                        configPath += L"XboxAppUpdaterArtCache\\config.json";

                        // Read the file
                        FILE* f = _wfopen(configPath.c_str(), L"r, ccs=UTF-8");
                        std::wstring fileContent;
                        if (f) {
                            wchar_t buf[4096];
                            while (fgetws(buf, 4096, f)) fileContent += buf;
                            fclose(f);
                        }
                        // Find the entry for the current game id and add/update customTitle
                        std::wstring idKey = L"id:\"" + gArtCurrentEntry.idStr + L"\"";
                        size_t pos = fileContent.find(idKey);
                        if (pos != std::wstring::npos) {
                            // Look for customTitle in this entry
                            size_t entryEnd = fileContent.find(L"},", pos);
                            if (entryEnd == std::wstring::npos) entryEnd = fileContent.find(L"}\n", pos);
                            if (entryEnd == std::wstring::npos) entryEnd = fileContent.find(L"}\r\n", pos);
                            if (entryEnd == std::wstring::npos) entryEnd = fileContent.find(L"}\r", pos);
                            if (entryEnd == std::wstring::npos) entryEnd = fileContent.find(L"}", pos);
                            if (entryEnd != std::wstring::npos) {
                                size_t customTitlePos = fileContent.find(L"customTitle:", pos);
                                if (customTitlePos != std::wstring::npos && customTitlePos < entryEnd) {
                                    // Update existing customTitle
                                    size_t valueStart = fileContent.find(L'"', customTitlePos);
                                    if (valueStart != std::wstring::npos && valueStart < entryEnd) {
                                        valueStart++;
                                        size_t valueEnd = fileContent.find(L'"', valueStart);
                                        if (valueEnd != std::wstring::npos && valueEnd < entryEnd) {
                                            fileContent.replace(valueStart, valueEnd - valueStart, newTitle);
                                        }
                                    }
                                } else {
                                    // Insert new customTitle before the end of the entry
                                    std::wstring insertStr = L",customTitle:\"" + newTitle + L"\"";
                                    fileContent.insert(entryEnd, insertStr);
                                }
                            }
                        }
                        // Write back the file
                        f = _wfopen(configPath.c_str(), L"w, ccs=UTF-8");
                        if (f) {
                            fputws(fileContent.c_str(), f);
                            fclose(f);
                        }

                        gArtCurrentEntry.customTitle = newTitle;
                        // Optionally update info label
                        if (gArtInfoLabel) {
                            std::wstring infoText = L"TITLE: " + gArtCurrentEntry.DisplayTitle() + L" (custom set)\r\n";
                            infoText += L"Original title: ";
                            if (!gArtCurrentEntry.title.empty())
                                infoText += gArtCurrentEntry.title + L"\r\n";
                            else
                                infoText += L"(unable to resolve automatically)\r\n";
                            infoText += L"STORE: " + gArtCurrentEntry.store + L"\r\n";
                            infoText += L"ID: " + gArtCurrentEntry.idStr + L"\r\n";
                            infoText += L"FILE: " + (gArtCurrentEntry.fileName.empty() ? gArtCurrentEntry.expectedFileName : gArtCurrentEntry.fileName) + L"\r\n";
                            infoText += L"HAS ART: " + std::wstring(gArtCurrentEntry.hasArt ? L"Yes" : L"No");
                            std::wstring fileName = gArtCurrentEntry.expectedFileName;
                            if (fileName.empty()) fileName = ExpectedPngFromManifestId(gArtCurrentEntry.store, gArtCurrentEntry.idStr);
                            bool hasBackup = HasBackup(gArtCurrentEntry.store, fileName);
                            infoText += L"\r\nBACKUP: " + std::wstring(hasBackup ? L"Available" : L"None");
                            SetWindowTextW(gArtInfoLabel, infoText.c_str());
                        }
                        // Start web art search (simulate click on web search or reload art)
                        PostMessageW(hWnd, WM_USER + 1, 0, 0); // Triggers reload
                    }
                }
                return 0;
            }
        }
    switch (msg) {
    case WM_CREATE: {
        HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        HFONT hFontBold = CreateFontW(-14, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");

        // --- Top section: Current art preview + game info ---

        // Current art label
        CreateWindowW(L"STATIC", L"Current Art:",
            WS_CHILD | WS_VISIBLE,
            12, 10, 100, 18,
            hWnd, nullptr,
            GetModuleHandleW(nullptr), nullptr);


        // --- Top right: Game Title Search and Save ---
        static const int IDC_ART_TITLE_EDIT = 4100;
        static const int IDC_ART_TITLE_SAVE_BTN = 4101;
        HWND gArtTitleEdit = CreateWindowW(L"EDIT", (!gArtCurrentEntry.customTitle.empty() ? gArtCurrentEntry.customTitle.c_str() : (gArtCurrentEntry.title.empty() ? L"" : gArtCurrentEntry.title.c_str())),
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            650, 10, 200, 24,
            hWnd, (HMENU)(INT_PTR)IDC_ART_TITLE_EDIT,
            GetModuleHandleW(nullptr), nullptr);
        SendMessageW(gArtTitleEdit, WM_SETFONT, (WPARAM)hFont, TRUE);
        SendMessageW(gArtTitleEdit, EM_SETCUEBANNER, TRUE, (LPARAM)L"Set game title...");

        HWND gArtTitleSaveBtn = CreateWindowW(L"BUTTON", L"Save and Search",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            860, 10, 120, 24,
            hWnd, (HMENU)(INT_PTR)IDC_ART_TITLE_SAVE_BTN,
            GetModuleHandleW(nullptr), nullptr);
        SendMessageW(gArtTitleSaveBtn, WM_SETFONT, (WPARAM)hFont, TRUE);

        // Current art preview (we'll draw it ourselves)
        gArtCurrentImg = CreateWindowW(L"STATIC", L"",
            WS_CHILD | WS_VISIBLE | SS_OWNERDRAW | WS_BORDER,
            12, 30, 120, 160,
            hWnd, (HMENU)(INT_PTR)IDC_ART_CURRENT_IMG,
            GetModuleHandleW(nullptr), nullptr);

        // Game info label (use ALL CAPS for field names to make them stand out)

        std::wstring infoText = L"TITLE: " + (gArtCurrentEntry.DisplayTitle().empty() ? L"(unknown)" : gArtCurrentEntry.DisplayTitle()) + L"\r\n";
        infoText += L"Original title: ";
        if (!gArtCurrentEntry.title.empty())
            infoText += gArtCurrentEntry.title + L"\r\n";
        else
            infoText += L"(unable to resolve automatically)\r\n";
        infoText += L"STORE: " + gArtCurrentEntry.store + L"\r\n";
        infoText += L"ID: " + gArtCurrentEntry.idStr + L"\r\n";
        infoText += L"FILE: " + (gArtCurrentEntry.fileName.empty() ? gArtCurrentEntry.expectedFileName : gArtCurrentEntry.fileName) + L"\r\n";
        infoText += L"HAS ART: " + std::wstring(gArtCurrentEntry.hasArt ? L"Yes" : L"No");

        std::wstring fileName = gArtCurrentEntry.expectedFileName;
        if (fileName.empty()) {
            fileName = ExpectedPngFromManifestId(gArtCurrentEntry.store, gArtCurrentEntry.idStr);
        }
        bool hasBackup = HasBackup(gArtCurrentEntry.store, fileName);
        infoText += L"\r\nBACKUP: " + std::wstring(hasBackup ? L"Available" : L"None");

        gArtInfoLabel = CreateWindowW(L"EDIT", infoText.c_str(),
            WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
            145, 30, 350, 120,
            hWnd, (HMENU)(INT_PTR)IDC_ART_INFO_LABEL,
            GetModuleHandleW(nullptr), nullptr);

        // Restore button (only enabled if backup exists)
        HWND restoreBtn = CreateWindowW(L"BUTTON", L"Restore Original",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | (hasBackup ? 0 : WS_DISABLED),
            145, 155, 120, 28,
            hWnd, (HMENU)(INT_PTR)IDC_ART_RESTORE_BTN,
            GetModuleHandleW(nullptr), nullptr);

        // Clear cache button for this game
        CreateWindowW(L"BUTTON", L"Clear Cache",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            275, 155, 100, 28,
            hWnd, (HMENU)(INT_PTR)IDC_ART_CLEAR_CACHE_BTN,
            GetModuleHandleW(nullptr), nullptr);

        // --- Middle section: Status ---

        gArtStatus = CreateWindowW(L"STATIC", L"Loading images from SteamGridDB...",
            WS_CHILD | WS_VISIBLE,
            12, 205, 500, 20,
            hWnd, (HMENU)(INT_PTR)IDC_ART_STATUS,
            GetModuleHandleW(nullptr), nullptr);
        
        // Red warning for missing API key (hidden initially)
        gArtApiWarning = CreateWindowW(L"STATIC", L"",
            WS_CHILD | SS_LEFT,  // Hidden initially
            520, 205, 400, 20,
            hWnd, (HMENU)(INT_PTR)4019,  // IDC_ART_API_WARNING
            GetModuleHandleW(nullptr), nullptr);

        // --- Tab control for art types ---
        gArtTab = CreateWindowW(WC_TABCONTROLW, L"",
            WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | TCS_OWNERDRAWFIXED | TCS_FIXEDWIDTH,
            12, 235, 500, 32,
            hWnd, (HMENU)(INT_PTR)IDC_ART_TAB,
            GetModuleHandleW(nullptr), nullptr);
        
        // Disable theming to remove bevel effect, subclass for custom background
        if (gArtTab) {
            SetWindowTheme(gArtTab, L"", L"");
            // Set fixed tab size - must be called after TCS_FIXEDWIDTH is set
            TabCtrl_SetItemSize(gArtTab, 60, 26);
            gOriginalTabProc = (WNDPROC)SetWindowLongPtrW(gArtTab, GWLP_WNDPROC, (LONG_PTR)TabSubclassProc);
        }

        // Add tabs
        TCITEMW tie = {};
        tie.mask = TCIF_TEXT;
        tie.pszText = (LPWSTR)L"Web";
        TabCtrl_InsertItem(gArtTab, 0, &tie);
        tie.pszText = (LPWSTR)L"Grids";
        TabCtrl_InsertItem(gArtTab, 1, &tie);
        tie.pszText = (LPWSTR)L"Heroes";
        TabCtrl_InsertItem(gArtTab, 2, &tie);
        tie.pszText = (LPWSTR)L"Logos";
        TabCtrl_InsertItem(gArtTab, 3, &tie);
        tie.pszText = (LPWSTR)L"Icons";
        TabCtrl_InsertItem(gArtTab, 4, &tie);

        gArtCurrentTab = ART_WEB;
        
        // Disable tabs that require API key if not set
        bool hasApiKey = !gConfig.steamGridDbKey.empty();
        if (!hasApiKey) {
            // Disable all tabs except Web (index 0)
            TCITEM tci = {};
            tci.mask = TCIF_STATE;
            tci.dwStateMask = TCIS_BUTTONPRESSED;
            tci.dwState = TCIS_BUTTONPRESSED;
            for (int i = 1; i <= 4; i++) {
                // There's no direct disable for tabs, but we'll handle it in click logic
            }
        }

        // Web tab controls (shown by default since Web is the default tab)
        // Position below the tabs, above the list view
        gArtWebSizeLabel = CreateWindowW(L"STATIC", L"Size:",
            WS_CHILD | WS_VISIBLE,
            320, 240, 35, 20,
            hWnd, (HMENU)(INT_PTR)4018,  // IDC_ART_WEB_SIZE_LABEL
            GetModuleHandleW(nullptr), nullptr);

        gArtWebResolutionCombo = CreateWindowW(L"COMBOBOX", L"",
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | CBS_HASSTRINGS,
            355, 237, 90, 200,
            hWnd, (HMENU)(INT_PTR)IDC_ART_WEB_RES_COMBO,
            GetModuleHandleW(nullptr), nullptr);

        // Add resolution options
        SendMessageW(gArtWebResolutionCombo, CB_ADDSTRING, 0, (LPARAM)L"265x265");
        SendMessageW(gArtWebResolutionCombo, CB_ADDSTRING, 0, (LPARAM)L"512x512");
        SendMessageW(gArtWebResolutionCombo, CB_ADDSTRING, 0, (LPARAM)L"1024x1024");
        SendMessageW(gArtWebResolutionCombo, CB_ADDSTRING, 0, (LPARAM)L"2160x2160");
        SendMessageW(gArtWebResolutionCombo, CB_ADDSTRING, 0, (LPARAM)L"4096x4096");
        SendMessageW(gArtWebResolutionCombo, CB_SETCURSEL, 2, 0);  // Default to 1024x1024
        gArtWebResolution = L"1024x1024";

        // Previous page button
        gArtWebPrevBtn = CreateWindowW(L"BUTTON", L"< Prev",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_DISABLED,  // Disabled initially (page 0)
            450, 237, 50, 24,
            hWnd, (HMENU)(INT_PTR)IDC_ART_WEB_PREV_BTN,
            GetModuleHandleW(nullptr), nullptr);

        // Next page button
        gArtWebNextBtn = CreateWindowW(L"BUTTON", L"Next >",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            502, 237, 50, 24,
            hWnd, (HMENU)(INT_PTR)IDC_ART_WEB_NEXT_BTN,
            GetModuleHandleW(nullptr), nullptr);

        // --- Selection area: Image grid (left) + Preview (right) with splitter ---

        gArtList = CreateWindowW(WC_LISTVIEWW, L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_ICON | LVS_SINGLESEL | LVS_AUTOARRANGE,
            12, 270, 500, 270,
            hWnd, (HMENU)(INT_PTR)IDC_ART_LIST,
            GetModuleHandleW(nullptr), nullptr);

        ListView_SetExtendedListViewStyle(gArtList, LVS_EX_DOUBLEBUFFER);

        // Selected preview label
        CreateWindowW(L"STATIC", L"Selected Preview:",
            WS_CHILD | WS_VISIBLE,
            530, 225, 130, 18,
            hWnd, (HMENU)(INT_PTR)IDC_ART_PREVIEW_LABEL,
            GetModuleHandleW(nullptr), nullptr);

        // Selected preview (fills right side)
        gArtPreviewImg = CreateWindowW(L"STATIC", L"",
            WS_CHILD | WS_VISIBLE | SS_OWNERDRAW | WS_BORDER,
            530, 245, 250, 280,
            hWnd, (HMENU)(INT_PTR)IDC_ART_PREVIEW_IMG,
            GetModuleHandleW(nullptr), nullptr);

        // --- Debug section ---
        
        CreateWindowW(L"STATIC", L"Debug Log:",
            WS_CHILD | WS_VISIBLE,
            12, 535, 100, 18,
            hWnd, (HMENU)(INT_PTR)IDC_ART_DEBUG_LABEL,
            GetModuleHandleW(nullptr), nullptr);

        gArtDebug = CreateWindowW(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
            12, 553, 580, 80,
            hWnd, (HMENU)(INT_PTR)IDC_ART_DEBUG,
            GetModuleHandleW(nullptr), nullptr);

        // --- Bottom section: Buttons ---

        CreateWindowW(L"BUTTON", L"Apply Selected",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            610, 553, 110, 30,
            hWnd, (HMENU)(INT_PTR)IDC_ART_APPLY_BTN,
            GetModuleHandleW(nullptr), nullptr);

        CreateWindowW(L"BUTTON", L"Cancel",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            610, 593, 110, 30,
            hWnd, (HMENU)(INT_PTR)IDC_ART_CANCEL_BTN,
            GetModuleHandleW(nullptr), nullptr);

        // Config button (hidden, shown when API key missing)
        gArtConfigBtn = CreateWindowW(L"BUTTON", L"Open Config",
            WS_CHILD | BS_OWNERDRAW,  // Not visible initially
            400, 203, 100, 24,
            hWnd, (HMENU)(INT_PTR)IDC_ART_CONFIG_BTN,
            GetModuleHandleW(nullptr), nullptr);

        // Set fonts
        for (HWND c = GetWindow(hWnd, GW_CHILD); c != nullptr; c = GetWindow(c, GW_HWNDNEXT)) {
            SendMessageW(c, WM_SETFONT, (WPARAM)hFont, TRUE);
        }

        // Load current art bitmap for preview
        if (gArtCurrentEntry.hasArt && !gArtCurrentEntry.filePath.empty() && fs::exists(gArtCurrentEntry.filePath)) {
            gArtCurrentBitmap = LoadPngToHBITMAP_AspectFit(gArtCurrentEntry.filePath, 200);
        }

        // Apply dark mode to this window
        ApplyDarkModeToWindow(hWnd);
        ApplyDarkModeToControls(hWnd);

        // Start loading images
        PostMessageW(hWnd, WM_USER + 1, 0, 0);

        if (hFontBold) DeleteObject(hFontBold);
        return 0;
    }

    case WM_SIZE: {
        RECT rc;
        GetClientRect(hWnd, &rc);
        int w = rc.right - rc.left;
        int h = rc.bottom - rc.top;

        // Clamp splitter position
        if (gArtSplitterPos < 150) gArtSplitterPos = 150;
        if (gArtSplitterPos > w - 300) gArtSplitterPos = w - 300;

        // Calculate layout regions
        int topSectionHeight = 200;
        int debugHeight = 100;
        int tabTop = 235;
        int tabHeight = 32;
        int listTop = tabTop + tabHeight + 3;  // Add 3px spacing
        int listHeight = h - listTop - debugHeight - 10;
        if (listHeight < 100) listHeight = 100;

        // List width = total width minus preview panel minus splitter
        int listWidth = w - 24 - gArtSplitterPos - SPLITTER_WIDTH;
        if (listWidth < 200) listWidth = 200;

        // Preview panel position
        int previewLeft = 12 + listWidth + SPLITTER_WIDTH;
        int previewWidth = w - previewLeft - 12;

        // Info label takes full width minus current art preview and restore button
        if (gArtInfoLabel) MoveWindow(gArtInfoLabel, 145, 30, w - 145 - 140 - 12, 120, TRUE);

        // Reposition title search field and button at top right
        HWND hTitleEdit = GetDlgItem(hWnd, 4100);
        HWND hTitleSaveBtn = GetDlgItem(hWnd, 4101);
        int titleEditWidth = 200;
        int titleBtnWidth = 120;
        int spacing = 10;
        int rightEdge = w - 12;
        int btnLeft = rightEdge - titleBtnWidth;
        int editLeft = btnLeft - spacing - titleEditWidth;
        if (hTitleEdit) MoveWindow(hTitleEdit, editLeft, 10, titleEditWidth, 24, TRUE);
        if (hTitleSaveBtn) MoveWindow(hTitleSaveBtn, btnLeft, 10, titleBtnWidth, 24, TRUE);
        
        // Status bar spans full width
        if (gArtStatus) MoveWindow(gArtStatus, 12, 205, w - 24, 20, TRUE);
        
        // Tab control on the left (above list)
        if (gArtTab) MoveWindow(gArtTab, 12, tabTop, listWidth, tabHeight, TRUE);
        
        // List on the left (below tabs)
        if (gArtList) MoveWindow(gArtList, 12, listTop, listWidth, listHeight, TRUE);
        
        // Preview label and image on the right
        HWND previewLabel = GetDlgItem(hWnd, IDC_ART_PREVIEW_LABEL);
        if (previewLabel) MoveWindow(previewLabel, previewLeft, tabTop, previewWidth, 18, TRUE);
        if (gArtPreviewImg) MoveWindow(gArtPreviewImg, previewLeft, tabTop + 20, previewWidth, listHeight + tabHeight - 20, TRUE);
        
        // Debug section at the bottom
        int debugTop = listTop + listHeight + 10;
        HWND debugLabel = GetDlgItem(hWnd, IDC_ART_DEBUG_LABEL);
        if (debugLabel) MoveWindow(debugLabel, 12, debugTop, 100, 18, TRUE);
        if (gArtDebug) MoveWindow(gArtDebug, 12, debugTop + 18, w - 150, debugHeight - 28, TRUE);
        
        // Buttons on the right of debug
        HWND applyBtn = GetDlgItem(hWnd, IDC_ART_APPLY_BTN);
        HWND cancelBtn = GetDlgItem(hWnd, IDC_ART_CANCEL_BTN);
        if (applyBtn) MoveWindow(applyBtn, w - 130, debugTop + 18, 110, 30, TRUE);
        if (cancelBtn) MoveWindow(cancelBtn, w - 130, debugTop + 53, 110, 30, TRUE);

        // Invalidate splitter region for redraw
        InvalidateRect(hWnd, nullptr, TRUE);
        
        return 0;
    }

    case WM_GETMINMAXINFO: {
        MINMAXINFO* mmi = (MINMAXINFO*)l;
        mmi->ptMinTrackSize.x = 700;
        mmi->ptMinTrackSize.y = 550;
        return 0;
    }

    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        
        // Draw the splitter bar
        RECT rc;
        GetClientRect(hWnd, &rc);
        int w = rc.right - rc.left;
        int tabTop = 225;
        int tabHeight = 32;
        int listTop = tabTop + tabHeight;
        int listHeight = rc.bottom - rc.top - listTop - 110;
        if (listHeight < 100) listHeight = 100;
        
        int listWidth = w - 24 - gArtSplitterPos - SPLITTER_WIDTH;
        if (listWidth < 200) listWidth = 200;
        int splitterLeft = 12 + listWidth;
        
        // Splitter spans from tab top to bottom of list
        RECT splitterRect = { splitterLeft, tabTop, splitterLeft + SPLITTER_WIDTH, listTop + listHeight };
        
        // Draw splitter - match the background color
        COLORREF splitterColor = gDarkMode ? RGB(32, 32, 32) : RGB(255, 255, 255);
        HBRUSH brushFace = CreateSolidBrush(splitterColor);
        FillRect(hdc, &splitterRect, brushFace);
        DeleteObject(brushFace);
        
        // Draw grip lines in the middle - subtle contrast
        int midY = tabTop + (listTop + listHeight - tabTop) / 2;
        COLORREF lightColor = gDarkMode ? RGB(70, 70, 70) : RGB(220, 220, 220);
        COLORREF shadowColor = gDarkMode ? RGB(50, 50, 50) : RGB(180, 180, 180);
        HPEN penLight = CreatePen(PS_SOLID, 1, lightColor);
        HPEN penShadow = CreatePen(PS_SOLID, 1, shadowColor);
        
        for (int i = -15; i <= 15; i += 5) {
            SelectObject(hdc, penShadow);
            MoveToEx(hdc, splitterLeft + 1, midY + i, nullptr);
            LineTo(hdc, splitterLeft + SPLITTER_WIDTH - 1, midY + i);
            SelectObject(hdc, penLight);
            MoveToEx(hdc, splitterLeft + 1, midY + i + 1, nullptr);
            LineTo(hdc, splitterLeft + SPLITTER_WIDTH - 1, midY + i + 1);
        }
        
        DeleteObject(penLight);
        DeleteObject(penShadow);
        
        EndPaint(hWnd, &ps);
        return 0;
    }

    case WM_SETCURSOR: {
        // Check if cursor is over the splitter
        POINT pt;
        GetCursorPos(&pt);
        ScreenToClient(hWnd, &pt);
        
        RECT rc;
        GetClientRect(hWnd, &rc);
        int w = rc.right - rc.left;
        int tabTop = 225;
        int tabHeight = 32;
        int listTop = tabTop + tabHeight;
        int listHeight = rc.bottom - rc.top - listTop - 110;
        if (listHeight < 100) listHeight = 100;
        
        int listWidth = w - 24 - gArtSplitterPos - SPLITTER_WIDTH;
        if (listWidth < 200) listWidth = 200;
        int splitterLeft = 12 + listWidth;
        
        if (pt.x >= splitterLeft && pt.x < splitterLeft + SPLITTER_WIDTH &&
            pt.y >= tabTop && pt.y < listTop + listHeight) {
            SetCursor(LoadCursor(nullptr, IDC_SIZEWE));
            return TRUE;
        }
        break;
    }

    case WM_LBUTTONDOWN: {
        POINT pt = { LOWORD(l), HIWORD(l) };
        
        RECT rc;
        GetClientRect(hWnd, &rc);
        int w = rc.right - rc.left;
        int tabTop = 225;
        int tabHeight = 32;
        int listTop = tabTop + tabHeight;
        int listHeight = rc.bottom - rc.top - listTop - 110;
        if (listHeight < 100) listHeight = 100;
        
        int listWidth = w - 24 - gArtSplitterPos - SPLITTER_WIDTH;
        if (listWidth < 200) listWidth = 200;
        int splitterLeft = 12 + listWidth;
        
        if (pt.x >= splitterLeft && pt.x < splitterLeft + SPLITTER_WIDTH &&
            pt.y >= tabTop && pt.y < listTop + listHeight) {
            gArtSplitterDragging = true;
            SetCapture(hWnd);
        }
        return 0;
    }

    case WM_LBUTTONUP: {
        if (gArtSplitterDragging) {
            gArtSplitterDragging = false;
            ReleaseCapture();
        }
        return 0;
    }

    case WM_MOUSEMOVE: {
        if (gArtSplitterDragging) {
            RECT rc;
            GetClientRect(hWnd, &rc);
            int w = rc.right - rc.left;
            int mouseX = (short)LOWORD(l);
            
            // Calculate new splitter position (from right edge)
            int newSplitterPos = w - mouseX - 12;
            
            // Clamp to reasonable range
            if (newSplitterPos < 150) newSplitterPos = 150;
            if (newSplitterPos > w - 300) newSplitterPos = w - 300;
            
            if (newSplitterPos != gArtSplitterPos) {
                gArtSplitterPos = newSplitterPos;
                // Trigger relayout
                SendMessageW(hWnd, WM_SIZE, 0, MAKELPARAM(rc.right - rc.left, rc.bottom - rc.top));
            }
        }
        return 0;
    }

    case WM_DRAWITEM: {
        DRAWITEMSTRUCT* dis = (DRAWITEMSTRUCT*)l;
        if (dis->CtlID == IDC_ART_CURRENT_IMG) {
            DrawArtPreview(dis->hDC, dis->rcItem, gArtCurrentBitmap, L"No current art");
            return TRUE;
        }
        if (dis->CtlID == IDC_ART_PREVIEW_IMG) {
            DrawArtPreview(dis->hDC, dis->rcItem, gArtPreviewBitmap, L"Select an image");
            return TRUE;
        }
        if (dis->CtlType == ODT_BUTTON) {
            DrawDarkButton(dis);
            return TRUE;
        }
        if (dis->CtlType == ODT_TAB && dis->CtlID == IDC_ART_TAB) {
            DrawDarkTab(dis, gArtTab);
            return TRUE;
        }
        return FALSE;
    }

    case WM_NOTIFY: {
        NMHDR* nmhdr = (NMHDR*)l;
        if (nmhdr->hwndFrom == gArtList) {
            if (nmhdr->code == LVN_ITEMCHANGED) {
                NMLISTVIEW* pnmv = (NMLISTVIEW*)l;
                // Check if selection changed (new state has LVIS_SELECTED)
                if ((pnmv->uChanged & LVIF_STATE) && (pnmv->uNewState & LVIS_SELECTED)) {
                    LoadSelectedPreview(pnmv->iItem);
                }
            }
        }
        if (nmhdr->hwndFrom == gArtTab && nmhdr->code == TCN_SELCHANGE) {
            int newTab = TabCtrl_GetCurSel(gArtTab);
            
            // If no API key and trying to switch to a non-Web tab, block it
            if (gConfig.steamGridDbKey.empty() && newTab != ART_WEB) {
                // Revert to Web tab
                TabCtrl_SetCurSel(gArtTab, ART_WEB);
                SetWindowTextW(gArtApiWarning, L"Note: Grids, Logos, Heroes, & Icons require API key from SteamGridDB.");
                ShowWindow(gArtApiWarning, SW_SHOW);
                if (gArtConfigBtn) {
                    ShowWindow(gArtConfigBtn, SW_SHOW);
                }
                return 0;
            }
            
            if (newTab != gArtCurrentTab) {
                gArtCurrentTab = newTab;
                
                // Show/hide web controls based on tab
                bool isWebTab = (gArtCurrentTab == ART_WEB);
                ShowWindow(gArtWebSizeLabel, isWebTab ? SW_SHOW : SW_HIDE);
                ShowWindow(gArtWebResolutionCombo, isWebTab ? SW_SHOW : SW_HIDE);
                ShowWindow(gArtWebPrevBtn, isWebTab ? SW_SHOW : SW_HIDE);
                ShowWindow(gArtWebNextBtn, isWebTab ? SW_SHOW : SW_HIDE);
                
                // Reset web search state when switching to web tab
                if (isWebTab) {
                    gArtWebPageIndex = 0;
                    gArtWebNoMoreResults = false;
                    // Disable Prev button on page 0
                    EnableWindow(gArtWebPrevBtn, FALSE);
                    EnableWindow(gArtWebNextBtn, TRUE);
                }
                
                // Reload images for new art type
                PostMessageW(hWnd, WM_USER + 2, 0, 0);  // Use new message to load for current tab
            }
        }
        return 0;
    }

    case WM_USER + 1: {
        // Load images - initial load
        gArtImages.clear();
        gArtGameId.clear();

        // If no API key and we're on Web tab, just proceed with web search
        if (gConfig.steamGridDbKey.empty() && gArtCurrentTab == ART_WEB) {
            SetWindowTextW(gArtStatus, L"Searching web for images...");
            PostMessageW(hWnd, WM_USER + 2, 0, 0);  // Load web images
            return 0;
        }

        if (gConfig.steamGridDbKey.empty()) {
            SetWindowTextW(gArtApiWarning, L"Note: Grids, Logos, Heroes, & Icons require API key from SteamGridDB.");
            ShowWindow(gArtApiWarning, SW_SHOW);
            if (gArtConfigBtn) {
                ShowWindow(gArtConfigBtn, SW_SHOW);
            }
            return 0;
        }
        
        // Hide warning and config button if shown
        ShowWindow(gArtApiWarning, SW_HIDE);
        if (gArtConfigBtn) {
            ShowWindow(gArtConfigBtn, SW_HIDE);
        }
        
        SetWindowTextW(gArtStatus, L"Searching for game on SteamGridDB...");
        UpdateWindow(gArtStatus);

        // Determine game ID
        std::wstring store, rest;
        if (TrySplitStoreAndRest(gArtCurrentEntry.idStr, store, rest)) {
            if (store == L"steam" && IsAllDigits(rest)) {
                // Use Steam app ID to find game
                gArtGameId = SteamGridDbGetGameIdBySteamAppId(rest, gConfig.steamGridDbKey);
            }
        }

        // If no direct lookup, search by title
        if (gArtGameId.empty() && !gArtCurrentEntry.title.empty()) {
            gArtGameId = SteamGridDbSearchGame(gArtCurrentEntry.title, gConfig.steamGridDbKey);
        }

        if (gArtGameId.empty()) {
            SetWindowTextW(gArtStatus, L"Could not find game on SteamGridDB. Try setting a title first.");
            return 0;
        }

        // Now load images for current tab
        PostMessageW(hWnd, WM_USER + 2, 0, 0);
        return 0;
    }

    case WM_USER + 2: {
        // Load images for current art type tab
        gArtImages.clear();
        
        // Clear preview
        if (gArtPreviewBitmap) {
            DeleteObject(gArtPreviewBitmap);
            gArtPreviewBitmap = nullptr;
        }
        gArtSelectedIndex = -1;
        if (gArtPreviewImg) InvalidateRect(gArtPreviewImg, nullptr, TRUE);

        // Check if API key is missing - don't change the error message (not needed for Web tab)
        if (gConfig.steamGridDbKey.empty() && gArtCurrentTab != ART_WEB) {
            // Clear the list but keep the error message
            ListView_DeleteAllItems(gArtList);
            DestroyArtImageList();
            return 0;
        }

        if (gArtGameId.empty() && gArtCurrentTab != ART_WEB) {
            SetWindowTextW(gArtStatus, L"No game ID available.");
            return 0;
        }

        const wchar_t* artTypeName = L"";
        bool isWebSearch = false;
        
        switch (gArtCurrentTab) {
        case ART_GRIDS:
            artTypeName = L"grids";
            SetWindowTextW(gArtStatus, L"Fetching grids from SteamGridDB...");
            UpdateWindow(gArtStatus);
            gArtImages = SteamGridDbGetGrids(gArtGameId, gConfig.steamGridDbKey);
            break;
        case ART_HEROES:
            artTypeName = L"heroes";
            SetWindowTextW(gArtStatus, L"Fetching heroes from SteamGridDB...");
            UpdateWindow(gArtStatus);
            gArtImages = SteamGridDbGetHeroes(gArtGameId, gConfig.steamGridDbKey);
            break;
        case ART_LOGOS:
            artTypeName = L"logos";
            SetWindowTextW(gArtStatus, L"Fetching logos from SteamGridDB...");
            UpdateWindow(gArtStatus);
            gArtImages = SteamGridDbGetLogos(gArtGameId, gConfig.steamGridDbKey);
            break;
        case ART_ICONS:
            artTypeName = L"icons";
            SetWindowTextW(gArtStatus, L"Fetching icons from SteamGridDB...");
            UpdateWindow(gArtStatus);
            gArtImages = SteamGridDbGetIcons(gArtGameId, gConfig.steamGridDbKey);
            break;
        case ART_WEB:
            isWebSearch = true;
            artTypeName = L"web images";
            {
                std::wstring statusMsg = L"Searching web for game art (page " + std::to_wstring(gArtWebPageIndex + 1) + L")...";
                SetWindowTextW(gArtStatus, statusMsg.c_str());
            }
            UpdateWindow(gArtStatus);
            // Get resolution from combo
            {
                int sel = (int)SendMessageW(gArtWebResolutionCombo, CB_GETCURSEL, 0, 0);
                if (sel != CB_ERR) {
                    wchar_t buf[32];
                    SendMessageW(gArtWebResolutionCombo, CB_GETLBTEXT, sel, (LPARAM)buf);
                    gArtWebResolution = buf;
                }
            }
            gArtImages = WebSearchGameImages(gArtCurrentEntry.title, gArtWebResolution, gArtWebPageIndex);
            // Update button states
            EnableWindow(gArtWebPrevBtn, gArtWebPageIndex > 0);
            // Disable next if no results found (end of results)
            gArtWebNoMoreResults = gArtImages.empty();
            EnableWindow(gArtWebNextBtn, !gArtWebNoMoreResults);
            break;
        }

        // Filter out small/blurry images based on minImageSize setting (skip for web search, already filtered)
        size_t totalBeforeFilter = gArtImages.size();
        size_t filteredOut = 0;
        if (!isWebSearch) {
            gArtImages = FilterImagesBySize(gArtImages, gConfig.minImageSize);
            filteredOut = totalBeforeFilter - gArtImages.size();
        }

        if (gArtImages.empty()) {
            std::wstring msg;
            if (isWebSearch) {
                if (gArtWebPageIndex == 0) {
                    msg = L"No web images found for \"" + gArtCurrentEntry.title + L"\".";
                } else {
                    msg = L"No more results on page " + std::to_wstring(gArtWebPageIndex + 1) + L".";
                }
            } else {
                msg = L"No " + std::wstring(artTypeName) + L" found for this game on SteamGridDB.";
                if (filteredOut > 0) {
                    msg += L" (" + std::to_wstring(filteredOut) + L" images filtered out as too small)";
                }
            }
            SetWindowTextW(gArtStatus, msg.c_str());
            // Clear the list
            ListView_DeleteAllItems(gArtList);
            DestroyArtImageList();
            return 0;
        }

        std::wstring statusText;
        if (isWebSearch) {
            statusText = L"Page " + std::to_wstring(gArtWebPageIndex + 1) + L": Found " + std::to_wstring(gArtImages.size()) + L" " + artTypeName;
        } else {
            statusText = L"Found " + std::to_wstring(gArtImages.size()) + L" " + artTypeName;
            if (filteredOut > 0) {
                statusText += L" (" + std::to_wstring(filteredOut) + L" too small)";
            }
        }
        statusText += L". Loading thumbnails...";
        SetWindowTextW(gArtStatus, statusText.c_str());

        PopulateArtList();
        return 0;
    }

    case WM_COMMAND: {
        int id = LOWORD(w);
        if (id == IDC_ART_APPLY_BTN) {
            ApplySelectedArt();
            return 0;
        }
        if (id == IDC_ART_CANCEL_BTN) {
            CloseArtWindow(false);
            return 0;
        }
        if (id == IDC_ART_RESTORE_BTN) {
            RestoreOriginalArt();
            return 0;
        }
        if (id == IDC_ART_CLEAR_CACHE_BTN) {
            int result = DarkMessageBox(gArtWnd, 
                L"This will delete all cached thumbnails for this game.\nThey will be re-downloaded when you browse art again.\n\nContinue?",
                L"Clear Game Cache", MB_YESNO | MB_ICONQUESTION);
            if (result == IDYES) {
                int deleted = ClearGameCache(gArtGameId);
                wchar_t msg[128];
                swprintf(msg, 128, L"Deleted %d cached images for this game.", deleted);
                DarkMessageBox(gArtWnd, msg, L"Cache Cleared", MB_OK | MB_ICONINFORMATION);
            }
            return 0;
        }
        if (id == IDC_ART_CONFIG_BTN) {
            // Close art window, open config
            CloseArtWindow(false);
            OpenConfigWindow(gMainWnd);
            return 0;
        }
        if (id == IDC_ART_WEB_PREV_BTN) {
            // Previous page
            if (gArtWebPageIndex > 0) {
                gArtWebPageIndex--;
                PostMessageW(hWnd, WM_USER + 2, 0, 0);
            }
            return 0;
        }
        if (id == IDC_ART_WEB_NEXT_BTN) {
            // Next page
            gArtWebPageIndex++;
            PostMessageW(hWnd, WM_USER + 2, 0, 0);
            return 0;
        }
        if (id == IDC_ART_WEB_RES_COMBO && HIWORD(w) == CBN_SELCHANGE) {
            // Resolution changed - reset search and reload
            gArtWebPageIndex = 0;
            gArtWebNoMoreResults = false;
            if (gArtCurrentTab == ART_WEB) {
                PostMessageW(hWnd, WM_USER + 2, 0, 0);
            }
            return 0;
        }
        return 0;
    }

    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)w;
        HWND hwndCtrl = (HWND)l;
        
        // Check if this is the API warning control - always paint in red
        if (hwndCtrl == gArtApiWarning) {
            SetTextColor(hdc, gDarkMode ? RGB(255, 80, 80) : RGB(220, 0, 0));
            SetBkColor(hdc, gDarkMode ? RGB(32, 32, 32) : RGB(255, 255, 255));
            return (LRESULT)(gDarkMode ? gDarkBgBrush : gLightBgBrush);
        }
        
        // Check if this is the status control with error message
        if (hwndCtrl == gArtStatus) {
            wchar_t text[256]{};
            GetWindowTextW(gArtStatus, text, _countof(text));
            if (wcsstr(text, L"ERROR") != nullptr || wcsstr(text, L"[!]") != nullptr) {
                // Red text for errors
                SetTextColor(hdc, RGB(255, 80, 80));
                SetBkColor(hdc, gDarkMode ? RGB(32, 32, 32) : RGB(255, 255, 255));
                return (LRESULT)(gDarkMode ? gDarkBgBrush : gLightBgBrush);
            }
        }
        
        if (gDarkMode) {
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkColor(hdc, RGB(32, 32, 32));
            return (LRESULT)gDarkBgBrush;
        } else {
            SetTextColor(hdc, RGB(0, 0, 0));
            SetBkColor(hdc, RGB(255, 255, 255));
            return (LRESULT)gLightBgBrush;
        }
    }
    
    case WM_CTLCOLOREDIT: {
        if (gDarkMode) {
            HDC hdc = (HDC)w;
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkColor(hdc, RGB(45, 45, 45));
            return (LRESULT)gDarkEditBrush;
        }
        break;
    }

    case WM_ERASEBKGND: {
        HDC hdc = (HDC)w;
        RECT rc;
        GetClientRect(hWnd, &rc);
        if (gDarkMode) {
            FillRect(hdc, &rc, gDarkBgBrush);
        } else {
            if (!gLightBgBrush) gLightBgBrush = CreateSolidBrush(RGB(255, 255, 255));
            FillRect(hdc, &rc, gLightBgBrush);
        }
        return 1;
    }

    case WM_CLOSE:
        CloseArtWindow(false);
        return 0;
    }
    return DefWindowProcW(hWnd, msg, w, l);
}

static void CloseArtWindow(bool applyChanges) {
    (void)applyChanges;
    if (!gArtWnd) return;

    // If a custom title was set during this session, reload the cache and grid exactly as at startup
    if (sCustomTitleChanged && gMainWnd) {
        CacheState st;
        if (LoadCacheFile(st)) {
            gConfig = st.config;
            gDarkMode = gConfig.darkMode;
            ApplyCombosFromState(st);
            RecreateListView(gMainWnd);
            gLastItems = st.items;
            PopulateListFromItems(gLastItems);
        }
        sCustomTitleChanged = false;
    }

    // Cancel any loading in progress
    gArtCancelLoading = true;
    gArtIsLoading = false;

    EnableWindow(gArtParent, TRUE);
    SetForegroundWindow(gArtParent);

    DestroyArtImageList();

    // Clean up current art bitmap
    if (gArtCurrentBitmap) {
        DeleteObject(gArtCurrentBitmap);
        gArtCurrentBitmap = nullptr;
    }

    // Clean up preview bitmap
    if (gArtPreviewBitmap) {
        DeleteObject(gArtPreviewBitmap);
        gArtPreviewBitmap = nullptr;
    }
    gArtSelectedIndex = -1;
    gArtGameId.clear();
    gArtCurrentTab = ART_WEB;

    DestroyWindow(gArtWnd);
    gArtWnd = nullptr;
    gArtList = nullptr;
    gArtTab = nullptr;
    gArtParent = nullptr;
    gArtStatus = nullptr;
    gArtCurrentImg = nullptr;
    gArtInfoLabel = nullptr;
    gArtDebug = nullptr;
    gArtPreviewImg = nullptr;
    gArtConfigBtn = nullptr;
    gArtImages.clear();
}

static void OpenArtWindow(HWND parent, const GameEntry& entry) {
    if (gArtWnd) {
        SetForegroundWindow(gArtWnd);
        return;
    }

    gArtCurrentEntry = entry;
    gArtParent = parent;
    EnableWindow(parent, FALSE);

    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = ArtWndProc;
    wc.hInstance = GetModuleHandleW(nullptr);
    wc.lpszClassName = L"XboxAppArtUpdaterArtWnd";
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hIcon = LoadIconW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(101));
    wc.hIconSm = LoadIconW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(101));
    RegisterClassExW(&wc);

    RECT pr{};
    GetWindowRect(parent, &pr);
    int w = 950, h = 750;
    int x = pr.left + ((pr.right - pr.left) - w) / 2;
    int y = pr.top + ((pr.bottom - pr.top) - h) / 2;

    std::wstring title = L"Update Art: " + (entry.title.empty() ? entry.idStr : entry.title);

    gArtWnd = CreateWindowExW(
        0,
        wc.lpszClassName,
        title.c_str(),
        WS_OVERLAPPEDWINDOW | WS_POPUP,
        x, y, w, h,
        parent, nullptr, GetModuleHandleW(nullptr), nullptr);

    ShowWindow(gArtWnd, SW_SHOW);
    UpdateWindow(gArtWnd);
}

static void OnComboChanged(HWND hWnd, int ctrlId) {
    if (ctrlId == IDC_CMB_SIZE) {
        int sel = (int)SendMessageW(gCmbSize, CB_GETCURSEL, 0, 0);
        int newSize = (sel == 0) ? 32 : (sel == 2) ? 128 : 64;

        if (newSize != gIconSize) {
            gIconSize = newSize;
            DestroyImageList();
        }

        RecreateListView(hWnd);
        PopulateListFromItems(gLastItems);
        SaveUiStateToCacheFile();
    }
    else if (ctrlId == IDC_CMB_LAYOUT) {
        int sel = (int)SendMessageW(gCmbLayout, CB_GETCURSEL, 0, 0);
        bool newGrid = (sel == 1);

        if (newGrid != gGridLayout) {
            gGridLayout = newGrid;
            DestroyImageList();
        }

        RecreateListView(hWnd);
        PopulateListFromItems(gLastItems);
        SaveUiStateToCacheFile();
    }
}

// ------------------------------------------------------------
// Config window
// ------------------------------------------------------------

static void CloseConfigWindow(bool applyChanges) {
    if (!gCfgWnd) return;

    bool themeChanged = false;
    
    if (applyChanges) {
        // Read values from edits
        wchar_t bufKey[2048]{};
        wchar_t bufPath[4096]{};

        GetWindowTextW(gCfgKeyEdit, bufKey, _countof(bufKey));
        GetWindowTextW(gCfgPathEdit, bufPath, _countof(bufPath));

        gConfig.steamGridDbKey = bufKey;
        gConfig.thirdPartyBasePath = bufPath;

        // Read theme selection
        int themeSel = (int)SendMessageW(gCfgTheme, CB_GETCURSEL, 0, 0);
        bool newDarkMode = (themeSel == 0);
        if (newDarkMode != gConfig.darkMode) {
            themeChanged = true;
        }
        gConfig.darkMode = newDarkMode;
        gDarkMode = gConfig.darkMode;

        // Read parallel downloads selection
        int parallelSel = (int)SendMessageW(gCfgParallel, CB_GETCURSEL, 0, 0);
        if (parallelSel == 0) {
            gConfig.parallelDownloads = 0; // Auto
        } else {
            gConfig.parallelDownloads = parallelSel * 2; // 1->2, 2->4, 3->6, etc.
        }

        // Read min image size
        wchar_t bufMinSize[32]{};
        GetWindowTextW(gCfgMinSizeEdit, bufMinSize, _countof(bufMinSize));
        gConfig.minImageSize = _wtoi(bufMinSize);
        if (gConfig.minImageSize < 0) gConfig.minImageSize = 0;

        // Persist to cache immediately
        SaveUiStateToCacheFile();
    }

    HWND parent = gCfgParent;
    
    EnableWindow(gCfgParent, TRUE);
    SetForegroundWindow(gCfgParent);

    DestroyWindow(gCfgWnd);
    gCfgWnd = nullptr;
    gCfgKeyEdit = nullptr;
    gCfgPathEdit = nullptr;
    gCfgTheme = nullptr;
    gCfgParallel = nullptr;
    gCfgMinSizeEdit = nullptr;
    gCfgParent = nullptr;
    
    // Apply theme change to parent window if theme changed
    if (themeChanged && parent) {
        ApplyDarkModeToWindow(parent);
        ApplyDarkModeToControls(parent);
        InvalidateRect(parent, nullptr, TRUE);
    }
}

static LRESULT CALLBACK ConfigWndProc(HWND hWnd, UINT msg, WPARAM w, LPARAM l) {
    switch (msg) {
    case WM_CREATE: {
        HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

        CreateWindowW(L"STATIC", L"SteamGridDB API Key:",
            WS_CHILD | WS_VISIBLE,
            12, 14, 160, 20,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_KEY_LABEL,
            GetModuleHandleW(nullptr), nullptr);

        gCfgKeyEdit = CreateWindowW(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            12, 36, 540, 24,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_KEY_EDIT,
            GetModuleHandleW(nullptr), nullptr);

        CreateWindowW(L"STATIC", L"ThirdPartyLibraries base path (optional):",
            WS_CHILD | WS_VISIBLE,
            12, 74, 320, 20,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_PATH_LABEL,
            GetModuleHandleW(nullptr), nullptr);

        gCfgPathEdit = CreateWindowW(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            12, 96, 470, 24,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_PATH_EDIT,
            GetModuleHandleW(nullptr), nullptr);

        CreateWindowW(L"BUTTON", L"Browse...",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            490, 96, 62, 24,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_BROWSE_BTN,
            GetModuleHandleW(nullptr), nullptr);

        CreateWindowW(L"STATIC", L"Theme:",
            WS_CHILD | WS_VISIBLE,
            12, 136, 50, 20,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_THEME_LABEL,
            GetModuleHandleW(nullptr), nullptr);

        gCfgTheme = CreateWindowW(WC_COMBOBOXW, L"",
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            70, 132, 120, 200,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_CMB_THEME,
            GetModuleHandleW(nullptr), nullptr);
        SendMessageW(gCfgTheme, CB_ADDSTRING, 0, (LPARAM)L"Dark");
        SendMessageW(gCfgTheme, CB_ADDSTRING, 0, (LPARAM)L"Light");
        SendMessageW(gCfgTheme, CB_SETCURSEL, gConfig.darkMode ? 0 : 1, 0);

        CreateWindowW(L"STATIC", L"Parallel Downloads:",
            WS_CHILD | WS_VISIBLE,
            210, 136, 130, 20,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_PARALLEL_LABEL,
            GetModuleHandleW(nullptr), nullptr);

        gCfgParallel = CreateWindowW(WC_COMBOBOXW, L"",
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            340, 132, 80, 200,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_CMB_PARALLEL,
            GetModuleHandleW(nullptr), nullptr);
        // Add options: Auto, 2, 4, 6, 8, ... up to CPU thread count
        unsigned int maxThreads = std::thread::hardware_concurrency();
        if (maxThreads < 2) maxThreads = 8; // Fallback if detection fails
        SendMessageW(gCfgParallel, CB_ADDSTRING, 0, (LPARAM)L"Auto");
        for (unsigned int i = 2; i <= maxThreads; i += 2) {
            wchar_t buf[16];
            swprintf(buf, 16, L"%u", i);
            SendMessageW(gCfgParallel, CB_ADDSTRING, 0, (LPARAM)buf);
        }
        // Select current setting
        if (gConfig.parallelDownloads == 0) {
            SendMessageW(gCfgParallel, CB_SETCURSEL, 0, 0); // Auto
        } else {
            // Find index: 0=Auto, 1=2, 2=4, 3=6, etc., so index = parallelDownloads/2
            int idx = gConfig.parallelDownloads / 2;
            if (idx < 1) idx = 1;
            int count = (int)SendMessageW(gCfgParallel, CB_GETCOUNT, 0, 0);
            if (idx >= count) idx = count - 1;
            SendMessageW(gCfgParallel, CB_SETCURSEL, idx, 0);
        }

        CreateWindowW(L"STATIC", L"Min Image Size (px):",
            WS_CHILD | WS_VISIBLE,
            12, 168, 130, 20,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_MINSIZE_LABEL,
            GetModuleHandleW(nullptr), nullptr);

        gCfgMinSizeEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_NUMBER,
            145, 166, 60, 22,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_MINSIZE_EDIT,
            GetModuleHandleW(nullptr), nullptr);
        wchar_t minSizeBuf[16];
        swprintf(minSizeBuf, 16, L"%d", gConfig.minImageSize);
        SetWindowTextW(gCfgMinSizeEdit, minSizeBuf);

        CreateWindowW(L"BUTTON", L"Clear Image Cache",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            12, 202, 130, 28,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_CLEAR_CACHE_BTN,
            GetModuleHandleW(nullptr), nullptr);

        CreateWindowW(L"BUTTON", L"Restore Original Art",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            152, 202, 145, 28,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_RESTORE_ART_BTN,
            GetModuleHandleW(nullptr), nullptr);

        CreateWindowW(L"BUTTON", L"Cancel",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            372, 242, 80, 28,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_CANCEL_BTN,
            GetModuleHandleW(nullptr), nullptr);

        CreateWindowW(L"BUTTON", L"OK",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            472, 242, 80, 28,
            hWnd, (HMENU)(INT_PTR)IDC_CFG_OK_BTN,
            GetModuleHandleW(nullptr), nullptr);

        // Set fonts + initial values
        for (HWND c = GetWindow(hWnd, GW_CHILD); c != nullptr; c = GetWindow(c, GW_HWNDNEXT)) {
            SendMessageW(c, WM_SETFONT, (WPARAM)hFont, TRUE);
        }

        SetWindowTextW(gCfgKeyEdit, gConfig.steamGridDbKey.c_str());
        SetWindowTextW(gCfgPathEdit, gConfig.thirdPartyBasePath.c_str());
        
        // Set placeholders using custom dark mode placeholder
        SetEditPlaceholder(gCfgKeyEdit, L"e.g., abcd1234efgh5678...");
        SetEditPlaceholder(gCfgPathEdit, L"e.g., C:\\XboxGames");

        // Apply dark mode
        ApplyDarkModeToWindow(hWnd);
        ApplyDarkModeToControls(hWnd);

        return 0;
    }

    case WM_COMMAND: {
        int id = LOWORD(w);
        if (id == IDC_CFG_BROWSE_BTN) {
            std::wstring picked = BrowseForFolder(hWnd, L"Select ThirdPartyLibraries base folder");
            if (!picked.empty()) {
                SetWindowTextW(gCfgPathEdit, picked.c_str());
            }
            return 0;
        }
        if (id == IDC_CFG_CLEAR_CACHE_BTN) {
            int result = DarkMessageBox(hWnd, 
                L"This will delete all cached thumbnails.\nYou will need to re-download them when browsing art.\n\nContinue?",
                L"Clear Image Cache", MB_YESNO | MB_ICONQUESTION);
            if (result == IDYES) {
                int deleted = ClearAllImageCache();
                wchar_t msg[128];
                swprintf(msg, 128, L"Deleted %d cached images.", deleted);
                DarkMessageBox(hWnd, msg, L"Cache Cleared", MB_OK | MB_ICONINFORMATION);
            }
            return 0;
        }
        if (id == IDC_CFG_RESTORE_ART_BTN) {
            int result = DarkMessageBox(hWnd, 
                L"This will restore all original art from backups.\n\nGames with backups: Restored to original art\nGames without backups: Custom art will be removed\n\nContinue?",
                L"Restore Original Art", MB_YESNO | MB_ICONWARNING);
            if (result == IDYES) {
                int restored = RestoreAllOriginalArt();
                wchar_t msg[128];
                swprintf(msg, 128, L"Restored/removed %d art files.\n\nPlease restart the Xbox app to see changes.", restored);
                DarkMessageBox(hWnd, msg, L"Restore Complete", MB_OK | MB_ICONINFORMATION);
                
                // Trigger a rescan to update the UI
                if (gMainWnd) {
                    PostMessageW(gMainWnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_SCAN, 0), 0);
                }
            }
            return 0;
        }
        if (id == IDC_CFG_OK_BTN) {
            CloseConfigWindow(true);
            return 0;
        }
        if (id == IDC_CFG_CANCEL_BTN) {
            CloseConfigWindow(false);
            return 0;
        }
        return 0;
    }

    case WM_DRAWITEM: {
        DRAWITEMSTRUCT* dis = (DRAWITEMSTRUCT*)l;
        if (dis->CtlType == ODT_BUTTON) {
            DrawDarkButton(dis);
            return TRUE;
        }
        if (dis->CtlType == ODT_COMBOBOX) {
            DrawDarkComboItem(dis);
            return TRUE;
        }
        break;
    }

    case WM_CTLCOLORLISTBOX: {
        // For combobox dropdown list
        HDC hdc = (HDC)w;
        if (gDarkMode) {
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkColor(hdc, RGB(45, 45, 45));
            return (LRESULT)gDarkEditBrush;
        } else {
            SetTextColor(hdc, RGB(0, 0, 0));
            SetBkColor(hdc, RGB(255, 255, 255));
            return (LRESULT)gLightBgBrush;
        }
    }

    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)w;
        if (gDarkMode) {
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkColor(hdc, RGB(32, 32, 32));
            return (LRESULT)gDarkBgBrush;
        } else {
            SetTextColor(hdc, RGB(0, 0, 0));
            SetBkColor(hdc, RGB(255, 255, 255));
            return (LRESULT)gLightBgBrush;
        }
    }
    
    case WM_CTLCOLOREDIT: {
        HDC hdc = (HDC)w;
        if (gDarkMode) {
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkColor(hdc, RGB(45, 45, 45));
            return (LRESULT)gDarkEditBrush;
        }
        break;
    }

    case WM_ERASEBKGND: {
        HDC hdc = (HDC)w;
        RECT rc;
        GetClientRect(hWnd, &rc);
        if (gDarkMode) {
            FillRect(hdc, &rc, gDarkBgBrush);
        } else {
            if (!gLightBgBrush) gLightBgBrush = CreateSolidBrush(RGB(255, 255, 255));
            FillRect(hdc, &rc, gLightBgBrush);
        }
        return 1;
    }

    case WM_CLOSE:
        CloseConfigWindow(false);
        return 0;
    }
    return DefWindowProcW(hWnd, msg, w, l);
}

static void OpenConfigWindow(HWND parent) {
    if (gCfgWnd) {
        SetForegroundWindow(gCfgWnd);
        return;
    }

    gCfgParent = parent;
    EnableWindow(parent, FALSE);

    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = ConfigWndProc;
    wc.hInstance = GetModuleHandleW(nullptr);
    wc.lpszClassName = L"XboxAppArtUpdaterConfigWnd";
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hIcon = LoadIconW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(101));
    wc.hIconSm = LoadIconW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(101));
    RegisterClassExW(&wc);

    RECT pr{};
    GetWindowRect(parent, &pr);
    int w = 580, h = 325;
    int x = pr.left + ((pr.right - pr.left) - w) / 2;
    int y = pr.top + ((pr.bottom - pr.top) - h) / 2;

    gCfgWnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        wc.lpszClassName,
        L"Config",
        WS_CAPTION | WS_SYSMENU | WS_POPUP,
        x, y, w, h,
        parent, nullptr, GetModuleHandleW(nullptr), nullptr);

    ShowWindow(gCfgWnd, SW_SHOW);
    UpdateWindow(gCfgWnd);
}

// ------------------------------------------------------------
// Help dialog
// ------------------------------------------------------------

// Helper to replace username in path with %USERNAME%
static std::wstring ReplaceUsernameInPath(const std::wstring& path) {
    wchar_t username[256]{};
    DWORD size = _countof(username);
    if (GetUserNameW(username, &size) && wcslen(username) > 0) {
        std::wstring result = path;
        std::wstring userStr = username;
        size_t pos = result.find(L"\\Users\\" + userStr + L"\\");
        if (pos != std::wstring::npos) {
            result.replace(pos, 7 + userStr.length() + 1, L"\\Users\\%USERNAME%\\");
        }
        return result;
    }
    return path;
}

// Help window proc
static HWND gHelpWnd = nullptr;
static HWND gHelpEdit = nullptr;
static HWND gHelpParent = nullptr;

// Helper to add text to RichEdit with optional bold (preserves dark mode color)
static void AppendRichText(HWND hRich, const wchar_t* text, bool bold = false) {
    // Move to end
    int len = GetWindowTextLengthW(hRich);
    SendMessageW(hRich, EM_SETSEL, len, len);
    
    // Set format (bold + color for dark mode)
    // Use CHARFORMAT with explicit size - more compatible
    CHARFORMAT cf{};
    cf.cbSize = sizeof(CHARFORMAT);
    cf.dwMask = CFM_BOLD | CFM_COLOR;
    cf.dwEffects = bold ? CFE_BOLD : 0;
    cf.crTextColor = gDarkMode ? RGB(255, 255, 255) : RGB(0, 0, 0);
    SendMessageW(hRich, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
    
    // Insert text
    SendMessageW(hRich, EM_REPLACESEL, FALSE, (LPARAM)text);
}

static LRESULT CALLBACK HelpWndProc(HWND hWnd, UINT msg, WPARAM w, LPARAM l) {
    switch (msg) {
    case WM_CREATE: {
        // Build help text with paths (using %USERNAME%)
        std::wstring defaultPath = ReplaceUsernameInPath(GetDefaultThirdPartyLibrariesPath());
        std::wstring backupPath = ReplaceUsernameInPath(GetBackupArtPath());
        
        // Build help text as a single string
        std::wstring helpText;
        helpText += L"XBOX APP ART UPDATER - HELP & FAQ\r\n";
        helpText += L"==================================\r\n\r\n";
        helpText += L"QUICK START - AUTOMATED WEB SEARCH:\r\n\r\n";
        helpText += L"Quick Apply now includes automatic web search using Google Images.\r\n";
        helpText += L"When enabled, it fetches game covers directly from the web without\r\n";
        helpText += L"needing a SteamGridDB API key.\r\n\r\n";
        helpText += L"IMPORTANT NOTES ABOUT WEB SEARCH:\r\n";
        helpText += L"- Web results are NOT curated and may be wrong or low quality\r\n";
        helpText += L"- Results can be unpredictable and vary between searches\r\n";
        helpText += L"- You can choose 'Random' or select a specific result number\r\n";
        helpText += L"- First results (1-3) are usually more accurate than later ones\r\n";
        helpText += L"- Use specific result numbers for more consistent selection\r\n";
        helpText += L"- For best quality, use SteamGridDB sources instead\r\n\r\n";
        helpText += L"==================================\r\n\r\n";
        helpText += L"HOW TO USE THIS APP:\r\n\r\n";
        helpText += L"1. GET A STEAMGRIDDB API KEY (Optional for web-only mode)\r\n";
        helpText += L"   - Go to https://www.steamgriddb.com\r\n";
        helpText += L"   - Create an account or sign in\r\n";
        helpText += L"   - Go to your Profile > Preferences > API\r\n";
        helpText += L"   - Copy your API key\r\n\r\n";
        helpText += L"2. SET UP THE APP\r\n";
        helpText += L"   - Click the \"Config\" button (top right)\r\n";
        helpText += L"   - Paste your SteamGridDB API key\r\n";
        helpText += L"   - Optionally set a custom ThirdPartyLibraries path\r\n";
        helpText += L"   - Choose Theme (Dark/Light) and Parallel Downloads\r\n";
        helpText += L"   - Click OK to save\r\n\r\n";
        helpText += L"3. SCAN FOR GAMES\r\n";
        helpText += L"   - Click the \"Scan\" button to find games\r\n";
        helpText += L"   - The app will scan Steam, Epic, GOG, and Ubisoft\r\n";
        helpText += L"   - Games will appear in the list with their current art\r\n\r\n";
        helpText += L"4. UPDATE GAME ART\r\n";
        helpText += L"   - Double-click any game in the list\r\n";
        helpText += L"   - Browse available art from SteamGridDB\r\n";
        helpText += L"   - Use tabs to switch between Grids, Heroes, Logos, Icons\r\n";
        helpText += L"   - Select an image and click \"Apply\" to update\r\n\r\n";
        helpText += L"5. RESTORE ORIGINAL ART\r\n";
        helpText += L"   - Open the art window for a game\r\n";
        helpText += L"   - Click \"Restore Original\" to revert changes\r\n";
        helpText += L"   - Only works if a backup was created\r\n\r\n";
        helpText += L"==================================\r\n\r\n";
        helpText += L"FILTERS:\r\n";
        helpText += L"- Store: Filter by game store (Steam, Epic, GOG, Ubisoft)\r\n";
        helpText += L"- Art: Show all games, only games with art, or missing art\r\n";
        helpText += L"- Search: Type to filter games by title\r\n\r\n";
        helpText += L"TIPS:\r\n";
        helpText += L"- Games with \"(missing)\" have no custom art yet\r\n";
        helpText += L"- Art is cached locally for faster loading\r\n";
        helpText += L"- Switch between List and Grid view using the Layout dropdown\r\n";
        helpText += L"- Change icon sizes with the Icon size dropdown\r\n\r\n";
        helpText += L"==================================\r\n\r\n";
        helpText += L"PARALLEL DOWNLOADS:\r\n";
        helpText += L"- Configure how many images download simultaneously\r\n";
        helpText += L"- Set in Config > Parallel Downloads dropdown\r\n";
        helpText += L"- \"Auto\" uses your CPU thread count for best performance\r\n";
        helpText += L"- Lower values (2-4) may help on slower connections\r\n\r\n";
        helpText += L"MINIMUM IMAGE SIZE:\r\n";
        helpText += L"- Filter out small/blurry images from search results\r\n";
        helpText += L"- Set in Config > Min Image Size (px) field\r\n";
        helpText += L"- Default is 200px - images smaller than this are hidden\r\n";
        helpText += L"- Set to 0 to show all images regardless of size\r\n";
        helpText += L"- Affects both manual art selection and Quick Apply\r\n\r\n";
        helpText += L"QUICK APPLY:\r\n";
        helpText += L"- Select one or more games in the list\r\n";
        helpText += L"- Click \"Quick Apply\" to apply art automatically\r\n";
        helpText += L"- Choose sources: Web Search, Grids, Heroes, Logos, Icons\r\n";
        helpText += L"- Web Search: Automatic Google Images search (may be unpredictable)\r\n";
        helpText += L"  * Random: Picks a random result from search\r\n";
        helpText += L"  * Result #: Select specific result (1=first, usually best)\r\n";
        helpText += L"  * First 3 results are typically most accurate\r\n";
        helpText += L"- SteamGridDB: Curated, high-quality community art\r\n";
        helpText += L"- Shows results with success/failure details\r\n";
        helpText += L"- Debug window (single game) shows URLs and selection\r\n";
        helpText += L"- Uses parallel processing for faster batch updates\r\n\r\n";
        helpText += L"READ-ONLY PROTECTION:\r\n";
        helpText += L"- Applied art is automatically marked as read-only\r\n";
        helpText += L"- Prevents Xbox app from overwriting your custom art\r\n";
        helpText += L"- Files are unlocked temporarily when updating art\r\n";
        helpText += L"- Restoring original art also removes read-only flag\r\n\r\n";
        helpText += L"CACHE MANAGEMENT:\r\n";
        helpText += L"- Thumbnails are cached locally for faster loading\r\n";
        helpText += L"- Clear all cache: Config > \"Clear Image Cache\" button\r\n";
        helpText += L"- Clear per-game: Art window > \"Clear Cache\" button\r\n";
        helpText += L"- Cache is stored in %TEMP%\\XboxAppUpdaterArtCache\r\n";
        helpText += L"- Safe to clear while browsing - images auto re-download\r\n";
        helpText += L"- Per-game clear only affects that game's thumbnails\r\n\r\n";
        helpText += L"==================================\r\n\r\n";
        helpText += L"PATHS (select and copy):\r\n\r\n";
        helpText += L"Default ThirdPartyLibraries:\r\n";
        helpText += defaultPath + L"\r\n\r\n";
        helpText += L"Art Backups Location:\r\n";
        helpText += backupPath;
        
        // Create a simple multiline edit control (easier to style)
        gHelpEdit = CreateWindowW(L"EDIT", helpText.c_str(),
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
            10, 10, 560, 420,
            hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
        
        // Apply dark mode theme to the edit control for dark scrollbar
        if (gDarkMode) {
            SetWindowTheme(gHelpEdit, L"DarkMode_Explorer", nullptr);
        }
        
        CreateWindowW(L"BUTTON", L"Close",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            250, 440, 80, 28,
            hWnd, (HMENU)1, GetModuleHandleW(nullptr), nullptr);
        
        HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        for (HWND c = GetWindow(hWnd, GW_CHILD); c != nullptr; c = GetWindow(c, GW_HWNDNEXT)) {
            SendMessageW(c, WM_SETFONT, (WPARAM)hFont, TRUE);
        }
        
        ApplyDarkModeToWindow(hWnd);
        return 0;
    }
    
    case WM_COMMAND:
        if (LOWORD(w) == 1) {
            DestroyWindow(hWnd);
            return 0;
        }
        break;
    
    case WM_DRAWITEM: {
        DRAWITEMSTRUCT* dis = (DRAWITEMSTRUCT*)l;
        if (dis->CtlType == ODT_BUTTON) {
            DrawDarkButton(dis);
            return TRUE;
        }
        break;
    }
    
    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)w;
        if (gDarkMode) {
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkColor(hdc, RGB(32, 32, 32));
            return (LRESULT)gDarkBgBrush;
        } else {
            SetTextColor(hdc, RGB(0, 0, 0));
            SetBkColor(hdc, RGB(255, 255, 255));
            return (LRESULT)gLightBgBrush;
        }
    }
    
    case WM_CTLCOLOREDIT: {
        HDC hdc = (HDC)w;
        if (gDarkMode) {
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkColor(hdc, RGB(45, 45, 45));
            return (LRESULT)gDarkEditBrush;
        }
        break;
    }
    
    case WM_ERASEBKGND: {
        HDC hdc = (HDC)w;
        RECT rc;
        GetClientRect(hWnd, &rc);
        if (gDarkMode) {
            FillRect(hdc, &rc, gDarkBgBrush);
        } else {
            if (!gLightBgBrush) gLightBgBrush = CreateSolidBrush(RGB(255, 255, 255));
            FillRect(hdc, &rc, gLightBgBrush);
        }
        return 1;
    }
    
    case WM_CLOSE:
    case WM_DESTROY:
        if (gHelpParent) {
            EnableWindow(gHelpParent, TRUE);
            SetForegroundWindow(gHelpParent);
        }
        gHelpWnd = nullptr;
        gHelpEdit = nullptr;
        gHelpParent = nullptr;
        return 0;
    }
    return DefWindowProcW(hWnd, msg, w, l);
}

static void ShowHelpDialog(HWND parent) {
    if (gHelpWnd) {
        SetForegroundWindow(gHelpWnd);
        return;
    }
    
    gHelpParent = parent;
    EnableWindow(parent, FALSE);
    
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = HelpWndProc;
    wc.hInstance = GetModuleHandleW(nullptr);
    wc.lpszClassName = L"XboxAppArtUpdaterHelpWnd";
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hIcon = LoadIconW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(101));
    wc.hIconSm = LoadIconW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(101));
    RegisterClassExW(&wc);
    
    int w = 590, h = 510;
    RECT parentRc;
    GetWindowRect(parent, &parentRc);
    int x = parentRc.left + (parentRc.right - parentRc.left - w) / 2;
    int y = parentRc.top + (parentRc.bottom - parentRc.top - h) / 2;
    
    gHelpWnd = CreateWindowExW(0,
        wc.lpszClassName,
        L"Help - Xbox App Art Updater",
        WS_CAPTION | WS_SYSMENU | WS_POPUP,
        x, y, w, h,
        parent, nullptr, GetModuleHandleW(nullptr), nullptr);
    
    ShowWindow(gHelpWnd, SW_SHOW);
    UpdateWindow(gHelpWnd);
}

// ------------------------------------------------------------
// Profile Management Window
// ------------------------------------------------------------

static HWND gProfilesWnd = nullptr;
static HWND gProfilesList = nullptr;
static HWND gProfileNameEdit = nullptr;
static HWND gProfileDetailsEdit = nullptr;
static HWND gProfileParent = nullptr;

static const int IDC_PROFILES_LIST = 7001;
static const int IDC_PROFILE_NAME_EDIT = 7002;
static const int IDC_PROFILE_CREATE_BTN = 7003;
static const int IDC_PROFILE_RESTORE_BTN = 7004;
static const int IDC_PROFILE_DELETE_BTN = 7005;
static const int IDC_PROFILE_DETAILS_EDIT = 7006;
static const int IDC_PROFILE_CLOSE_BTN = 7007;

static void RefreshProfilesList() {
    if (!gProfilesList) return;
    
    SendMessageW(gProfilesList, LB_RESETCONTENT, 0, 0);
    
    auto profiles = GetAllProfiles();
    for (const auto& profile : profiles) {
        int idx = (int)SendMessageW(gProfilesList, LB_ADDSTRING, 0, (LPARAM)profile.displayName.c_str());
        SendMessageW(gProfilesList, LB_SETITEMDATA, idx, (LPARAM)new ProfileInfo(profile));
    }
}

static void UpdateProfileDetails() {
    if (!gProfilesList || !gProfileDetailsEdit) return;
    
    int sel = (int)SendMessageW(gProfilesList, LB_GETCURSEL, 0, 0);
    if (sel == LB_ERR) {
        SetWindowTextW(gProfileDetailsEdit, L"Select a profile to view details");
        return;
    }
    
    ProfileInfo* pInfo = (ProfileInfo*)SendMessageW(gProfilesList, LB_GETITEMDATA, sel, 0);
    if (!pInfo) return;
    
    std::wstring details;
    details += L"Profile: " + pInfo->baseName + L"\r\n\r\n";
    
    wchar_t dateStr[100];
    GetDateFormatW(LOCALE_USER_DEFAULT, DATE_LONGDATE, &pInfo->creationTime, nullptr, dateStr, 100);
    wchar_t timeStr[100];
    GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &pInfo->creationTime, nullptr, timeStr, 100);
    details += L"Created: " + std::wstring(dateStr) + L" " + std::wstring(timeStr) + L"\r\n\r\n";
    
    details += L"Art Files: " + std::to_wstring(pInfo->artCount) + L"\r\n\r\n";
    
    // Show breakdown by store
    std::wstring profileFullPath = GetProfilesPath() + L"\\" + pInfo->fullName;
    std::vector<std::wstring> stores = {L"steam", L"epic", L"gog", L"ubi"};
    details += L"Breakdown by Store:\r\n";
    for (const auto& store : stores) {
        fs::path storePath = fs::path(profileFullPath) / store;
        int count = 0;
        if (fs::exists(storePath) && fs::is_directory(storePath)) {
            try {
                for (const auto& entry : fs::directory_iterator(storePath)) {
                    if (entry.is_regular_file()) count++;
                }
            } catch (...) {}
        }
        details += L"  " + store + L": " + std::to_wstring(count) + L" files\r\n";
    }
    
    SetWindowTextW(gProfileDetailsEdit, details.c_str());
}

static LRESULT CALLBACK ProfilesWndProc(HWND hWnd, UINT msg, WPARAM w, LPARAM l) {
    switch (msg) {
    case WM_CREATE: {
        HFONT hFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        
        CreateWindowW(L"STATIC", L"Saved Profiles:",
            WS_CHILD | WS_VISIBLE,
            12, 10, 100, 20,
            hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
        
        gProfilesList = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | LBS_NOTIFY,
            12, 32, 250, 200,
            hWnd, (HMENU)(INT_PTR)IDC_PROFILES_LIST,
            GetModuleHandleW(nullptr), nullptr);
        
        CreateWindowW(L"STATIC", L"Profile Details:",
            WS_CHILD | WS_VISIBLE,
            275, 10, 100, 20,
            hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
        
        gProfileDetailsEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL,
            275, 32, 305, 200,
            hWnd, (HMENU)(INT_PTR)IDC_PROFILE_DETAILS_EDIT,
            GetModuleHandleW(nullptr), nullptr);
        
        CreateWindowW(L"STATIC", L"New Profile Name:",
            WS_CHILD | WS_VISIBLE,
            12, 245, 120, 20,
            hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);
        
        gProfileNameEdit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL,
            12, 267, 180, 24,
            hWnd, (HMENU)(INT_PTR)IDC_PROFILE_NAME_EDIT,
            GetModuleHandleW(nullptr), nullptr);
        
        // Set placeholder text
        SendMessageW(gProfileNameEdit, EM_SETCUEBANNER, TRUE, (LPARAM)L"Enter profile name");
        
        CreateWindowW(L"BUTTON", L"Create Profile",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            200, 265, 120, 28,
            hWnd, (HMENU)(INT_PTR)IDC_PROFILE_CREATE_BTN,
            GetModuleHandleW(nullptr), nullptr);
        
        CreateWindowW(L"BUTTON", L"Restore Selected",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            335, 265, 120, 28,
            hWnd, (HMENU)(INT_PTR)IDC_PROFILE_RESTORE_BTN,
            GetModuleHandleW(nullptr), nullptr);
        
        CreateWindowW(L"BUTTON", L"Delete Selected",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            460, 265, 120, 28,
            hWnd, (HMENU)(INT_PTR)IDC_PROFILE_DELETE_BTN,
            GetModuleHandleW(nullptr), nullptr);
        
        CreateWindowW(L"BUTTON", L"Close",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            500, 308, 80, 28,
            hWnd, (HMENU)(INT_PTR)IDC_PROFILE_CLOSE_BTN,
            GetModuleHandleW(nullptr), nullptr);
        
        for (HWND c = GetWindow(hWnd, GW_CHILD); c != nullptr; c = GetWindow(c, GW_HWNDNEXT)) {
            SendMessageW(c, WM_SETFONT, (WPARAM)hFont, TRUE);
        }
        
        RefreshProfilesList();
        UpdateProfileDetails();
        ApplyDarkModeToWindow(hWnd);
        return 0;
    }
    
    case WM_COMMAND: {
        int id = LOWORD(w);
        int notif = HIWORD(w);
        
        if (id == IDC_PROFILES_LIST && notif == LBN_SELCHANGE) {
            UpdateProfileDetails();
            return 0;
        }
        
        if (id == IDC_PROFILE_CREATE_BTN) {
            wchar_t name[256]{};
            GetWindowTextW(gProfileNameEdit, name, 256);
            std::wstring profileName = name;
            
            if (profileName.empty()) {
                DarkMessageBox(hWnd, L"Please enter a profile name.", L"Profile Name Required", MB_OK | MB_ICONWARNING);
                return 0;
            }
            
            // Validate name (no special characters)
            if (profileName.find_first_of(L"\\/:*?\"<>|") != std::wstring::npos) {
                DarkMessageBox(hWnd, L"Profile name cannot contain: \\ / : * ? \" < > |", L"Invalid Name", MB_OK | MB_ICONWARNING);
                return 0;
            }
            
            int result = DarkMessageBox(hWnd,
                L"This will create a snapshot of your current library folders.\n\nContinue?",
                L"Create Profile", MB_YESNO | MB_ICONQUESTION);
            
            if (result == IDYES) {
                if (CreateProfile(profileName)) {
                    DarkMessageBox(hWnd, L"Profile created successfully!", L"Success", MB_OK | MB_ICONINFORMATION);
                    SetWindowTextW(gProfileNameEdit, L"");
                    RefreshProfilesList();
                } else {
                    DarkMessageBox(hWnd, L"Failed to create profile.", L"Error", MB_OK | MB_ICONERROR);
                }
            }
            return 0;
        }
        
        if (id == IDC_PROFILE_RESTORE_BTN) {
            int sel = (int)SendMessageW(gProfilesList, LB_GETCURSEL, 0, 0);
            if (sel == LB_ERR) {
                DarkMessageBox(hWnd, L"Please select a profile to restore.", L"No Selection", MB_OK | MB_ICONWARNING);
                return 0;
            }
            
            ProfileInfo* pInfo = (ProfileInfo*)SendMessageW(gProfilesList, LB_GETITEMDATA, sel, 0);
            if (!pInfo) return 0;
            
            int result = DarkMessageBox(hWnd,
                (L"This will restore library folders from:\n" + pInfo->displayName + L"\n\nYour current libraries will be replaced!\n\nContinue?").c_str(),
                L"Restore Profile", MB_YESNO | MB_ICONWARNING);
            
            if (result == IDYES) {
                if (RestoreProfile(*pInfo)) {
                    DarkMessageBox(hWnd, L"Profile restored successfully!\n\nPlease restart the Xbox app to see changes.", L"Success", MB_OK | MB_ICONINFORMATION);
                    
                    // Trigger rescan
                    if (gMainWnd) {
                        PostMessageW(gMainWnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_SCAN, 0), 0);
                    }
                } else {
                    DarkMessageBox(hWnd, L"Failed to restore profile.", L"Error", MB_OK | MB_ICONERROR);
                }
            }
            return 0;
        }
        
        if (id == IDC_PROFILE_DELETE_BTN) {
            int sel = (int)SendMessageW(gProfilesList, LB_GETCURSEL, 0, 0);
            if (sel == LB_ERR) {
                DarkMessageBox(hWnd, L"Please select a profile to delete.", L"No Selection", MB_OK | MB_ICONWARNING);
                return 0;
            }
            
            ProfileInfo* pInfo = (ProfileInfo*)SendMessageW(gProfilesList, LB_GETITEMDATA, sel, 0);
            if (!pInfo) return 0;
            
            int result = DarkMessageBox(hWnd,
                (L"Delete profile:\n" + pInfo->displayName + L"\n\nThis cannot be undone!").c_str(),
                L"Delete Profile", MB_YESNO | MB_ICONWARNING);
            
            if (result == IDYES) {
                if (DeleteProfile(*pInfo)) {
                    DarkMessageBox(hWnd, L"Profile deleted successfully!", L"Success", MB_OK | MB_ICONINFORMATION);
                    RefreshProfilesList();
                    UpdateProfileDetails();
                } else {
                    DarkMessageBox(hWnd, L"Failed to delete profile.", L"Error", MB_OK | MB_ICONERROR);
                }
            }
            return 0;
        }
        
        if (id == IDC_PROFILE_CLOSE_BTN) {
            PostMessageW(hWnd, WM_CLOSE, 0, 0);
            return 0;
        }
        return 0;
    }
    
    case WM_DRAWITEM: {
        DRAWITEMSTRUCT* dis = (DRAWITEMSTRUCT*)l;
        if (dis->CtlType == ODT_BUTTON) {
            DrawDarkButton(dis);
            return TRUE;
        }
        break;
    }
    
    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)w;
        if (gDarkMode) {
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkColor(hdc, RGB(32, 32, 32));
            return (LRESULT)gDarkBgBrush;
        } else {
            SetTextColor(hdc, RGB(0, 0, 0));
            SetBkColor(hdc, RGB(255, 255, 255));
            return (LRESULT)gLightBgBrush;
        }
    }
    
    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORLISTBOX: {
        HDC hdc = (HDC)w;
        if (gDarkMode) {
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkColor(hdc, RGB(45, 45, 45));
            return (LRESULT)gDarkEditBrush;
        }
        break;
    }
    
    case WM_ERASEBKGND: {
        HDC hdc = (HDC)w;
        RECT rc;
        GetClientRect(hWnd, &rc);
        if (gDarkMode) {
            FillRect(hdc, &rc, gDarkBgBrush);
        } else {
            if (!gLightBgBrush) gLightBgBrush = CreateSolidBrush(RGB(255, 255, 255));
            FillRect(hdc, &rc, gLightBgBrush);
        }
        return 1;
    }
    
    case WM_CLOSE: {
        // Free allocated ProfileInfo objects
        if (gProfilesList) {
            int count = (int)SendMessageW(gProfilesList, LB_GETCOUNT, 0, 0);
            for (int i = 0; i < count; i++) {
                ProfileInfo* pInfo = (ProfileInfo*)SendMessageW(gProfilesList, LB_GETITEMDATA, i, 0);
                delete pInfo;
            }
        }
        
        if (gProfileParent) {
            EnableWindow(gProfileParent, TRUE);
            SetForegroundWindow(gProfileParent);
        }
        gProfilesWnd = nullptr;
        gProfilesList = nullptr;
        gProfileNameEdit = nullptr;
        gProfileDetailsEdit = nullptr;
        gProfileParent = nullptr;
        DestroyWindow(hWnd);
        return 0;
    }
    }
    return DefWindowProcW(hWnd, msg, w, l);
}

static void ShowProfilesDialog(HWND parent) {
    if (gProfilesWnd) {
        SetForegroundWindow(gProfilesWnd);
        return;
    }
    
    gProfileParent = parent;
    EnableWindow(parent, FALSE);
    
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = ProfilesWndProc;
    wc.hInstance = GetModuleHandleW(nullptr);
    wc.lpszClassName = L"XboxAppArtUpdaterProfilesWnd";
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hIcon = LoadIconW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(101));
    wc.hIconSm = LoadIconW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(101));
    RegisterClassExW(&wc);
    
    RECT pr{};
    GetWindowRect(parent, &pr);
    int w = 600, h = 380;
    int x = pr.left + ((pr.right - pr.left) - w) / 2;
    int y = pr.top + ((pr.bottom - pr.top) - h) / 2;
    
    gProfilesWnd = CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        wc.lpszClassName,
        L"Profile Management",
        WS_CAPTION | WS_SYSMENU | WS_POPUP,
        x, y, w, h,
        parent, nullptr, GetModuleHandleW(nullptr), nullptr);
    
    ShowWindow(gProfilesWnd, SW_SHOW);
    UpdateWindow(gProfilesWnd);
}

// ------------------------------------------------------------
// Window
// ------------------------------------------------------------

static LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM w, LPARAM l) {
    switch (msg) {
    case WM_CREATE: {
        InitCommonControls();

        CreateWindowW(L"STATIC", L"Icon size:",
            WS_CHILD | WS_VISIBLE,
            10, 14, 70, 20,
            hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);

        gCmbSize = CreateWindowW(WC_COMBOBOXW, L"",
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            85, 10, 90, 200,
            hWnd, (HMENU)(INT_PTR)IDC_CMB_SIZE,
            GetModuleHandleW(nullptr), nullptr);

        CreateWindowW(L"STATIC", L"Layout:",
            WS_CHILD | WS_VISIBLE,
            190, 14, 50, 20,
            hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);

        gCmbLayout = CreateWindowW(WC_COMBOBOXW, L"",
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            245, 10, 120, 200,
            hWnd, (HMENU)(INT_PTR)IDC_CMB_LAYOUT,
            GetModuleHandleW(nullptr), nullptr);

        // Filter controls
        CreateWindowW(L"STATIC", L"Store:",
            WS_CHILD | WS_VISIBLE,
            380, 14, 40, 20,
            hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);

        gCmbStore = CreateWindowW(WC_COMBOBOXW, L"",
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            425, 10, 90, 200,
            hWnd, (HMENU)(INT_PTR)IDC_CMB_STORE,
            GetModuleHandleW(nullptr), nullptr);
        SendMessageW(gCmbStore, CB_ADDSTRING, 0, (LPARAM)L"All");
        SendMessageW(gCmbStore, CB_ADDSTRING, 0, (LPARAM)L"Steam");
        SendMessageW(gCmbStore, CB_ADDSTRING, 0, (LPARAM)L"Epic");
        SendMessageW(gCmbStore, CB_ADDSTRING, 0, (LPARAM)L"GOG");
        SendMessageW(gCmbStore, CB_ADDSTRING, 0, (LPARAM)L"Ubisoft");
        SendMessageW(gCmbStore, CB_SETCURSEL, 0, 0);

        CreateWindowW(L"STATIC", L"Art:",
            WS_CHILD | WS_VISIBLE,
            525, 14, 30, 20,
            hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);

        gCmbArtFilter = CreateWindowW(WC_COMBOBOXW, L"",
            WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS,
            560, 10, 100, 200,
            hWnd, (HMENU)(INT_PTR)IDC_CMB_ART_FILTER,
            GetModuleHandleW(nullptr), nullptr);
        SendMessageW(gCmbArtFilter, CB_ADDSTRING, 0, (LPARAM)L"All");
        SendMessageW(gCmbArtFilter, CB_ADDSTRING, 0, (LPARAM)L"Has Art");
        SendMessageW(gCmbArtFilter, CB_ADDSTRING, 0, (LPARAM)L"Missing Art");
        SendMessageW(gCmbArtFilter, CB_SETCURSEL, 0, 0);

        CreateWindowW(L"STATIC", L"Search:",
            WS_CHILD | WS_VISIBLE,
            675, 14, 48, 20,
            hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);

        gSearchEdit = CreateWindowW(L"EDIT", L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
            728, 10, 150, 24,
            hWnd, (HMENU)(INT_PTR)IDC_SEARCH_EDIT,
            GetModuleHandleW(nullptr), nullptr);
        SetEditPlaceholder(gSearchEdit, L"Search games...");

        // Right-aligned buttons (will be repositioned in WM_SIZE)
        gBtnHelp = CreateWindowW(L"BUTTON", L"Help",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            0, 10, 60, 30,
            hWnd, (HMENU)(INT_PTR)IDC_BTN_HELP,
            GetModuleHandleW(nullptr), nullptr);

        gBtnConfig = CreateWindowW(L"BUTTON", L"Config",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            0, 10, 80, 30,
            hWnd, (HMENU)(INT_PTR)IDC_BTN_CONFIG,
            GetModuleHandleW(nullptr), nullptr);

        gBtnScan = CreateWindowW(L"BUTTON", L"Scan",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            0, 10, 80, 30,
            hWnd, (HMENU)(INT_PTR)IDC_BTN_SCAN,
            GetModuleHandleW(nullptr), nullptr);

        gBtnQuickApply = CreateWindowW(L"BUTTON", L"Quick Apply",
            WS_CHILD | BS_OWNERDRAW,  // Hidden initially, shown after scan
            0, 10, 90, 30,
            hWnd, (HMENU)(INT_PTR)IDC_BTN_QUICK_APPLY,
            GetModuleHandleW(nullptr), nullptr);

        // Create instruction note above the list
        gListInstructionNote = CreateWindowW(L"STATIC", 
            L"Tip: Double-click a game to customize it. Use CTRL to select multiple games for Quick Apply.",
            WS_CHILD | WS_VISIBLE | SS_CENTER,
            10, 90, 800, 20,
            hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);


        // Create unresolved checkbox above Select All, under icon size
        gChkUnresolved = CreateWindowW(L"BUTTON", L"Show unresolved titles",
            WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX,
            10, 35, 200, 20,
            hWnd, (HMENU)11001, GetModuleHandleW(nullptr), nullptr);
        SendMessageW(gChkUnresolved, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);

        gBtnSelectAll = CreateWindowW(L"BUTTON", L"Select All",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            10, 80, 80, 30,
            hWnd, (HMENU)(INT_PTR)IDC_BTN_SELECT_ALL,
            GetModuleHandleW(nullptr), nullptr);

        gBtnProfiles = CreateWindowW(L"BUTTON", L"Profiles",
            WS_CHILD | WS_VISIBLE | BS_OWNERDRAW,
            0, 10, 75, 30,
            hWnd, (HMENU)(INT_PTR)IDC_BTN_PROFILES,
            GetModuleHandleW(nullptr), nullptr);

        // Store main window handle
        gMainWnd = hWnd;

        // Create tooltip control
        gTooltip = CreateWindowExW(0, TOOLTIPS_CLASSW, nullptr,
            WS_POPUP | TTS_ALWAYSTIP | TTS_NOPREFIX,
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
            hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);

        // Add tooltips to buttons


        // Status bar at the bottom
        gStatusBar = CreateWindowW(STATUSCLASSNAMEW, L"Ready",
            WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
            0, 0, 0, 0,
            hWnd, (HMENU)(INT_PTR)IDC_STATUS_BAR,
            GetModuleHandleW(nullptr), nullptr);
        
        // Subclass status bar for dark mode
        if (gStatusBar) {
            gOriginalStatusBarProc = (WNDPROC)SetWindowLongPtrW(gStatusBar, GWLP_WNDPROC, (LONG_PTR)StatusBarSubclassProc);
        }

        InitCombosAndDefaults();

        RecreateListView(hWnd);
        
        // Create the "no scan" message overlay (positioned in WM_SIZE)
        std::wstring defaultPath = ReplaceUsernameInPath(GetDefaultThirdPartyLibrariesPath());
        std::wstring noScanMsg = 
            L"Welcome to Xbox App Art Updater!\r\n\r\n"
            L"Please click the SCAN button to start the automatic game detection.\r\n\r\n"
            L"If the scan doesn't find your games, please configure the ThirdPartyLibraries path in Config.\r\n\r\n"
            L"The default path is:\r\n" + defaultPath;
        
        gNoScanMessage = CreateWindowW(L"EDIT", noScanMsg.c_str(),
            WS_CHILD | ES_MULTILINE | ES_READONLY | ES_CENTER,
            50, 100, 600, 200,
            hWnd, nullptr, GetModuleHandleW(nullptr), nullptr);

        CacheState st;
        if (LoadCacheFile(st)) {
            // Apply config first (affects scan paths)
            gConfig = st.config;
            gDarkMode = gConfig.darkMode;

            ApplyCombosFromState(st);
            RecreateListView(hWnd);

            gLastItems = st.items;
            PopulateListFromItems(gLastItems);
            
            // Mark as already scanned and update button to "Refresh"
            if (!gLastItems.empty()) {
                gHasScanned = true;
                if (gBtnScan) {
                    SetWindowTextW(gBtnScan, L"Refresh");
                }
            }
            
            // Update status bar
            std::wstring statusText = std::to_wstring(gLastItems.size()) + L" games loaded from cache";
            SendMessageW(gStatusBar, SB_SETTEXTW, 0, (LPARAM)statusText.c_str());
        } else {
            gConfig = AppConfig{};
            gLastItems.clear();
            PopulateListFromItems(gLastItems);
            SendMessageW(gStatusBar, SB_SETTEXTW, 0, (LPARAM)L"Click Scan to find games");
        }
        
        // Apply dark mode to window
        ApplyDarkModeToWindow(hWnd);
        ApplyDarkModeToControls(hWnd);
        
        return 0;
    }

    case WM_COMMAND: {
        int code = HIWORD(w);
        int id = LOWORD(w);
        HWND src = (HWND)l;

        if (id == 11001 && src == gChkUnresolved) {
            OnUnresolvedCheckChanged();
            return 0;
        }

        if (code == CBN_SELCHANGE) {
            int ctrlId = id;
            if (ctrlId == IDC_CMB_STORE) {
                if (gQuickApplyRunning) return 0; // Ignore during Quick Apply
                int sel = (int)SendMessageW(gCmbStore, CB_GETCURSEL, 0, 0);
                switch (sel) {
                    case 0: gFilterStore = L""; break;
                    case 1: gFilterStore = L"steam"; break;
                    case 2: gFilterStore = L"epic"; break;
                    case 3: gFilterStore = L"gog"; break;
                    case 4: gFilterStore = L"ubi"; break;
                }
                PopulateListFromItems(gLastItems);
                return 0;
            }
            if (ctrlId == IDC_CMB_ART_FILTER) {
                if (gQuickApplyRunning) return 0; // Ignore during Quick Apply
                gFilterArt = (int)SendMessageW(gCmbArtFilter, CB_GETCURSEL, 0, 0);
                PopulateListFromItems(gLastItems);
                return 0;
            }
            OnComboChanged(hWnd, ctrlId);
            return 0;
        }
        if (code == EN_CHANGE && id == IDC_SEARCH_EDIT) {
            if (gQuickApplyRunning) return 0; // Ignore during Quick Apply
            wchar_t buf[256]{};
            GetWindowTextW(gSearchEdit, buf, _countof(buf));
            gFilterSearch = buf;
            PopulateListFromItems(gLastItems);
            return 0;
        }
        if (id == IDC_BTN_CONFIG) {
            OpenConfigWindow(hWnd);
            return 0;
        }
        if (id == IDC_BTN_SCAN) {
            DoScanAndCache(hWnd);
            return 0;
        }
        if (id == IDC_BTN_HELP) {
            ShowHelpDialog(hWnd);
            return 0;
        }
        if (id == IDC_BTN_QUICK_APPLY) {
            DoQuickApply(hWnd);
            return 0;
        }
        if (id == IDC_BTN_SELECT_ALL) {
            // Select all items in the list
            int count = ListView_GetItemCount(gList);
            for (int i = 0; i < count; i++) {
                ListView_SetItemState(gList, i, LVIS_SELECTED, LVIS_SELECTED);
            }
            return 0;
        }
        if (id == IDC_BTN_PROFILES) {
            ShowProfilesDialog(hWnd);
            return 0;
        }
        return 0;
    }

    case WM_NOTIFY: {
        NMHDR* nmhdr = (NMHDR*)l;
        
        // Custom draw for ListView header
        if (nmhdr->code == NM_CUSTOMDRAW && gDarkMode) {
            NMCUSTOMDRAW* nmcd = (NMCUSTOMDRAW*)l;
            
            // Check if this is from the header control
            HWND hHeader = ListView_GetHeader(gList);
            if (nmhdr->hwndFrom == hHeader) {
                switch (nmcd->dwDrawStage) {
                case CDDS_PREPAINT:
                    return CDRF_NOTIFYITEMDRAW;
                case CDDS_ITEMPREPAINT: {
                    // Draw dark header item
                    HDC hdc = nmcd->hdc;
                    RECT rc = nmcd->rc;
                    
                    // Dark background
                    HBRUSH bgBrush = CreateSolidBrush(RGB(45, 45, 45));
                    FillRect(hdc, &rc, bgBrush);
                    DeleteObject(bgBrush);
                    
                    // Border
                    HPEN borderPen = CreatePen(PS_SOLID, 1, RGB(70, 70, 70));
                    HPEN oldPen = (HPEN)SelectObject(hdc, borderPen);
                    MoveToEx(hdc, rc.right - 1, rc.top, nullptr);
                    LineTo(hdc, rc.right - 1, rc.bottom);
                    MoveToEx(hdc, rc.left, rc.bottom - 1, nullptr);
                    LineTo(hdc, rc.right, rc.bottom - 1);
                    SelectObject(hdc, oldPen);
                    DeleteObject(borderPen);
                    
                    // Get header text
                    wchar_t text[256]{};
                    HDITEMW hdi{};
                    hdi.mask = HDI_TEXT;
                    hdi.pszText = text;
                    hdi.cchTextMax = _countof(text);
                    Header_GetItem(hHeader, nmcd->dwItemSpec, &hdi);
                    
                    // Draw bold white text
                    SetBkMode(hdc, TRANSPARENT);
                    SetTextColor(hdc, RGB(255, 255, 255));
                    
                    HFONT hFont = (HFONT)SendMessageW(hHeader, WM_GETFONT, 0, 0);
                    LOGFONTW lf{};
                    GetObjectW(hFont, sizeof(lf), &lf);
                    lf.lfWeight = FW_BOLD;
                    HFONT boldFont = CreateFontIndirectW(&lf);
                    HFONT oldFont = (HFONT)SelectObject(hdc, boldFont);
                    
                    RECT textRc = rc;
                    textRc.left += 6;
                    textRc.right -= 6;
                    DrawTextW(hdc, text, -1, &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
                    
                    SelectObject(hdc, oldFont);
                    DeleteObject(boldFont);
                    
                    return CDRF_SKIPDEFAULT;
                }
                }
            }
        }
        
        if (nmhdr->hwndFrom == gList && nmhdr->code == NM_DBLCLK) {
            NMITEMACTIVATE* nmia = (NMITEMACTIVATE*)l;
            int idx = nmia->iItem;
            if (idx >= 0 && idx < (int)gFilteredItems.size()) {
                OpenArtWindow(hWnd, gFilteredItems[idx]);
            }
            return 0;
        }
        return 0;
    }

    case WM_SIZE: {
        int wdt = LOWORD(l);
        int hgt = HIWORD(l);
        
        // Position buttons to the right side
        if (gBtnHelp) {
            MoveWindow(gBtnHelp, wdt - 350, 10, 60, 30, TRUE);
        }
        if (gBtnScan) {
            MoveWindow(gBtnScan, wdt - 280, 10, 80, 30, TRUE);
        }
        if (gBtnConfig) {
            MoveWindow(gBtnConfig, wdt - 190, 10, 80, 30, TRUE);
        }
        if (gBtnProfiles) {
            MoveWindow(gBtnProfiles, wdt - 100, 10, 75, 30, TRUE);
        }
        
        // Position Select All and Quick Apply buttons on the left below filters
        if (gBtnSelectAll) {
            MoveWindow(gBtnSelectAll, 10, 80, 80, 30, TRUE);
        }
        if (gChkUnresolved) {
            MoveWindow(gChkUnresolved, 10, 35, 200, 20, TRUE);
        }
        if (gBtnQuickApply) {
            MoveWindow(gBtnQuickApply, wdt - 100, 50, 90, 30, TRUE);
        }
        
        // Resize status bar first (it auto-sizes)
        if (gStatusBar) {
            SendMessageW(gStatusBar, WM_SIZE, 0, 0);
            
            // Get status bar height
            RECT sbRect;
            GetWindowRect(gStatusBar, &sbRect);
            int sbHeight = sbRect.bottom - sbRect.top;
            
            // Resize instruction note
            if (gListInstructionNote) {
                MoveWindow(gListInstructionNote, 10, 90, wdt - 20, 20, TRUE);
            }
            
            if (gList) {
                MoveWindow(gList, 10, 115, wdt - 20, hgt - 125 - sbHeight, TRUE);
            }
            
            // Center the no-scan message in the list area
            if (gNoScanMessage) {
                int msgW = 500;
                int msgH = 180;
                int listH = hgt - 125 - sbHeight;
                int msgX = 10 + (wdt - 20 - msgW) / 2;
                int msgY = 115 + (listH - msgH) / 2;
                MoveWindow(gNoScanMessage, msgX, msgY, msgW, msgH, TRUE);
            }
        } else if (gList) {
            MoveWindow(gList, 10, 115, wdt - 20, hgt - 125, TRUE);
        }
        return 0;
    }

    case WM_DESTROY:
        DestroyImageList();
        if (gDarkBgBrush) { DeleteObject(gDarkBgBrush); gDarkBgBrush = nullptr; }
        if (gDarkEditBrush) { DeleteObject(gDarkEditBrush); gDarkEditBrush = nullptr; }
        PostQuitMessage(0);
        return 0;
    
    case WM_ERASEBKGND: {
        if (gDarkMode) {
            HDC hdc = (HDC)w;
            RECT rc;
            GetClientRect(hWnd, &rc);
            FillRect(hdc, &rc, gDarkBgBrush);
            return 1;
        }
        break;
    }
    
    case WM_DRAWITEM: {
        DRAWITEMSTRUCT* dis = (DRAWITEMSTRUCT*)l;
        if (dis->CtlType == ODT_BUTTON) {
            DrawDarkButton(dis);
            return TRUE;
        }
        if (dis->CtlType == ODT_COMBOBOX) {
            DrawDarkComboItem(dis);
            return TRUE;
        }
        break;
    }
    
    case WM_CTLCOLORLISTBOX: {
        // For combobox dropdown list
        HDC hdc = (HDC)w;
        if (gDarkMode) {
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkColor(hdc, RGB(45, 45, 45));
            if (!gDarkEditBrush) {
                gDarkEditBrush = CreateSolidBrush(RGB(45, 45, 45));
            }
            return (LRESULT)gDarkEditBrush;
        } else {
            SetTextColor(hdc, RGB(0, 0, 0));
            SetBkColor(hdc, RGB(255, 255, 255));
            if (!gLightBgBrush) gLightBgBrush = CreateSolidBrush(RGB(255, 255, 255));
            return (LRESULT)gLightBgBrush;
        }
    }
    
    case WM_CTLCOLORSTATIC: {
        HDC hdc = (HDC)w;
        if (gDarkMode) {
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkColor(hdc, RGB(32, 32, 32));
            return (LRESULT)gDarkBgBrush;
        } else {
            SetTextColor(hdc, RGB(0, 0, 0));
            SetBkColor(hdc, RGB(255, 255, 255));
            if (!gLightBgBrush) gLightBgBrush = CreateSolidBrush(RGB(255, 255, 255));
            return (LRESULT)gLightBgBrush;
        }
    }
    
    case WM_CTLCOLOREDIT: {
        HDC hdc = (HDC)w;
        if (gDarkMode) {
            SetTextColor(hdc, RGB(255, 255, 255));
            SetBkColor(hdc, RGB(45, 45, 45));
            return (LRESULT)gDarkEditBrush;
        }
        break;
    }
    }
    return DefWindowProcW(hWnd, msg, w, l);
}

// ------------------------------------------------------------
// Entry point
// ------------------------------------------------------------

int WINAPI wWinMain(HINSTANCE h, HINSTANCE, PWSTR, int nCmd) {
    GdiplusStartupInput gdiplusStartupInput;
    if (GdiplusStartup(&gGdiplusToken, &gdiplusStartupInput, nullptr) != Ok) {
        MessageBoxW(nullptr, L"Failed to initialize GDI+.", APP_TITLE, MB_OK | MB_ICONERROR);
        return 1;
    }

    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    
    // Initialize dark mode APIs and set app preference before creating windows
    LoadDarkModeApis();
    if (SetPreferredAppMode) {
        SetPreferredAppMode(AllowDark);
    }
    if (RefreshImmersiveColorPolicyState) {
        RefreshImmersiveColorPolicyState();
    }

    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = h;
    wc.lpszClassName = L"XboxAppArtUpdaterWnd";
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hIcon = LoadIconW(h, MAKEINTRESOURCEW(101));
    wc.hIconSm = LoadIconW(h, MAKEINTRESOURCEW(101));
    RegisterClassExW(&wc);

    HWND wnd = CreateWindowW(
        wc.lpszClassName, APP_TITLE,
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT,
        1300, 820,
        nullptr, nullptr, h, nullptr);

    ShowWindow(wnd, nCmd);

    MSG msg;
while (GetMessageW(&msg, nullptr, 0, 0)) {
    TranslateMessage(&msg);
    DispatchMessageW(&msg);
}
    if (gGdiplusToken) GdiplusShutdown(gGdiplusToken);
    return 0;
}

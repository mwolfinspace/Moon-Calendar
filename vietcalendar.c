#define WIN32_LEAN_AND_MEAN
#ifndef UNICODE
#define UNICODE
#endif
#ifndef _UNICODE
#define _UNICODE
#endif
#define _WIN32_IE 0x0500
#define _WIN32_WINNT 0x0600

#define COBJMACROS
#include <windows.h>
#include <commdlg.h>  // Đã thêm thư viện này để sửa lỗi CHOOSEFONTW/CHOOSECOLOR
#include <shellapi.h>
#include <shlobj.h>
#include <wininet.h>
#include <dxgiformat.h>
#include <d2d1.h>
#include <dwrite.h>
#include <dwrite_2.h>
#include <stdio.h>
#include <stdlib.h>
#include <math.h>
#include <wctype.h>

#pragma comment(lib, "wininet.lib")
#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "dwrite.lib")

#ifdef __MINGW32__
// MinGW headers sometimes don't provide a linkable IID_IDWriteFactory symbol.
// Define it here to satisfy the linker when using DWriteCreateFactory.
const IID IID_IDWriteFactory = {0xb859ee5a, 0xd838, 0x4b5b, {0xa2, 0xe8, 0x1a, 0xdc, 0x7d, 0x93, 0xdb, 0x48}};
const IID IID_IDWriteFactory2 = {0x0439fc60, 0xca44, 0x4994, {0x8d, 0xee, 0x3a, 0x9a, 0xf7, 0xb7, 0x32, 0xec}};
const IID IID_IDWritePixelSnapping = {0xeaf3a2da, 0xecf4, 0x4d24, {0xb6, 0x44, 0xb3, 0x4f, 0x68, 0x42, 0x02, 0x4b}};
const IID IID_IDWriteTextRenderer = {0xef8a8135, 0x5cc6, 0x45fe, {0x88, 0x25, 0xc5, 0xa0, 0x72, 0x4e, 0xb8, 0x19}};
#endif

#ifndef WM_DPICHANGED
#define WM_DPICHANGED 0x02E0
#endif

#define APP_NAME L"VietCalendar"
#define WM_TRAYICON (WM_USER + 1)
#define WM_ONLINE_UPDATE (WM_APP + 1)
#define ID_EXIT 1001
#define ID_SETTINGS 1002
#define ID_TOGGLE_AUTORUN 1003
#define ID_RESET_POS 1004
#define ID_PREV_MONTH 1010
#define ID_NEXT_MONTH 1011
#define ID_GOTO_TODAY 1012
#define ID_TOGGLE_MOVE 1013
#define ID_OPACITY_TRACK 2001
#define ID_BTN_COLOR 2002
#define ID_BTN_FONT 2003
#define ID_EDIT_X 2004
#define ID_EDIT_Y 2005
#define ID_CHK_AUTORUN 2006
#define ID_CHK_STICK 2007
#define ID_CHK_PRIMARY 2008
#define ID_CHK_MOVE 2011
#define ID_APPLY 2009
#define ID_EMOJI_STYLE 2012
#define ID_BTN_EMOJI_FONT 2013
#define ID_LBL_EMOJI_FONT 2014
#define ID_SLIDER_SOLAR 2015
#define ID_SLIDER_EMOJI 2016
#define ID_SLIDER_LUNAR 2017
#define ID_EDIT_SOLAR 2018
#define ID_EDIT_EMOJI 2019
#define ID_EDIT_LUNAR 2020
#define ID_LBL_PCT_SOLAR 2021
#define ID_LBL_PCT_EMOJI 2022
#define ID_LBL_PCT_LUNAR 2023
#define ID_LBL_TITLE 2100
#define ID_LBL_POSX 2101
#define ID_LBL_POSY 2102
#define ID_LBL_OPACITY 2103
#define ID_LBL_FONTINFO 2010
#define ID_LBL_EMOJI_STYLE 2104
#define ID_LBL_SOLAR 2110
#define ID_LBL_EMOJI 2111
#define ID_LBL_LUNAR 2112

// --- CẤU TRÚC DỮ LIỆU & BIẾN TOÀN CỤC ---
typedef struct {
    int x, y;
    int width, height;
    int alpha;
    COLORREF bgColor;
    COLORREF textColor;
    COLORREF highlightColor;
    COLORREF lunarColor;
    BOOL autoRun;
    BOOL stickDesktop;
    BOOL primaryScreenOnly;
    BOOL allowMove;
    wchar_t fontName[32];
    int fontSize;
    wchar_t emojiFontColor[64];
    wchar_t emojiFontMono[64];
    BOOL emojiMono;
    int solarScalePct;
    int emojiScalePct;
    int lunarScalePct;
    float relX, relY;
    wchar_t monitorId[64];
} Config;

static Config g_cfg = {-1, -1, 0, 0, 245, RGB(248,249,252), RGB(24,32,45), RGB(14,120,200), RGB(90,103,120), FALSE, TRUE, FALSE, TRUE, L"Segoe UI Variable", 11, L"Segoe UI Emoji", L"Segoe UI Symbol", FALSE, 125, 100, 85, -1.0f, -1.0f, L""};
static HWND g_hwnd = NULL, g_settingsHwnd = NULL, g_hTooltip = NULL;
static NOTIFYICONDATAW g_nid = {0};
static HFONT g_fontBig = NULL, g_fontSmall = NULL, g_fontMonth = NULL, g_fontLunar = NULL, g_fontEmoji = NULL;
static HICON g_hIcon = NULL;
static int g_hoverDay = -1;
static float g_dpiScale = 1.0f;
static int g_cellWidth = 0, g_cellHeight = 0, g_headerHeight = 0;
static int g_dayHeaderHeight = 0;
static int g_shadowSize = 0;
static int g_cardWidth = 0, g_cardHeight = 0;
static int g_sideMargin = 0, g_topMargin = 0, g_bottomMargin = 0;
static RECT g_cardRect = {0};
static RECT g_rcPrev = {0}, g_rcNext = {0}, g_rcHeader = {0};
static BOOL g_trackingMouse = FALSE;
static BOOL g_hoverHeader = FALSE;
static float g_headerPulse = 0.0f;
static BOOL g_mouseInside = FALSE;
static BOOL g_stickRaised = FALSE;
static int g_viewMonth = 0, g_viewYear = 0;
static BOOL g_followToday = TRUE;
static BOOL g_dragging = FALSE;
static POINT g_dragOffset = {0};
static int g_visibleRows = 6;
static float g_fontScale = 1.0f;
static float g_effectiveFontSize = 11.0f;
static float g_solarScale = 1.25f;
static float g_lunarScale = 0.85f;
static float g_emojiScale = 1.0f;
static BOOL g_showLunar = TRUE;
static int g_hoverNav = -1;
static HFONT g_settingsFont = NULL;
static UINT g_settingsFontDpi = 0;
static BOOL g_syncingScaleEdits = FALSE;
static int g_lastTodayYmd = 0;
static ID2D1Factory* g_d2dFactory = NULL;
static IDWriteFactory* g_dwriteFactory = NULL;
static IDWriteFactory2* g_dwriteFactory2 = NULL;
static ID2D1DCRenderTarget* g_d2dTarget = NULL;
static ID2D1SolidColorBrush* g_d2dTextBrush = NULL;
static IDWriteTextFormat* g_dwriteEmojiFormatColor = NULL;
static IDWriteTextFormat* g_dwriteEmojiFormatMono = NULL;

typedef struct {
    int day, month, year;
    int leap; // 1 nếu là tháng nhuận
} LunarDate;

static CRITICAL_SECTION g_onlineLock;
static BOOL g_onlineLockInit = FALSE;
static BOOL g_onlineHasData = FALSE;
static SYSTEMTIME g_onlineSolar = {0};
static LunarDate g_onlineLunar = {0};
static LONG g_onlineCheckInProgress = 0;
static int g_onlineLastSuccessYMD = 0;
static DWORD g_onlineNextAttemptTick = 0;

/* ============================================================================
 * ENGINE TÍNH TOÁN ÂM LỊCH CHÍNH XÁC CAO (HỒ NGỌC ĐỨC / JEAN MEEUS)
 * Được viết lại để xử lý chính xác múi giờ UTC+7
 * ============================================================================ */

#define PI 3.14159265358979323846

static int IntFloor(double d) { return (int)floor(d); }

// Tính Julian Day Number từ ngày Dương lịch (theo chuẩn Ho Ngọc Đức)
static int jdFromDate(int d, int m, int y) {
    int a = (14 - m) / 12;
    int yy = y + 4800 - a;
    int mm = m + 12 * a - 3;
    int jd = d + IntFloor((153 * mm + 2) / 5.0) + 365 * yy + IntFloor(yy / 4.0)
        - IntFloor(yy / 100.0) + IntFloor(yy / 400.0) - 32045;
    if (jd < 2299161) {
        jd = d + IntFloor((153 * mm + 2) / 5.0) + 365 * yy + IntFloor(yy / 4.0) - 32083;
    }
    return jd;
}

static const wchar_t* kCan[] = {L"Gi\u00E1p", L"\u1EA4t", L"B\u00EDnh", L"\u0110inh", L"M\u1EADu", L"K\u1EF7", L"Canh", L"T\u00E2n", L"Nh\u00E2m", L"Qu\u00FD"};
static const wchar_t* kChi[] = {L"T\u00FD", L"S\u1EEDu", L"D\u1EA7n", L"M\u00E3o", L"Th\u00ECn", L"T\u1EF5", L"Ng\u1ECD", L"M\u00F9i", L"Th\u00E2n", L"D\u1EADu", L"Tu\u1EA5t", L"H\u1EE3i"};
static const wchar_t* kDayNames[] = {L"CN", L"T2", L"T3", L"T4", L"T5", L"T6", L"T7"};
static const wchar_t* kChiEmoji[] = {
    L"\U0001F400", // Rat
    L"\U0001F402", // Ox
    L"\U0001F405", // Tiger
    L"\U0001F408", // Cat (Vietnamese Zodiac)
    L"\U0001F409", // Dragon
    L"\U0001F40D", // Snake
    L"\U0001F40E", // Horse
    L"\U0001F410", // Goat
    L"\U0001F412", // Monkey
    L"\U0001F413", // Rooster
    L"\U0001F415", // Dog
    L"\U0001F416"  // Pig
};

static void GetCanChiDay(int d, int m, int y, wchar_t* out, int outLen) {
    int jdn = jdFromDate(d, m, y);
    swprintf(out, outLen, L"%s %s", kCan[(jdn + 9) % 10], kChi[(jdn + 1) % 12]);
}

static int GetChiIndexDay(int d, int m, int y) {
    int jdn = jdFromDate(d, m, y);
    return (jdn + 1) % 12;
}

static void GetCanChiMonth(int lunarMonth, int lunarYear, wchar_t* out, int outLen) {
    int can = (lunarYear * 12 + lunarMonth + 3) % 10;
    int chi = (lunarMonth + 1) % 12;
    swprintf(out, outLen, L"%s %s", kCan[can], kChi[chi]);
}

static void GetCanChiYear(int lunarYear, wchar_t* out, int outLen) {
    int can = (lunarYear + 6) % 10;
    int chi = (lunarYear + 8) % 12;
    swprintf(out, outLen, L"%s %s", kCan[can], kChi[chi]);
}

static int getNewMoonDay(int k, int timeZone) {
    double T = k / 1236.85;
    double T2 = T * T;
    double T3 = T2 * T;
    double dr = PI / 180.0;
    double Jd1 = 2415020.75933 + 29.53058868 * k + 0.0001178 * T2 - 0.000000155 * T3;
    Jd1 += 0.00033 * sin((166.56 + 132.87 * T - 0.009173 * T2) * dr);
    double M = 359.2242 + 29.10535608 * k - 0.0000333 * T2 - 0.00000347 * T3;
    double Mpr = 306.0253 + 385.81691806 * k + 0.0107306 * T2 + 0.00001236 * T3;
    double F = 21.2964 + 390.67050646 * k - 0.0016528 * T2 - 0.00000239 * T3;
    double C1 = (0.1734 - 0.000393 * T) * sin(M * dr) + 0.0021 * sin(2 * dr * M);
    C1 = C1 - 0.4068 * sin(Mpr * dr) + 0.0161 * sin(dr * 2 * Mpr);
    C1 = C1 - 0.0004 * sin(dr * 3 * Mpr);
    C1 = C1 + 0.0104 * sin(dr * 2 * F) - 0.0051 * sin(dr * (M + Mpr));
    C1 = C1 - 0.0074 * sin(dr * (M - Mpr)) + 0.0004 * sin(dr * (2 * F + M));
    C1 = C1 - 0.0004 * sin(dr * (2 * F - M)) - 0.0006 * sin(dr * (2 * F + Mpr));
    C1 = C1 + 0.0010 * sin(dr * (2 * F - Mpr)) + 0.0005 * sin(dr * (2 * Mpr + M));
    double deltat;
    if (T < -11) {
        deltat = 0.001 + 0.000839 * T + 0.0002261 * T2 - 0.00000845 * T3 - 0.000000081 * T * T3;
    } else {
        deltat = -0.000278 + 0.000265 * T + 0.000262 * T2;
    }
    double JdNew = Jd1 + C1 - deltat;
    return IntFloor(JdNew + 0.5 + timeZone / 24.0);
}

static int getSunLongitude(int jdn, int timeZone) {
    double T = (jdn - 2451545.5 - timeZone / 24.0) / 36525;
    double T2 = T * T;
    double dr = PI / 180.0;
    double M = 357.52910 + 35999.05030 * T - 0.0001559 * T2 - 0.00000048 * T * T2;
    double L0 = 280.46645 + 36000.76983 * T + 0.0003032 * T2;
    double DL = (1.914600 - 0.004817 * T - 0.000014 * T2) * sin(dr * M);
    DL = DL + (0.019993 - 0.000101 * T) * sin(dr * 2 * M) + 0.000290 * sin(dr * 3 * M);
    double L = L0 + DL;
    double omega = 125.04 - 1934.136 * T;
    L = L - 0.00569 - 0.00478 * sin(omega * dr);
    L = L * dr;
    L = L - PI * 2 * IntFloor(L / (PI * 2));
    return IntFloor(L / PI * 6);
}

static int getLunarMonth11(int yy, int timeZone) {
    int off = jdFromDate(31, 12, yy) - 2415021;
    int k = IntFloor(off / 29.530588853);
    int nm = getNewMoonDay(k, timeZone);
    int sunLong = getSunLongitude(nm, timeZone);
    if (sunLong >= 9) nm = getNewMoonDay(k - 1, timeZone);
    return nm;
}

static int getLeapMonthOffset(int a11, int timeZone) {
    int k = IntFloor(0.5 + (a11 - 2415021.076998695) / 29.530588853);
    int last = 0;
    int i = 1;
    int arc = getSunLongitude(getNewMoonDay(k + i, timeZone), timeZone);
    do {
        last = arc;
        i++;
        arc = getSunLongitude(getNewMoonDay(k + i, timeZone), timeZone);
    } while (arc != last && i < 14);
    return i - 1;
}

static BOOL TryGetOnlineLunar(int d, int m, int y, LunarDate* outLunar);

// Chuyển đổi Dương -> Âm (theo Ho Ngọc Đức)
static LunarDate SolarToLunar(int d, int m, int y) {
    LunarDate online;
    if (TryGetOnlineLunar(d, m, y, &online)) return online;

    int timeZone = 7;
    int dayNumber = jdFromDate(d, m, y);
    int k = IntFloor((dayNumber - 2415021.076998695) / 29.530588853);
    int monthStart = getNewMoonDay(k + 1, timeZone);
    if (monthStart > dayNumber) monthStart = getNewMoonDay(k, timeZone);
    int a11 = getLunarMonth11(y, timeZone);
    int b11 = a11;
    int lunarYear;
    if (a11 >= monthStart) {
        lunarYear = y;
        a11 = getLunarMonth11(y - 1, timeZone);
    } else {
        lunarYear = y + 1;
        b11 = getLunarMonth11(y + 1, timeZone);
    }
    int lunarDay = dayNumber - monthStart + 1;
    int diff = IntFloor((monthStart - a11) / 29.0);
    int lunarLeap = 0;
    int lunarMonth = diff + 11;
    if (b11 - a11 > 365) {
        int leapMonthDiff = getLeapMonthOffset(a11, timeZone);
        if (diff >= leapMonthDiff) {
            lunarMonth = diff + 10;
            if (diff == leapMonthDiff) lunarLeap = 1;
        }
    }
    if (lunarMonth > 12) lunarMonth -= 12;
    if (lunarMonth >= 11 && diff < 4) lunarYear -= 1;

    LunarDate res;
    res.day = lunarDay;
    res.month = lunarMonth;
    res.year = lunarYear;
    res.leap = lunarLeap;
    return res;
}

static int daysInMonth(int m, int y) {
    if (m == 2) return (y % 4 == 0 && y % 100 != 0) || (y % 400 == 0) ? 29 : 28;
    if (m == 4 || m == 6 || m == 9 || m == 11) return 30;
    return 31;
}

static int dayOfWeek(int d, int m, int y) {
    static int t[] = {0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4};
    if (m < 3) y -= 1;
    return (y + y/4 - y/100 + y/400 + t[m-1] + d) % 7;
}

static COLORREF BlendColor(COLORREF c1, COLORREF c2, float t) {
    if (t < 0.0f) t = 0.0f;
    if (t > 1.0f) t = 1.0f;
    int r1 = GetRValue(c1), g1 = GetGValue(c1), b1 = GetBValue(c1);
    int r2 = GetRValue(c2), g2 = GetGValue(c2), b2 = GetBValue(c2);
    int r = (int)(r1 + (r2 - r1) * t);
    int g = (int)(g1 + (g2 - g1) * t);
    int b = (int)(b1 + (b2 - b1) * t);
    return RGB(r, g, b);
}

static DWORD PremultiplyColor(COLORREF c, BYTE a) {
    DWORD r = (DWORD)GetRValue(c) * a / 255;
    DWORD g = (DWORD)GetGValue(c) * a / 255;
    DWORD b = (DWORD)GetBValue(c) * a / 255;
    return (DWORD)(a << 24) | (r << 16) | (g << 8) | b;
}

static int ClampInt(int v, int minV, int maxV) {
    if (v < minV) return minV;
    if (v > maxV) return maxV;
    return v;
}

static int CenterY(int rowTop, int rowH, int ctrlH) {
    return rowTop + (rowH - ctrlH) / 2;
}

static int GetTodayYmd(SYSTEMTIME* outSt) {
    SYSTEMTIME st;
    SYSTEMTIME* p = outSt ? outSt : &st;
    GetLocalTime(p);
    return p->wYear * 10000 + p->wMonth * 100 + p->wDay;
}

static BOOL IsSupportedWindows() {
    RTL_OSVERSIONINFOW vi;
    ZeroMemory(&vi, sizeof(vi));
    vi.dwOSVersionInfoSize = sizeof(vi);
    HMODULE hNt = GetModuleHandleW(L"ntdll.dll");
    if (hNt) {
        typedef LONG (WINAPI *RtlGetVersion_t)(PRTL_OSVERSIONINFOW);
        RtlGetVersion_t pRtlGetVersion = (RtlGetVersion_t)GetProcAddress(hNt, "RtlGetVersion");
        if (pRtlGetVersion && pRtlGetVersion(&vi) == 0) {
            return (vi.dwMajorVersion >= 10);
        }
    }
    OSVERSIONINFOW osv;
    ZeroMemory(&osv, sizeof(osv));
    osv.dwOSVersionInfoSize = sizeof(osv);
    if (GetVersionExW(&osv)) {
        return (osv.dwMajorVersion >= 10);
    }
    return FALSE;
}

static int Luminance(COLORREF c) {
    int r = GetRValue(c);
    int g = GetGValue(c);
    int b = GetBValue(c);
    return (int)(0.299f * r + 0.587f * g + 0.114f * b);
}

static COLORREF PickTextColorOn(COLORREF bg, COLORREF darkText, COLORREF lightText) {
    return (Luminance(bg) > 160) ? darkText : lightText;
}

static float GetResolutionScale() {
    int h = GetSystemMetrics(SM_CYSCREEN);
    if (h >= 2160) return 1.40f;
    if (h >= 2000) return 1.35f;
    if (h >= 1600) return 1.25f;
    if (h >= 1440) return 1.15f;
    return 1.00f;
}

static int GetVisibleRowCount(int mm, int yy) {
    int days = daysInMonth(mm, yy);
    int first = dayOfWeek(1, mm, yy);
    int cells = first + days;
    int rows = (cells + 6) / 7;
    if (rows < 4) rows = 4;
    if (rows > 6) rows = 6;
    return rows;
}

static void SyncScaleFromConfig() {
    if (g_cfg.solarScalePct < 90) g_cfg.solarScalePct = 90;
    if (g_cfg.solarScalePct > 180) g_cfg.solarScalePct = 180;
    if (g_cfg.emojiScalePct < 70) g_cfg.emojiScalePct = 70;
    if (g_cfg.emojiScalePct > 160) g_cfg.emojiScalePct = 160;
    if (g_cfg.lunarScalePct < 60) g_cfg.lunarScalePct = 60;
    if (g_cfg.lunarScalePct > 120) g_cfg.lunarScalePct = 120;
    g_solarScale = g_cfg.solarScalePct / 100.0f;
    g_emojiScale = g_cfg.emojiScalePct / 100.0f;
    g_lunarScale = g_cfg.lunarScalePct / 100.0f;
}

static float DistanceToRoundedRect(int x, int y, RECT rc, int radius) {
    float px = (float)x + 0.5f;
    float py = (float)y + 0.5f;
    float left = (float)rc.left + radius;
    float right = (float)rc.right - radius;
    float top = (float)rc.top + radius;
    float bottom = (float)rc.bottom - radius;
    float dx = 0.0f, dy = 0.0f;
    if (px < left) dx = left - px;
    else if (px > right) dx = px - right;
    if (py < top) dy = top - py;
    else if (py > bottom) dy = py - bottom;
    return (float)sqrt(dx * dx + dy * dy) - radius;
}

static void RestoreCardAlpha(DWORD* pixel, int radius) {
    if (!pixel) return;
    int w = g_cfg.width;
    for (int y = g_cardRect.top; y < g_cardRect.bottom; y++) {
        for (int x = g_cardRect.left; x < g_cardRect.right; x++) {
            if (DistanceToRoundedRect(x, y, g_cardRect, radius) <= 0.0f) {
                DWORD p = pixel[y * w + x];
                pixel[y * w + x] = (0xFFu << 24) | (p & 0x00FFFFFF);
            }
        }
    }
}

// ... CÁC HÀM UI (KHÔNG THAY ĐỔI LOGIC) ...

static void GetConfigPath(wchar_t* path, int size) {
    wchar_t* appdata = _wgetenv(L"APPDATA");
    if (!appdata) appdata = L".";
    swprintf(path, size, L"%s\\VietLunarCal", appdata);
    CreateDirectoryW(path, NULL);
    swprintf(path, size, L"%s\\VietLunarCal\\config.ini", appdata);
}

static float ReadIniFloat(const wchar_t* section, const wchar_t* key, float def, const wchar_t* path) {
    wchar_t buf[64];
    swprintf(buf, 64, L"%.4f", def);
    GetPrivateProfileStringW(section, key, buf, buf, 64, path);
    return (float)_wtof(buf);
}

static void LoadConfig() {
    wchar_t path[MAX_PATH];
    GetConfigPath(path, MAX_PATH);
    g_cfg.x = GetPrivateProfileIntW(L"Pos", L"X", -1, path);
    g_cfg.y = GetPrivateProfileIntW(L"Pos", L"Y", -1, path);
    g_cfg.alpha = GetPrivateProfileIntW(L"Style", L"Alpha", 230, path);
    g_cfg.autoRun = GetPrivateProfileIntW(L"Settings", L"AutoRun", 0, path);
    g_cfg.stickDesktop = GetPrivateProfileIntW(L"Settings", L"StickDesktop", 1, path);
    g_cfg.primaryScreenOnly = GetPrivateProfileIntW(L"Settings", L"PrimaryScreenOnly", 0, path);
    g_cfg.allowMove = GetPrivateProfileIntW(L"Settings", L"AllowMove", 1, path);
    g_cfg.bgColor = GetPrivateProfileIntW(L"Style", L"BgColor", RGB(248,249,252), path);
    g_cfg.textColor = GetPrivateProfileIntW(L"Style", L"TextColor", RGB(24,32,45), path);
    g_cfg.highlightColor = GetPrivateProfileIntW(L"Style", L"HighlightColor", RGB(14,120,200), path);
    g_cfg.lunarColor = GetPrivateProfileIntW(L"Style", L"LunarColor", RGB(90,103,120), path);
    g_cfg.fontSize = GetPrivateProfileIntW(L"Style", L"FontSize", 11, path);
    GetPrivateProfileStringW(L"Style", L"FontName", L"Segoe UI Variable", g_cfg.fontName, 32, path);
    GetPrivateProfileStringW(L"Style", L"EmojiFontColor", L"Segoe UI Emoji", g_cfg.emojiFontColor, 64, path);
    GetPrivateProfileStringW(L"Style", L"EmojiFontMono", L"Segoe UI Symbol", g_cfg.emojiFontMono, 64, path);
    g_cfg.emojiMono = GetPrivateProfileIntW(L"Style", L"EmojiMono", 0, path);
    g_cfg.solarScalePct = GetPrivateProfileIntW(L"Style", L"SolarScale", 125, path);
    g_cfg.emojiScalePct = GetPrivateProfileIntW(L"Style", L"EmojiScale", 100, path);
    g_cfg.lunarScalePct = GetPrivateProfileIntW(L"Style", L"LunarScale", 85, path);
    g_cfg.relX = ReadIniFloat(L"Pos", L"RelX", -1.0f, path);
    g_cfg.relY = ReadIniFloat(L"Pos", L"RelY", -1.0f, path);
    GetPrivateProfileStringW(L"Pos", L"MonitorId", L"", g_cfg.monitorId, 64, path);
}

static void SaveConfig() {
    wchar_t path[MAX_PATH], buf[32];
    GetConfigPath(path, MAX_PATH);
    swprintf(buf, 32, L"%d", g_cfg.x); WritePrivateProfileStringW(L"Pos", L"X", buf, path);
    swprintf(buf, 32, L"%d", g_cfg.y); WritePrivateProfileStringW(L"Pos", L"Y", buf, path);
    swprintf(buf, 32, L"%.4f", g_cfg.relX); WritePrivateProfileStringW(L"Pos", L"RelX", buf, path);
    swprintf(buf, 32, L"%.4f", g_cfg.relY); WritePrivateProfileStringW(L"Pos", L"RelY", buf, path);
    WritePrivateProfileStringW(L"Pos", L"MonitorId", g_cfg.monitorId, path);
    swprintf(buf, 32, L"%d", g_cfg.alpha); WritePrivateProfileStringW(L"Style", L"Alpha", buf, path);
    swprintf(buf, 32, L"%d", g_cfg.bgColor); WritePrivateProfileStringW(L"Style", L"BgColor", buf, path);
    swprintf(buf, 32, L"%d", g_cfg.textColor); WritePrivateProfileStringW(L"Style", L"TextColor", buf, path);
    swprintf(buf, 32, L"%d", g_cfg.highlightColor); WritePrivateProfileStringW(L"Style", L"HighlightColor", buf, path);
    swprintf(buf, 32, L"%d", g_cfg.lunarColor); WritePrivateProfileStringW(L"Style", L"LunarColor", buf, path);
    swprintf(buf, 32, L"%d", g_cfg.autoRun); WritePrivateProfileStringW(L"Settings", L"AutoRun", buf, path);
    swprintf(buf, 32, L"%d", g_cfg.stickDesktop); WritePrivateProfileStringW(L"Settings", L"StickDesktop", buf, path);
    swprintf(buf, 32, L"%d", g_cfg.primaryScreenOnly); WritePrivateProfileStringW(L"Settings", L"PrimaryScreenOnly", buf, path);
    swprintf(buf, 32, L"%d", g_cfg.allowMove); WritePrivateProfileStringW(L"Settings", L"AllowMove", buf, path);
    swprintf(buf, 32, L"%d", g_cfg.fontSize); WritePrivateProfileStringW(L"Style", L"FontSize", buf, path);
    WritePrivateProfileStringW(L"Style", L"FontName", g_cfg.fontName, path);
    WritePrivateProfileStringW(L"Style", L"EmojiFontColor", g_cfg.emojiFontColor, path);
    WritePrivateProfileStringW(L"Style", L"EmojiFontMono", g_cfg.emojiFontMono, path);
    swprintf(buf, 32, L"%d", g_cfg.emojiMono); WritePrivateProfileStringW(L"Style", L"EmojiMono", buf, path);
    swprintf(buf, 32, L"%d", g_cfg.solarScalePct); WritePrivateProfileStringW(L"Style", L"SolarScale", buf, path);
    swprintf(buf, 32, L"%d", g_cfg.emojiScalePct); WritePrivateProfileStringW(L"Style", L"EmojiScale", buf, path);
    swprintf(buf, 32, L"%d", g_cfg.lunarScalePct); WritePrivateProfileStringW(L"Style", L"LunarScale", buf, path);
}

static UINT GetDpiForWindowCompat(HWND hwnd) {
    typedef UINT (WINAPI *GetDpiForWindow_t)(HWND);
    static GetDpiForWindow_t pGetDpiForWindow = NULL;
    static BOOL resolved = FALSE;
    if (!resolved) {
        HMODULE hUser32 = GetModuleHandleW(L"user32.dll");
        if (hUser32) pGetDpiForWindow = (GetDpiForWindow_t)GetProcAddress(hUser32, "GetDpiForWindow");
        resolved = TRUE;
    }
    if (pGetDpiForWindow) return pGetDpiForWindow(hwnd);
    HDC hdc = GetDC(hwnd);
    UINT dpi = (UINT)GetDeviceCaps(hdc, LOGPIXELSX);
    ReleaseDC(hwnd, hdc);
    return dpi ? dpi : 96;
}

static int ScaleByDpi(int value, UINT dpi) {
    return MulDiv(value, (int)dpi, 96);
}

static HFONT EnsureSettingsFont(UINT dpi) {
    if (g_settingsFont && g_settingsFontDpi == dpi) return g_settingsFont;
    if (g_settingsFont) {
        DeleteObject(g_settingsFont);
        g_settingsFont = NULL;
    }
    NONCLIENTMETRICS ncm = {0};
    ncm.cbSize = sizeof(ncm);
    typedef BOOL (WINAPI *SystemParametersInfoForDpi_t)(UINT, UINT, PVOID, UINT, UINT);
    static SystemParametersInfoForDpi_t pSpiForDpi = NULL;
    static BOOL spiResolved = FALSE;
    if (!spiResolved) {
        HMODULE hUser32 = GetModuleHandleW(L"user32.dll");
        if (hUser32) pSpiForDpi = (SystemParametersInfoForDpi_t)GetProcAddress(hUser32, "SystemParametersInfoForDpi");
        spiResolved = TRUE;
    }
    BOOL ok = FALSE;
    if (pSpiForDpi) {
        ok = pSpiForDpi(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0, dpi);
    } else {
        ok = SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, sizeof(ncm), &ncm, 0);
    }
    if (!ok) {
        LOGFONTW lf = {0};
        lf.lfHeight = -MulDiv(9, (int)dpi, 72);
        wcscpy(lf.lfFaceName, L"Segoe UI");
        g_settingsFont = CreateFontIndirectW(&lf);
        g_settingsFontDpi = dpi;
        return g_settingsFont;
    }
    LOGFONTW lf = ncm.lfMessageFont;
    if (!pSpiForDpi) {
        lf.lfHeight = MulDiv(lf.lfHeight, (int)dpi, 96);
    }
    g_settingsFont = CreateFontIndirectW(&lf);
    g_settingsFontDpi = dpi;
    return g_settingsFont;
}

static void ApplySettingsFont(HWND hwnd, HFONT font) {
    if (!font) return;
    int ids[] = {
        ID_LBL_TITLE, ID_LBL_POSX, ID_LBL_POSY, ID_LBL_OPACITY, ID_LBL_EMOJI_STYLE,
        ID_LBL_SOLAR, ID_LBL_EMOJI, ID_LBL_LUNAR,
        ID_EDIT_X, ID_EDIT_Y, ID_OPACITY_TRACK, ID_BTN_COLOR, ID_BTN_FONT, ID_LBL_FONTINFO,
        ID_SLIDER_SOLAR, ID_SLIDER_EMOJI, ID_SLIDER_LUNAR,
        ID_EDIT_SOLAR, ID_EDIT_EMOJI, ID_EDIT_LUNAR,
        ID_LBL_PCT_SOLAR, ID_LBL_PCT_EMOJI, ID_LBL_PCT_LUNAR,
        ID_EMOJI_STYLE, ID_BTN_EMOJI_FONT, ID_LBL_EMOJI_FONT, ID_CHK_AUTORUN, ID_CHK_STICK,
        ID_CHK_PRIMARY, ID_CHK_MOVE, ID_APPLY, IDOK, IDCANCEL
    };
    for (int i = 0; i < (int)(sizeof(ids)/sizeof(ids[0])); i++) {
        HWND h = GetDlgItem(hwnd, ids[i]);
        if (h) SendMessageW(h, WM_SETFONT, (WPARAM)font, TRUE);
    }
}

static void UpdateScaleEditsFromSliders(HWND hwnd) {
    if (g_syncingScaleEdits) return;
    g_syncingScaleEdits = TRUE;
    wchar_t buf[16];
    int solar = (int)SendMessage(GetDlgItem(hwnd, ID_SLIDER_SOLAR), TBM_GETPOS, 0, 0);
    int emoji = (int)SendMessage(GetDlgItem(hwnd, ID_SLIDER_EMOJI), TBM_GETPOS, 0, 0);
    int lunar = (int)SendMessage(GetDlgItem(hwnd, ID_SLIDER_LUNAR), TBM_GETPOS, 0, 0);
    swprintf(buf, 16, L"%d", solar);
    SetWindowTextW(GetDlgItem(hwnd, ID_EDIT_SOLAR), buf);
    swprintf(buf, 16, L"%d", emoji);
    SetWindowTextW(GetDlgItem(hwnd, ID_EDIT_EMOJI), buf);
    swprintf(buf, 16, L"%d", lunar);
    SetWindowTextW(GetDlgItem(hwnd, ID_EDIT_LUNAR), buf);
    g_syncingScaleEdits = FALSE;
}

static void ApplyScaleEditToSlider(HWND hwnd, int editId, int sliderId, int minV, int maxV) {
    if (g_syncingScaleEdits) return;
    wchar_t buf[16];
    GetWindowTextW(GetDlgItem(hwnd, editId), buf, 16);
    int val = _wtoi(buf);
    val = ClampInt(val, minV, maxV);
    g_syncingScaleEdits = TRUE;
    swprintf(buf, 16, L"%d", val);
    SetWindowTextW(GetDlgItem(hwnd, editId), buf);
    SendMessage(GetDlgItem(hwnd, sliderId), TBM_SETPOS, TRUE, val);
    g_syncingScaleEdits = FALSE;
}

static int GetClampedEditValue(HWND hwnd, int editId, int minV, int maxV, int fallback) {
    wchar_t buf[16];
    GetWindowTextW(GetDlgItem(hwnd, editId), buf, 16);
    int val = fallback;
    if (buf[0] != 0) val = _wtoi(buf);
    val = ClampInt(val, minV, maxV);
    swprintf(buf, 16, L"%d", val);
    SetWindowTextW(GetDlgItem(hwnd, editId), buf);
    return val;
}

static void SetSettingsWindowSize(HWND hwnd, UINT dpi, int clientW, int clientH) {
    RECT rc = {0, 0, clientW, clientH};
    DWORD style = (DWORD)GetWindowLongW(hwnd, GWL_STYLE);
    DWORD exStyle = (DWORD)GetWindowLongW(hwnd, GWL_EXSTYLE);
    typedef BOOL (WINAPI *AdjustWindowRectExForDpi_t)(LPRECT, DWORD, BOOL, DWORD, UINT);
    static AdjustWindowRectExForDpi_t pAdjust = NULL;
    static BOOL resolved = FALSE;
    if (!resolved) {
        HMODULE hUser32 = GetModuleHandleW(L"user32.dll");
        if (hUser32) pAdjust = (AdjustWindowRectExForDpi_t)GetProcAddress(hUser32, "AdjustWindowRectExForDpi");
        resolved = TRUE;
    }
    if (pAdjust) pAdjust(&rc, style, FALSE, exStyle, dpi);
    else AdjustWindowRectEx(&rc, style, FALSE, exStyle);
    SetWindowPos(hwnd, NULL, 0, 0, rc.right - rc.left, rc.bottom - rc.top,
        SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
}

static void LayoutSettingsControls(HWND hwnd, UINT dpi) {
    int m = ScaleByDpi(20, dpi);
    int y = ScaleByDpi(10, dpi);
    int rowH = ScaleByDpi(24, dpi);
    int editH = ScaleByDpi(24, dpi);
    int btnH = ScaleByDpi(32, dpi);
    int trackH = ScaleByDpi(30, dpi);
    int gapY = ScaleByDpi(8, dpi);
    int gapX = ScaleByDpi(8, dpi);
    int labelW = ScaleByDpi(100, dpi);
    int editW = ScaleByDpi(100, dpi);
    int smallBtnW = ScaleByDpi(100, dpi);
    int comboW = ScaleByDpi(140, dpi);
    int scaleEditW = ScaleByDpi(52, dpi);
    int pctW = ScaleByDpi(16, dpi);
    int clientW = ScaleByDpi(520, dpi);
    int clientH = ScaleByDpi(500, dpi);

    SetSettingsWindowSize(hwnd, dpi, clientW, clientH);
    HFONT font = EnsureSettingsFont(dpi);
    ApplySettingsFont(hwnd, font);

    MoveWindow(GetDlgItem(hwnd, ID_LBL_TITLE), m, y, clientW - m * 2, rowH, TRUE);
    y += ScaleByDpi(30, dpi);

    int rowHPos = (rowH > editH) ? rowH : editH;
    MoveWindow(GetDlgItem(hwnd, ID_LBL_POSX), m, CenterY(y, rowHPos, rowH), labelW, rowH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_EDIT_X), m + labelW, CenterY(y, rowHPos, editH), editW, editH, TRUE);
    int x2 = m + labelW + editW + ScaleByDpi(20, dpi);
    MoveWindow(GetDlgItem(hwnd, ID_LBL_POSY), x2, CenterY(y, rowHPos, rowH), labelW, rowH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_EDIT_Y), x2 + labelW, CenterY(y, rowHPos, editH), editW, editH, TRUE);
    y += rowHPos + gapY;

    int rowHOpacity = trackH;
    MoveWindow(GetDlgItem(hwnd, ID_LBL_OPACITY), m, CenterY(y, rowHOpacity, rowH), labelW, rowH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_OPACITY_TRACK), m + labelW, CenterY(y, rowHOpacity, trackH),
        clientW - (m + labelW) - m, trackH, TRUE);
    y += rowHOpacity + gapY;

    int rowHButtons = btnH;
    MoveWindow(GetDlgItem(hwnd, ID_BTN_COLOR), m, CenterY(y, rowHButtons, btnH), ScaleByDpi(140, dpi), btnH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_BTN_FONT), m + ScaleByDpi(150, dpi), CenterY(y, rowHButtons, btnH), smallBtnW, btnH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_LBL_FONTINFO), m + ScaleByDpi(260, dpi), CenterY(y, rowHButtons, rowH),
        clientW - (m + ScaleByDpi(260, dpi)) - m, rowH, TRUE);
    y += rowHButtons + gapY;

    int sliderW = clientW - (m + labelW) - m - scaleEditW - pctW - gapX * 2;
    int rowHSlider = trackH;
    MoveWindow(GetDlgItem(hwnd, ID_LBL_SOLAR), m, CenterY(y, rowHSlider, rowH), labelW, rowH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_SLIDER_SOLAR), m + labelW, CenterY(y, rowHSlider, trackH),
        sliderW, trackH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_EDIT_SOLAR), m + labelW + sliderW + gapX, CenterY(y, rowHSlider, editH),
        scaleEditW, editH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_LBL_PCT_SOLAR), m + labelW + sliderW + gapX + scaleEditW + gapX,
        CenterY(y, rowHSlider, rowH), pctW, rowH, TRUE);
    y += rowHSlider + gapY;

    MoveWindow(GetDlgItem(hwnd, ID_LBL_EMOJI), m, CenterY(y, rowHSlider, rowH), labelW, rowH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_SLIDER_EMOJI), m + labelW, CenterY(y, rowHSlider, trackH),
        sliderW, trackH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_EDIT_EMOJI), m + labelW + sliderW + gapX, CenterY(y, rowHSlider, editH),
        scaleEditW, editH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_LBL_PCT_EMOJI), m + labelW + sliderW + gapX + scaleEditW + gapX,
        CenterY(y, rowHSlider, rowH), pctW, rowH, TRUE);
    y += rowHSlider + gapY;

    MoveWindow(GetDlgItem(hwnd, ID_LBL_LUNAR), m, CenterY(y, rowHSlider, rowH), labelW, rowH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_SLIDER_LUNAR), m + labelW, CenterY(y, rowHSlider, trackH),
        sliderW, trackH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_EDIT_LUNAR), m + labelW + sliderW + gapX, CenterY(y, rowHSlider, editH),
        scaleEditW, editH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_LBL_PCT_LUNAR), m + labelW + sliderW + gapX + scaleEditW + gapX,
        CenterY(y, rowHSlider, rowH), pctW, rowH, TRUE);
    y += rowHSlider + gapY;

    int rowHCombo = btnH;
    MoveWindow(GetDlgItem(hwnd, ID_LBL_EMOJI_STYLE), m, CenterY(y, rowHCombo, rowH), labelW, rowH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_EMOJI_STYLE), m + labelW, CenterY(y, rowHCombo, editH),
        comboW, ScaleByDpi(200, dpi), TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_BTN_EMOJI_FONT), m + labelW + comboW + ScaleByDpi(10, dpi), CenterY(y, rowHCombo, btnH),
        smallBtnW + ScaleByDpi(20, dpi), btnH, TRUE);
    MoveWindow(GetDlgItem(hwnd, ID_LBL_EMOJI_FONT),
        m + labelW + comboW + smallBtnW + ScaleByDpi(40, dpi), CenterY(y, rowHCombo, rowH),
        clientW - (m + labelW + comboW + smallBtnW + ScaleByDpi(40, dpi)) - m, rowH, TRUE);
    y += rowHCombo + gapY;

    MoveWindow(GetDlgItem(hwnd, ID_CHK_AUTORUN), m, y, clientW - m * 2, rowH, TRUE);
    y += rowH + ScaleByDpi(2, dpi);
    MoveWindow(GetDlgItem(hwnd, ID_CHK_STICK), m, y, clientW - m * 2, rowH, TRUE);
    y += rowH + ScaleByDpi(2, dpi);
    MoveWindow(GetDlgItem(hwnd, ID_CHK_PRIMARY), m, y, clientW - m * 2, rowH, TRUE);
    y += rowH + ScaleByDpi(2, dpi);
    MoveWindow(GetDlgItem(hwnd, ID_CHK_MOVE), m, y, clientW - m * 2, rowH, TRUE);

    int btnY = clientH - m - btnH;
    int btnW = ScaleByDpi(80, dpi);
    int btnGap = ScaleByDpi(20, dpi);
    int totalBtnW = btnW * 3 + btnGap * 2;
    int btnX = (clientW - totalBtnW) / 2;
    MoveWindow(GetDlgItem(hwnd, ID_APPLY), btnX, btnY, btnW, btnH, TRUE);
    MoveWindow(GetDlgItem(hwnd, IDOK), btnX + btnW + btnGap, btnY, btnW, btnH, TRUE);
    MoveWindow(GetDlgItem(hwnd, IDCANCEL), btnX + (btnW + btnGap) * 2, btnY, btnW, btnH, TRUE);
}

static BOOL ParseDateAfterLabel(const wchar_t* html, const wchar_t* label, int* d, int* m, int* y) {
    const wchar_t* p = wcsstr(html, label);
    if (!p) return FALSE;
    p += wcslen(label);
    while (*p && !iswdigit(*p)) p++;
    if (!*p) return FALSE;
    int dd = (int)wcstol(p, (wchar_t**)&p, 10);
    while (*p && !iswdigit(*p)) p++;
    if (!*p) return FALSE;
    int mm = (int)wcstol(p, (wchar_t**)&p, 10);
    while (*p && !iswdigit(*p)) p++;
    if (!*p) return FALSE;
    int yy = (int)wcstol(p, (wchar_t**)&p, 10);
    if (dd <= 0 || mm <= 0 || yy <= 0) return FALSE;
    *d = dd;
    *m = mm;
    *y = yy;
    return TRUE;
}

static BOOL DownloadUrlToWideBuffer(const wchar_t* url, wchar_t** outBuf) {
    *outBuf = NULL;
    HINTERNET hInternet = InternetOpenW(L"VietCalendar", INTERNET_OPEN_TYPE_PRECONFIG, NULL, NULL, 0);
    if (!hInternet) return FALSE;
    HINTERNET hUrl = InternetOpenUrlW(
        hInternet, url, NULL, 0,
        INTERNET_FLAG_RELOAD | INTERNET_FLAG_NO_CACHE_WRITE | INTERNET_FLAG_SECURE, 0);
    if (!hUrl) {
        InternetCloseHandle(hInternet);
        return FALSE;
    }
    BYTE* data = NULL;
    DWORD total = 0;
    BYTE buffer[4096];
    DWORD read = 0;
    BOOL ok = TRUE;
    while (InternetReadFile(hUrl, buffer, sizeof(buffer), &read) && read > 0) {
        BYTE* tmp = (BYTE*)realloc(data, total + read + 1);
        if (!tmp) { ok = FALSE; break; }
        data = tmp;
        memcpy(data + total, buffer, read);
        total += read;
    }
    if (!ok || total == 0) {
        if (data) free(data);
        InternetCloseHandle(hUrl);
        InternetCloseHandle(hInternet);
        return FALSE;
    }
    data[total] = 0;
    int wlen = MultiByteToWideChar(CP_UTF8, 0, (LPCCH)data, (int)total, NULL, 0);
    if (wlen <= 0) {
        free(data);
        InternetCloseHandle(hUrl);
        InternetCloseHandle(hInternet);
        return FALSE;
    }
    wchar_t* wbuf = (wchar_t*)calloc((size_t)wlen + 1, sizeof(wchar_t));
    if (!wbuf) {
        free(data);
        InternetCloseHandle(hUrl);
        InternetCloseHandle(hInternet);
        return FALSE;
    }
    MultiByteToWideChar(CP_UTF8, 0, (LPCCH)data, (int)total, wbuf, wlen);
    wbuf[wlen] = 0;
    free(data);
    InternetCloseHandle(hUrl);
    InternetCloseHandle(hInternet);
    *outBuf = wbuf;
    return TRUE;
}

static BOOL FetchOnlineTodayLunar(LunarDate* outLunar, SYSTEMTIME* outSolar) {
    DWORD flags = 0;
    if (!InternetGetConnectedState(&flags, 0)) return FALSE;
    wchar_t* html = NULL;
    if (!DownloadUrlToWideBuffer(L"https://www.xemlicham.com/", &html)) return FALSE;
    int sd = 0, sm = 0, sy = 0;
    int ld = 0, lm = 0, ly = 0;
    BOOL ok = ParseDateAfterLabel(html, L"Ng\u00E0y D\u01B0\u01A1ng L\u1ECBch:", &sd, &sm, &sy) &&
              ParseDateAfterLabel(html, L"Ng\u00E0y \u00C2m L\u1ECBch:", &ld, &lm, &ly);
    free(html);
    if (!ok) return FALSE;
    SYSTEMTIME local;
    GetLocalTime(&local);
    if (sd != local.wDay || sm != local.wMonth || sy != local.wYear) return FALSE;
    ZeroMemory(outSolar, sizeof(*outSolar));
    outSolar->wDay = (WORD)sd;
    outSolar->wMonth = (WORD)sm;
    outSolar->wYear = (WORD)sy;
    outLunar->day = ld;
    outLunar->month = lm;
    outLunar->year = ly;
    outLunar->leap = 0;
    return TRUE;
}

static DWORD WINAPI OnlineCheckThread(LPVOID param) {
    int targetYMD = (int)(INT_PTR)param;
    LunarDate lunar = {0};
    SYSTEMTIME solar = {0};
    BOOL ok = FetchOnlineTodayLunar(&lunar, &solar);
    if (g_onlineLockInit) EnterCriticalSection(&g_onlineLock);
    if (ok) {
        g_onlineHasData = TRUE;
        g_onlineSolar = solar;
        g_onlineLunar = lunar;
        g_onlineLastSuccessYMD = targetYMD;
    }
    g_onlineNextAttemptTick = GetTickCount() + (ok ? 6 * 60 * 60 * 1000 : 30 * 60 * 1000);
    if (g_onlineLockInit) LeaveCriticalSection(&g_onlineLock);
    InterlockedExchange(&g_onlineCheckInProgress, 0);
    if (ok && g_hwnd) PostMessageW(g_hwnd, WM_ONLINE_UPDATE, 0, 0);
    return 0;
}

static void MaybeStartOnlineCheck() {
    SYSTEMTIME st;
    GetLocalTime(&st);
    int ymd = st.wYear * 10000 + st.wMonth * 100 + st.wDay;
    DWORD now = GetTickCount();
    if (ymd == g_onlineLastSuccessYMD) {
        if (g_onlineNextAttemptTick && now < g_onlineNextAttemptTick) return;
    }
    if (InterlockedCompareExchange(&g_onlineCheckInProgress, 1, 0) != 0) return;
    HANDLE hThread = CreateThread(NULL, 0, OnlineCheckThread, (LPVOID)(INT_PTR)ymd, 0, NULL);
    if (hThread) CloseHandle(hThread);
    else InterlockedExchange(&g_onlineCheckInProgress, 0);
}

static BOOL TryGetOnlineLunar(int d, int m, int y, LunarDate* outLunar) {
    if (!outLunar || !g_onlineHasData || !g_onlineLockInit) return FALSE;
    BOOL ok = FALSE;
    EnterCriticalSection(&g_onlineLock);
    if (g_onlineHasData &&
        g_onlineSolar.wDay == d &&
        g_onlineSolar.wMonth == m &&
        g_onlineSolar.wYear == y) {
        *outLunar = g_onlineLunar;
        ok = TRUE;
    }
    LeaveCriticalSection(&g_onlineLock);
    return ok;
}

static void SetAutoRun(BOOL enable) {
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Run", 0, KEY_WRITE, &hKey) == ERROR_SUCCESS) {
        if (enable) {
            wchar_t exePath[MAX_PATH];
            GetModuleFileNameW(NULL, exePath, MAX_PATH);
            RegSetValueExW(hKey, L"VietCalendar", 0, REG_SZ, (BYTE*)exePath, (wcslen(exePath)+1)*2);
        } else {
            RegDeleteValueW(hKey, L"VietCalendar");
        }
        RegCloseKey(hKey);
    }
}

static void GetPrimaryScreenRect(RECT* rc) {
    rc->left = 0;
    rc->top = 0;
    rc->right = GetSystemMetrics(SM_CXSCREEN);
    rc->bottom = GetSystemMetrics(SM_CYSCREEN);
    SystemParametersInfoW(SPI_GETWORKAREA, 0, rc, 0);
}

typedef struct {
    const wchar_t* target;
    RECT rc;
    BOOL found;
} MonitorFindCtx;

static BOOL CALLBACK FindMonitorByDeviceProc(HMONITOR hMon, HDC hdc, LPRECT rcMon, LPARAM data) {
    MonitorFindCtx* ctx = (MonitorFindCtx*)data;
    MONITORINFOEXW mi;
    ZeroMemory(&mi, sizeof(mi));
    mi.cbSize = sizeof(mi);
    if (GetMonitorInfoW(hMon, (MONITORINFO*)&mi)) {
        if (ctx->target && wcscmp(mi.szDevice, ctx->target) == 0) {
            ctx->rc = mi.rcWork;
            ctx->found = TRUE;
            return FALSE;
        }
    }
    return TRUE;
}

static BOOL GetMonitorWorkAreaByDevice(const wchar_t* device, RECT* rc) {
    MonitorFindCtx ctx;
    ZeroMemory(&ctx, sizeof(ctx));
    ctx.target = device;
    EnumDisplayMonitors(NULL, NULL, FindMonitorByDeviceProc, (LPARAM)&ctx);
    if (ctx.found && rc) *rc = ctx.rc;
    return ctx.found;
}

static void GetMonitorWorkAreaFromPoint(POINT pt, RECT* rc, wchar_t* device, int deviceLen) {
    HMONITOR hMon = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
    MONITORINFOEXW mi;
    ZeroMemory(&mi, sizeof(mi));
    mi.cbSize = sizeof(mi);
    GetMonitorInfoW(hMon, (MONITORINFO*)&mi);
    if (rc) *rc = mi.rcWork;
    if (device && deviceLen > 0) {
        wcsncpy(device, mi.szDevice, deviceLen - 1);
        device[deviceLen - 1] = L'\0';
    }
}

static void UpdateRelativePositionFromCurrent() {
    RECT rc;
    wchar_t device[64] = {0};
    POINT pt = {g_cfg.x + g_cfg.width / 2, g_cfg.y + g_cfg.height / 2};
    GetMonitorWorkAreaFromPoint(pt, &rc, device, 64);
    if (!g_cfg.primaryScreenOnly) {
        wcsncpy(g_cfg.monitorId, device, 63);
        g_cfg.monitorId[63] = L'\0';
    }
    int maxX = max(1, rc.right - rc.left - g_cfg.width);
    int maxY = max(1, rc.bottom - rc.top - g_cfg.height);
    g_cfg.relX = (float)(g_cfg.x - rc.left) / (float)maxX;
    g_cfg.relY = (float)(g_cfg.y - rc.top) / (float)maxY;
    if (g_cfg.relX < 0.0f) g_cfg.relX = 0.0f;
    if (g_cfg.relX > 1.0f) g_cfg.relX = 1.0f;
    if (g_cfg.relY < 0.0f) g_cfg.relY = 0.0f;
    if (g_cfg.relY > 1.0f) g_cfg.relY = 1.0f;
}

static void RestorePositionFromConfig() {
    RECT rc;
    BOOL ok = FALSE;
    if (g_cfg.primaryScreenOnly) {
        GetPrimaryScreenRect(&rc);
        ok = TRUE;
    } else if (g_cfg.monitorId[0]) {
        ok = GetMonitorWorkAreaByDevice(g_cfg.monitorId, &rc);
    }
    if (!ok) {
        POINT pt = {g_cfg.x, g_cfg.y};
        GetMonitorWorkAreaFromPoint(pt, &rc, NULL, 0);
    }
    int maxX = max(1, rc.right - rc.left - g_cfg.width);
    int maxY = max(1, rc.bottom - rc.top - g_cfg.height);
    if (g_cfg.relX >= 0.0f && g_cfg.relY >= 0.0f) {
        g_cfg.x = rc.left + (int)roundf(g_cfg.relX * maxX);
        g_cfg.y = rc.top + (int)roundf(g_cfg.relY * maxY);
    }
    g_cfg.x = ClampInt(g_cfg.x, rc.left, rc.right - g_cfg.width);
    g_cfg.y = ClampInt(g_cfg.y, rc.top, rc.bottom - g_cfg.height);
}

static void ClampPositionToVisibleMonitor() {
    RECT rc;
    if (g_cfg.primaryScreenOnly) {
        GetPrimaryScreenRect(&rc);
    } else {
        POINT pt = {g_cfg.x, g_cfg.y};
        GetMonitorWorkAreaFromPoint(pt, &rc, NULL, 0);
    }
    g_cfg.x = ClampInt(g_cfg.x, rc.left, rc.right - g_cfg.width);
    g_cfg.y = ClampInt(g_cfg.y, rc.top, rc.bottom - g_cfg.height);
}

static void SendToBack() {
    if (g_cfg.stickDesktop) {
        SetWindowPos(g_hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
    }
}

static void UpdateStickTimer(HWND hwnd) {
    if (!hwnd) return;
    if (g_cfg.stickDesktop) {
        SetTimer(hwnd, 2, 500, NULL);
    } else {
        KillTimer(hwnd, 2);
    }
}

static void UpdateClickThroughStyle() {
    if (!g_hwnd) return;
    LONG_PTR ex = GetWindowLongPtrW(g_hwnd, GWL_EXSTYLE);
    ex &= ~WS_EX_TRANSPARENT;
    SetWindowLongPtrW(g_hwnd, GWL_EXSTYLE, ex);
    SetWindowPos(g_hwnd, NULL, 0, 0, 0, 0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
}

static void DrawCalendar(HWND hwnd);

static void SetHeaderHover(HWND hwnd, BOOL hover) {
    if (hover == g_hoverHeader) return;
    g_hoverHeader = hover;
    if (hover) {
        g_headerPulse = 0.0f;
        SetTimer(hwnd, 3, 30, NULL);
    } else {
        KillTimer(hwnd, 3);
        g_headerPulse = 0.0f;
    }
    DrawCalendar(hwnd);
}

static void CalculateMetrics() {
    HDC hdc = GetDC(NULL);
    float baseScale = GetResolutionScale();
    float effectiveFont = g_cfg.fontSize * baseScale;
    for (int pass = 0; pass < 2; pass++) {
        int bigSize = (int)(effectiveFont * g_solarScale * g_dpiScale);
        int lunarSize = (int)((effectiveFont - 2.0f) * g_lunarScale * g_dpiScale);
        int emojiSize = (int)((effectiveFont * g_emojiScale + 1.0f) * g_dpiScale);
        if (lunarSize < (int)(6 * g_dpiScale)) lunarSize = (int)(6 * g_dpiScale);
        if (emojiSize < (int)(8 * g_dpiScale)) emojiSize = (int)(8 * g_dpiScale);

        HFONT hBig = CreateFontW(bigSize, 0, 0, 0, FW_BOLD, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY,
            VARIABLE_PITCH, g_cfg.fontName);
        HFONT hLunar = CreateFontW(lunarSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY,
            VARIABLE_PITCH, g_cfg.fontName);
        HFONT oldFont = (HFONT)SelectObject(hdc, hBig);
        SIZE szNum, szLunar;
        GetTextExtentPoint32W(hdc, L"00", 2, &szNum);
        SelectObject(hdc, hLunar);
        GetTextExtentPoint32W(hdc, L"00/00", 5, &szLunar);
        SelectObject(hdc, oldFont);
        DeleteObject(hBig);
        DeleteObject(hLunar);

        int paddingX = (int)(8 * g_dpiScale);
        int bandPad = (int)(4 * g_dpiScale);
        int topBand = bigSize + bandPad * 2;
        int bottomBand = g_showLunar ? (lunarSize + bandPad * 2) : (bandPad * 2);
        int emojiBand = emojiSize + bandPad * 2;
        int minEmojiBand = (int)(16 * g_dpiScale);
        if (emojiBand < minEmojiBand) emojiBand = minEmojiBand;
        int baseHeight = topBand + emojiBand + bottomBand + (int)(2 * g_dpiScale);
        int textWidth = g_showLunar ? max(szNum.cx, szLunar.cx) : szNum.cx;
        int cellW = textWidth + (paddingX * 2);
        if (cellW < emojiSize + paddingX * 2) cellW = emojiSize + paddingX * 2;
        int cellH = baseHeight;
        if (cellW < (int)(36 * g_dpiScale)) cellW = (int)(36 * g_dpiScale);
        if (cellH < (int)(64 * g_dpiScale)) cellH = (int)(64 * g_dpiScale);
        int cell = max(cellW, cellH);

        if (pass == 0) {
            float fillScale = (float)cell / (float)baseHeight;
            if (fillScale > 1.05f) {
                if (fillScale > 1.30f) fillScale = 1.30f;
                effectiveFont *= fillScale;
                continue;
            }
        }

        g_effectiveFontSize = effectiveFont;
        g_fontScale = effectiveFont / (float)g_cfg.fontSize;
        g_cellWidth = cell;
        g_cellHeight = cell;
        break;
    }
    ReleaseDC(NULL, hdc);

    g_sideMargin = (int)(16 * g_dpiScale);
    g_topMargin = (int)(12 * g_dpiScale);
    g_bottomMargin = (int)(12 * g_dpiScale);
    g_dayHeaderHeight = (int)(24 * g_dpiScale);

    int titleHeight = (int)((g_effectiveFontSize + 6) * g_dpiScale);
    int subHeight = (int)((g_effectiveFontSize) * g_dpiScale);
    g_headerHeight = (int)(10 * g_dpiScale) + titleHeight + subHeight;

    g_cardWidth = (g_cellWidth * 7) + (g_sideMargin * 2);
    if (g_cardWidth < (int)(320 * g_dpiScale)) g_cardWidth = (int)(320 * g_dpiScale);
    g_visibleRows = GetVisibleRowCount(g_viewMonth, g_viewYear);
    g_cardHeight = g_topMargin + g_headerHeight + g_dayHeaderHeight + (g_cellHeight * g_visibleRows) + g_bottomMargin;

    g_shadowSize = (int)(12 * g_dpiScale);
    g_cfg.width = g_cardWidth + (g_shadowSize * 2);
    g_cfg.height = g_cardHeight + (g_shadowSize * 2);
    g_cardRect.left = g_shadowSize;
    g_cardRect.top = g_shadowSize;
    g_cardRect.right = g_shadowSize + g_cardWidth;
    g_cardRect.bottom = g_shadowSize + g_cardHeight;
}

static void UpdateEmojiTextFormats(void);

static void CreateFonts() {
    DeleteObject(g_fontBig);
    DeleteObject(g_fontSmall);
    DeleteObject(g_fontMonth);
    DeleteObject(g_fontLunar);
    DeleteObject(g_fontEmoji);
    float fs = g_effectiveFontSize;
    int bigSize = (int)(fs * g_solarScale * g_dpiScale);
    int lunarSize = (int)((fs - 2) * g_lunarScale * g_dpiScale);
    int smallSize = (int)((fs - 1) * g_dpiScale);
    int monthSize = (int)((fs + 4) * g_dpiScale);
    int emojiSize = (int)((fs * g_emojiScale + 1) * g_dpiScale);
    const wchar_t* emojiFace = g_cfg.emojiMono ? g_cfg.emojiFontMono : g_cfg.emojiFontColor;
    int emojiQuality = g_cfg.emojiMono ? ANTIALIASED_QUALITY : CLEARTYPE_QUALITY;
    g_fontBig = CreateFontW(bigSize, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, VARIABLE_PITCH, g_cfg.fontName);
    g_fontMonth = CreateFontW(monthSize, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, VARIABLE_PITCH, g_cfg.fontName);
    g_fontSmall = CreateFontW(smallSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, VARIABLE_PITCH, g_cfg.fontName);
    g_fontLunar = CreateFontW(lunarSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, ANTIALIASED_QUALITY, VARIABLE_PITCH, g_cfg.fontName);
    g_fontEmoji = CreateFontW(emojiSize, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
        OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, emojiQuality, VARIABLE_PITCH, emojiFace);
    UpdateEmojiTextFormats();
}

static void EnsureFontFaceOrFallback(wchar_t* faceName, const wchar_t* fallback) {
    HDC hdc = GetDC(NULL);
    if (!hdc) return;
    LOGFONTW lf = {0};
    wcscpy(lf.lfFaceName, faceName);
    lf.lfHeight = -MulDiv(max(g_cfg.fontSize, 8), GetDeviceCaps(hdc, LOGPIXELSY), 72);
    HFONT hFont = CreateFontIndirectW(&lf);
    HFONT oldFont = (HFONT)SelectObject(hdc, hFont);
    wchar_t face[LF_FACESIZE] = {0};
    GetTextFaceW(hdc, LF_FACESIZE, face);
    SelectObject(hdc, oldFont);
    DeleteObject(hFont);
    ReleaseDC(NULL, hdc);
    if (face[0] && _wcsicmp(face, faceName) != 0) {
        wcscpy(faceName, fallback);
    }
}

static void EnsureFontFallback() {
    EnsureFontFaceOrFallback(g_cfg.fontName, L"Segoe UI");
    EnsureFontFaceOrFallback(g_cfg.emojiFontColor, L"Segoe UI Emoji");
    EnsureFontFaceOrFallback(g_cfg.emojiFontMono, L"Segoe UI Symbol");
}

static BOOL EnsureDWrite() {
    HRESULT hr;
    if (!g_d2dFactory) {
        hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_SINGLE_THREADED, &IID_ID2D1Factory, NULL, (void**)&g_d2dFactory);
        if (FAILED(hr)) return FALSE;
    }
    if (!g_dwriteFactory) {
        hr = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED, &IID_IDWriteFactory, (IUnknown**)&g_dwriteFactory);
        if (FAILED(hr)) return FALSE;
    }
    if (!g_dwriteFactory2 && g_dwriteFactory) {
        IUnknown_QueryInterface((IUnknown*)g_dwriteFactory, &IID_IDWriteFactory2, (void**)&g_dwriteFactory2);
    }
    if (!g_d2dTarget) {
        D2D1_RENDER_TARGET_PROPERTIES props;
        props.type = D2D1_RENDER_TARGET_TYPE_DEFAULT;
        props.pixelFormat.format = DXGI_FORMAT_B8G8R8A8_UNORM;
        props.pixelFormat.alphaMode = D2D1_ALPHA_MODE_PREMULTIPLIED;
        props.dpiX = 96.0f;
        props.dpiY = 96.0f;
        props.usage = D2D1_RENDER_TARGET_USAGE_NONE;
        props.minLevel = D2D1_FEATURE_LEVEL_DEFAULT;
        hr = ID2D1Factory_CreateDCRenderTarget(g_d2dFactory, &props, &g_d2dTarget);
        if (FAILED(hr)) return FALSE;
    }
    if (!g_d2dTextBrush && g_d2dTarget) {
        ID2D1RenderTarget* rt = (ID2D1RenderTarget*)g_d2dTarget;
        D2D1_COLOR_F color = {0, 0, 0, 1.0f};
        hr = ID2D1RenderTarget_CreateSolidColorBrush(rt, &color, NULL, &g_d2dTextBrush);
        if (FAILED(hr)) return FALSE;
    }
    return TRUE;
}

static void ReleaseDWriteResources() {
    if (g_dwriteEmojiFormatColor) { IUnknown_Release((IUnknown*)g_dwriteEmojiFormatColor); g_dwriteEmojiFormatColor = NULL; }
    if (g_dwriteEmojiFormatMono) { IUnknown_Release((IUnknown*)g_dwriteEmojiFormatMono); g_dwriteEmojiFormatMono = NULL; }
    if (g_d2dTextBrush) { IUnknown_Release((IUnknown*)g_d2dTextBrush); g_d2dTextBrush = NULL; }
    if (g_d2dTarget) { IUnknown_Release((IUnknown*)g_d2dTarget); g_d2dTarget = NULL; }
    if (g_dwriteFactory2) { IUnknown_Release((IUnknown*)g_dwriteFactory2); g_dwriteFactory2 = NULL; }
    if (g_dwriteFactory) { IUnknown_Release((IUnknown*)g_dwriteFactory); g_dwriteFactory = NULL; }
    if (g_d2dFactory) { IUnknown_Release((IUnknown*)g_d2dFactory); g_d2dFactory = NULL; }
}

static void UpdateEmojiTextFormats() {
    if (!EnsureDWrite()) return;
    if (g_dwriteEmojiFormatColor) { IUnknown_Release((IUnknown*)g_dwriteEmojiFormatColor); g_dwriteEmojiFormatColor = NULL; }
    if (g_dwriteEmojiFormatMono) { IUnknown_Release((IUnknown*)g_dwriteEmojiFormatMono); g_dwriteEmojiFormatMono = NULL; }
    float size = (float)((g_effectiveFontSize * g_emojiScale + 2) * g_dpiScale);
    HRESULT hr = IDWriteFactory_CreateTextFormat(
        g_dwriteFactory,
        g_cfg.emojiFontColor,
        NULL,
        DWRITE_FONT_WEIGHT_NORMAL,
        DWRITE_FONT_STYLE_NORMAL,
        DWRITE_FONT_STRETCH_NORMAL,
        size,
        L"en-us",
        &g_dwriteEmojiFormatColor);
    if (SUCCEEDED(hr) && g_dwriteEmojiFormatColor) {
        IDWriteTextFormat_SetTextAlignment(g_dwriteEmojiFormatColor, DWRITE_TEXT_ALIGNMENT_CENTER);
        IDWriteTextFormat_SetParagraphAlignment(g_dwriteEmojiFormatColor, DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
        IDWriteTextFormat_SetWordWrapping(g_dwriteEmojiFormatColor, DWRITE_WORD_WRAPPING_NO_WRAP);
    }
    hr = IDWriteFactory_CreateTextFormat(
        g_dwriteFactory,
        g_cfg.emojiFontMono,
        NULL,
        DWRITE_FONT_WEIGHT_NORMAL,
        DWRITE_FONT_STYLE_NORMAL,
        DWRITE_FONT_STRETCH_NORMAL,
        size,
        L"en-us",
        &g_dwriteEmojiFormatMono);
    if (SUCCEEDED(hr) && g_dwriteEmojiFormatMono) {
        IDWriteTextFormat_SetTextAlignment(g_dwriteEmojiFormatMono, DWRITE_TEXT_ALIGNMENT_CENTER);
        IDWriteTextFormat_SetParagraphAlignment(g_dwriteEmojiFormatMono, DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
        IDWriteTextFormat_SetWordWrapping(g_dwriteEmojiFormatMono, DWRITE_WORD_WRAPPING_NO_WRAP);
    }
}

typedef struct {
    RECT rc;
    wchar_t emoji[8];
    COLORREF color;
} EmojiDraw;

typedef struct {
    IDWriteTextRenderer iface;
    ULONG refCount;
    ID2D1RenderTarget* rt;
    ID2D1SolidColorBrush* brush;
    IDWriteFactory2* factory2;
    BOOL enableColor;
    COLORREF monoColor;
} EmojiTextRenderer;

static HRESULT STDMETHODCALLTYPE EmojiTextRenderer_QueryInterface(IDWriteTextRenderer* This, REFIID riid, void** ppvObject) {
    if (!ppvObject) return E_POINTER;
    *ppvObject = NULL;
    if (IsEqualIID(riid, &IID_IUnknown) ||
        IsEqualIID(riid, &IID_IDWriteTextRenderer) ||
        IsEqualIID(riid, &IID_IDWritePixelSnapping)) {
        *ppvObject = This;
        IUnknown_AddRef((IUnknown*)This);
        return S_OK;
    }
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE EmojiTextRenderer_AddRef(IDWriteTextRenderer* This) {
    EmojiTextRenderer* self = (EmojiTextRenderer*)This;
    return ++self->refCount;
}

static ULONG STDMETHODCALLTYPE EmojiTextRenderer_Release(IDWriteTextRenderer* This) {
    EmojiTextRenderer* self = (EmojiTextRenderer*)This;
    if (self->refCount > 0) self->refCount--;
    return self->refCount;
}

static HRESULT STDMETHODCALLTYPE EmojiTextRenderer_IsPixelSnappingDisabled(
    IDWriteTextRenderer* This, void* context, BOOL* isDisabled) {
    (void)This; (void)context;
    if (!isDisabled) return E_POINTER;
    *isDisabled = FALSE;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE EmojiTextRenderer_GetCurrentTransform(
    IDWriteTextRenderer* This, void* context, DWRITE_MATRIX* transform) {
    (void)This; (void)context;
    if (!transform) return E_POINTER;
    transform->m11 = 1.0f; transform->m12 = 0.0f;
    transform->m21 = 0.0f; transform->m22 = 1.0f;
    transform->dx = 0.0f; transform->dy = 0.0f;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE EmojiTextRenderer_GetPixelsPerDip(
    IDWriteTextRenderer* This, void* context, FLOAT* pixelsPerDip) {
    (void)This; (void)context;
    if (!pixelsPerDip) return E_POINTER;
    *pixelsPerDip = 1.0f;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE EmojiTextRenderer_DrawGlyphRun(
    IDWriteTextRenderer* This,
    void* context,
    FLOAT baselineOriginX,
    FLOAT baselineOriginY,
    DWRITE_MEASURING_MODE measuringMode,
    const DWRITE_GLYPH_RUN* glyphRun,
    const DWRITE_GLYPH_RUN_DESCRIPTION* glyphRunDescription,
    IUnknown* clientDrawingEffect) {
    (void)context; (void)clientDrawingEffect;
    EmojiTextRenderer* self = (EmojiTextRenderer*)This;
    if (!self->rt || !self->brush || !glyphRun) return E_FAIL;

    if (self->enableColor && self->factory2) {
        IDWriteColorGlyphRunEnumerator* colorLayers = NULL;
        HRESULT hr = IDWriteFactory2_TranslateColorGlyphRun(
            self->factory2,
            baselineOriginX,
            baselineOriginY,
            glyphRun,
            glyphRunDescription,
            measuringMode,
            NULL,
            0,
            &colorLayers);
        if (SUCCEEDED(hr) && colorLayers) {
            BOOL hasRun = FALSE;
            while (SUCCEEDED(IDWriteColorGlyphRunEnumerator_MoveNext(colorLayers, &hasRun)) && hasRun) {
                const DWRITE_COLOR_GLYPH_RUN* colorRun = NULL;
                if (SUCCEEDED(IDWriteColorGlyphRunEnumerator_GetCurrentRun(colorLayers, &colorRun)) && colorRun) {
                    ID2D1SolidColorBrush_SetColor(self->brush, &colorRun->runColor);
                    D2D1_POINT_2F pt = {colorRun->baselineOriginX, colorRun->baselineOriginY};
                    ID2D1RenderTarget_DrawGlyphRun(
                        self->rt,
                        pt,
                        &colorRun->glyphRun,
                        (ID2D1Brush*)self->brush,
                        measuringMode);
                }
            }
            IUnknown_Release((IUnknown*)colorLayers);
            return S_OK;
        }
    }

    D2D1_COLOR_F mono = {
        GetRValue(self->monoColor) / 255.0f,
        GetGValue(self->monoColor) / 255.0f,
        GetBValue(self->monoColor) / 255.0f,
        1.0f
    };
    ID2D1SolidColorBrush_SetColor(self->brush, &mono);
    D2D1_POINT_2F pt = {baselineOriginX, baselineOriginY};
    ID2D1RenderTarget_DrawGlyphRun(
        self->rt,
        pt,
        glyphRun,
        (ID2D1Brush*)self->brush,
        measuringMode);
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE EmojiTextRenderer_DrawUnderline(
    IDWriteTextRenderer* This,
    void* context,
    FLOAT baselineOriginX,
    FLOAT baselineOriginY,
    const DWRITE_UNDERLINE* underline,
    IUnknown* clientDrawingEffect) {
    (void)This; (void)context; (void)baselineOriginX; (void)baselineOriginY; (void)underline; (void)clientDrawingEffect;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE EmojiTextRenderer_DrawStrikethrough(
    IDWriteTextRenderer* This,
    void* context,
    FLOAT baselineOriginX,
    FLOAT baselineOriginY,
    const DWRITE_STRIKETHROUGH* strikethrough,
    IUnknown* clientDrawingEffect) {
    (void)This; (void)context; (void)baselineOriginX; (void)baselineOriginY; (void)strikethrough; (void)clientDrawingEffect;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE EmojiTextRenderer_DrawInlineObject(
    IDWriteTextRenderer* This,
    void* context,
    FLOAT originX,
    FLOAT originY,
    IDWriteInlineObject* inlineObject,
    BOOL isSideways,
    BOOL isRightToLeft,
    IUnknown* clientDrawingEffect) {
    (void)This; (void)context; (void)originX; (void)originY; (void)inlineObject; (void)isSideways; (void)isRightToLeft; (void)clientDrawingEffect;
    return S_OK;
}

static IDWriteTextRendererVtbl g_emojiTextRendererVtbl = {
    EmojiTextRenderer_QueryInterface,
    EmojiTextRenderer_AddRef,
    EmojiTextRenderer_Release,
    EmojiTextRenderer_IsPixelSnappingDisabled,
    EmojiTextRenderer_GetCurrentTransform,
    EmojiTextRenderer_GetPixelsPerDip,
    EmojiTextRenderer_DrawGlyphRun,
    EmojiTextRenderer_DrawUnderline,
    EmojiTextRenderer_DrawStrikethrough,
    EmojiTextRenderer_DrawInlineObject
};

static BOOL DrawEmojiWithDWrite(HDC hdc, const EmojiDraw* items, int count, COLORREF monoColor, BOOL mono) {
    if (count <= 0) return FALSE;
    if (!EnsureDWrite()) return FALSE;
    IDWriteTextFormat* fmt = mono ? g_dwriteEmojiFormatMono : g_dwriteEmojiFormatColor;
    if (!fmt || !g_d2dTarget || !g_d2dTextBrush) return FALSE;
    RECT rc = {0, 0, g_cfg.width, g_cfg.height};
    HRESULT hr = ID2D1DCRenderTarget_BindDC(g_d2dTarget, hdc, &rc);
    if (FAILED(hr)) return FALSE;
    EmojiTextRenderer renderer;
    ZeroMemory(&renderer, sizeof(renderer));
    renderer.iface.lpVtbl = &g_emojiTextRendererVtbl;
    renderer.refCount = 1;
    renderer.rt = (ID2D1RenderTarget*)g_d2dTarget;
    renderer.brush = g_d2dTextBrush;
    renderer.factory2 = g_dwriteFactory2;
    renderer.enableColor = (!mono && g_dwriteFactory2 != NULL);
    renderer.monoColor = monoColor;

    ID2D1RenderTarget_BeginDraw((ID2D1RenderTarget*)g_d2dTarget);
    for (int i = 0; i < count; i++) {
        UINT32 len = (UINT32)wcslen(items[i].emoji);
        if (len == 0) continue;
        IDWriteTextLayout* layout = NULL;
        FLOAT w = (FLOAT)(items[i].rc.right - items[i].rc.left);
        FLOAT h = (FLOAT)(items[i].rc.bottom - items[i].rc.top);
        if (g_dwriteFactory2) {
            IDWriteFactory2_CreateTextLayout(g_dwriteFactory2, items[i].emoji, len, fmt, w, h, &layout);
        } else {
            IDWriteFactory_CreateTextLayout(g_dwriteFactory, items[i].emoji, len, fmt, w, h, &layout);
        }
            if (layout) {
                IDWriteTextLayout_SetTextAlignment(layout, DWRITE_TEXT_ALIGNMENT_CENTER);
                IDWriteTextLayout_SetParagraphAlignment(layout, DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
                renderer.monoColor = items[i].color;
            IDWriteTextLayout_Draw(layout, NULL, (IDWriteTextRenderer*)&renderer,
                (FLOAT)items[i].rc.left, (FLOAT)items[i].rc.top);
                IUnknown_Release((IUnknown*)layout);
            }
    }
    HRESULT endHr = ID2D1RenderTarget_EndDraw((ID2D1RenderTarget*)g_d2dTarget, NULL, NULL);
    if (endHr == D2DERR_RECREATE_TARGET) {
        if (g_d2dTextBrush) { IUnknown_Release((IUnknown*)g_d2dTextBrush); g_d2dTextBrush = NULL; }
        if (g_d2dTarget) { IUnknown_Release((IUnknown*)g_d2dTarget); g_d2dTarget = NULL; }
    }
    return TRUE;
}

static BOOL DrawTextWithDWrite(HDC hdc, const RECT* rc, const wchar_t* text, COLORREF color, float sizePx) {
    if (!text || !rc) return FALSE;
    if (!EnsureDWrite()) return FALSE;
    if (!g_d2dTarget || !g_d2dTextBrush) return FALSE;
    RECT bind = {0, 0, g_cfg.width, g_cfg.height};
    HRESULT hr = ID2D1DCRenderTarget_BindDC(g_d2dTarget, hdc, &bind);
    if (FAILED(hr)) return FALSE;

    IDWriteTextFormat* fmt = NULL;
    if (g_dwriteFactory2) {
        IDWriteFactory2_CreateTextFormat(
            g_dwriteFactory2, g_cfg.fontName, NULL,
            DWRITE_FONT_WEIGHT_SEMI_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL,
            sizePx, L"vi-VN", &fmt);
    } else {
        IDWriteFactory_CreateTextFormat(
            g_dwriteFactory, g_cfg.fontName, NULL,
            DWRITE_FONT_WEIGHT_SEMI_BOLD, DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL,
            sizePx, L"vi-VN", &fmt);
    }
    if (!fmt) return FALSE;

    IDWriteTextFormat_SetTextAlignment(fmt, DWRITE_TEXT_ALIGNMENT_CENTER);
    IDWriteTextFormat_SetParagraphAlignment(fmt, DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
    IDWriteTextFormat_SetWordWrapping(fmt, DWRITE_WORD_WRAPPING_NO_WRAP);

    EmojiTextRenderer renderer;
    ZeroMemory(&renderer, sizeof(renderer));
    renderer.iface.lpVtbl = &g_emojiTextRendererVtbl;
    renderer.refCount = 1;
    renderer.rt = (ID2D1RenderTarget*)g_d2dTarget;
    renderer.brush = g_d2dTextBrush;
    renderer.factory2 = g_dwriteFactory2;
    renderer.enableColor = (g_dwriteFactory2 != NULL);
    renderer.monoColor = color;

    ID2D1RenderTarget_BeginDraw((ID2D1RenderTarget*)g_d2dTarget);
    IDWriteTextLayout* layout = NULL;
    UINT32 len = (UINT32)wcslen(text);
    float w = (float)(rc->right - rc->left);
    float h = (float)(rc->bottom - rc->top);
    if (g_dwriteFactory2) {
        IDWriteFactory2_CreateTextLayout(g_dwriteFactory2, text, len, fmt, w, h, &layout);
    } else {
        IDWriteFactory_CreateTextLayout(g_dwriteFactory, text, len, fmt, w, h, &layout);
    }
    if (layout) {
        IDWriteTextLayout_SetTextAlignment(layout, DWRITE_TEXT_ALIGNMENT_CENTER);
        IDWriteTextLayout_SetParagraphAlignment(layout, DWRITE_PARAGRAPH_ALIGNMENT_CENTER);
        IDWriteTextLayout_Draw(layout, NULL, (IDWriteTextRenderer*)&renderer,
            (FLOAT)rc->left, (FLOAT)rc->top);
        IUnknown_Release((IUnknown*)layout);
    }
    IUnknown_Release((IUnknown*)fmt);

    HRESULT endHr = ID2D1RenderTarget_EndDraw((ID2D1RenderTarget*)g_d2dTarget, NULL, NULL);
    if (endHr == D2DERR_RECREATE_TARGET) {
        if (g_d2dTextBrush) { IUnknown_Release((IUnknown*)g_d2dTextBrush); g_d2dTextBrush = NULL; }
        if (g_d2dTarget) { IUnknown_Release((IUnknown*)g_d2dTarget); g_d2dTarget = NULL; }
    }
    return TRUE;
}

static void NormalizeMonthYear(int* mm, int* yy) {
    while (*mm < 1) { *mm += 12; (*yy)--; }
    while (*mm > 12) { *mm -= 12; (*yy)++; }
}

static void SetViewMonthYear(int mm, int yy, BOOL followToday) {
    NormalizeMonthYear(&mm, &yy);
    g_viewMonth = mm;
    g_viewYear = yy;
    g_followToday = followToday;
    g_hoverDay = -1;
    if (g_hTooltip) SendMessageW(g_hTooltip, TTM_POP, 0, 0);
    CalculateMetrics();
    if (g_hwnd) {
        ClampPositionToVisibleMonitor();
        SetWindowPos(g_hwnd, NULL, g_cfg.x, g_cfg.y, g_cfg.width, g_cfg.height, SWP_NOZORDER | SWP_NOACTIVATE);
        UpdateRelativePositionFromCurrent();
    }
}

static void SetViewToToday() {
    SYSTEMTIME st;
    GetLocalTime(&st);
    SetViewMonthYear(st.wMonth, st.wYear, TRUE);
}

static void UpdateTooltip(int day) {
    if (!g_hTooltip) return;
    if (day < 1 || day > 31) {
        SendMessageW(g_hTooltip, TTM_POP, 0, 0);
        return;
    }
    if (day > daysInMonth(g_viewMonth, g_viewYear)) return;
    LunarDate lunar = SolarToLunar(day, g_viewMonth, g_viewYear);
    wchar_t canChiDay[32], canChiMonth[32], canChiYear[32];
    GetCanChiDay(day, g_viewMonth, g_viewYear, canChiDay, 32);
    GetCanChiMonth(lunar.month, lunar.year, canChiMonth, 32);
    GetCanChiYear(lunar.year, canChiYear, 32);
    int dow = dayOfWeek(day, g_viewMonth, g_viewYear);
    wchar_t tip[256];
    swprintf(tip, 256,
        L"Ng\u00E0y %s, th\u00E1ng %s, n\u0103m %s\nD\u01B0\u1EDBng l\u1ECBch: %02d/%02d/%d (%s)\n\u00C2m l\u1ECBch: %02d/%02d/%d%s",
        canChiDay, canChiMonth, canChiYear,
        day, g_viewMonth, g_viewYear, kDayNames[dow],
        lunar.day, lunar.month, lunar.year,
        lunar.leap ? L" (Nhu\u1EADn)" : L"");
    TOOLINFOW ti = {0};
    ti.cbSize = sizeof(ti);
    ti.uFlags = TTF_SUBCLASS;
    ti.hwnd = g_hwnd;
    ti.uId = 1;
    ti.lpszText = tip;
    GetClientRect(g_hwnd, &ti.rect);
    SendMessageW(g_hTooltip, TTM_UPDATETIPTEXTW, 0, (LPARAM)&ti);
    SendMessageW(g_hTooltip, TTM_POPUP, 0, 0);
}

static int GetDayFromPoint(int x, int y) {
    int gridWidth = g_cellWidth * 7;
    int startX = g_cardRect.left + (g_cardWidth - gridWidth) / 2;
    int startY = g_cardRect.top + g_topMargin + g_headerHeight + g_dayHeaderHeight;
    if (x < startX || y < startY) return -1;
    int col = (x - startX) / g_cellWidth;
    int row = (y - startY) / g_cellHeight;
    if (col < 0 || col > 6 || row < 0 || row > (g_visibleRows - 1)) return -1;
    int firstDay = dayOfWeek(1, g_viewMonth, g_viewYear);
    int day = row * 7 + col - firstDay + 1;
    if (day < 1 || day > daysInMonth(g_viewMonth, g_viewYear)) return -1;
    return day;
}

static void DrawCalendar(HWND hwnd) {
    HDC hdcScreen = GetDC(NULL);
    HDC hdcMem = CreateCompatibleDC(hdcScreen);
    BITMAPINFO bmi = {0};
    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
    bmi.bmiHeader.biWidth = g_cfg.width;
    bmi.bmiHeader.biHeight = -g_cfg.height;
    bmi.bmiHeader.biPlanes = 1;
    bmi.bmiHeader.biBitCount = 32;
    bmi.bmiHeader.biCompression = BI_RGB;
    void* bits = NULL;
    HBITMAP hbm = CreateDIBSection(hdcMem, &bmi, DIB_RGB_COLORS, &bits, NULL, 0);
    HBITMAP oldBmp = (HBITMAP)SelectObject(hdcMem, hbm);
    ZeroMemory(bits, g_cfg.width * g_cfg.height * 4);

    DWORD* pixel = (DWORD*)bits;
    COLORREF bg = g_cfg.bgColor;
    COLORREF bgTop = BlendColor(bg, RGB(255,255,255), 0.06f);
    COLORREF bgBottom = BlendColor(bg, RGB(0,0,0), 0.02f);
    COLORREF border = BlendColor(bg, RGB(0,0,0), 0.15f);
    COLORREF grid = BlendColor(bg, RGB(0,0,0), 0.08f);
    float headerShine = 0.04f;
    if (g_hoverHeader) {
        float pulse = 0.5f + 0.5f * sinf(g_headerPulse);
        headerShine = 0.06f + 0.10f * pulse;
        if (headerShine > 0.22f) headerShine = 0.22f;
    }
    COLORREF headerBg = BlendColor(bg, RGB(255,255,255), headerShine);
    COLORREF accent = g_cfg.highlightColor;
    COLORREF todayBg = BlendColor(accent, RGB(255,255,255), 0.78f);
    COLORREF hoverBg = BlendColor(bg, RGB(0,0,0), 0.04f);
    COLORREF tetBg = (Luminance(bg) < 90)
        ? BlendColor(accent, bg, 0.60f)
        : RGB(255, 236, 214);
    COLORREF sundayColor = RGB(220, 38, 38);
    COLORREF saturdayColor = RGB(22, 163, 74);
    COLORREF dayNameColor = BlendColor(g_cfg.textColor, bg, 0.45f);
    COLORREF todayText = PickTextColorOn(todayBg, RGB(18, 18, 18), RGB(245, 245, 245));

    int radius = (int)(16 * g_dpiScale);
    if (radius < 8) radius = 8;
    int shadow = g_shadowSize;
    BYTE cardAlpha = 255;
    BYTE shadowMax = (BYTE)min(140, g_cfg.alpha / 2 + 30);

    for (int py = 0; py < g_cfg.height; py++) {
        for (int px = 0; px < g_cfg.width; px++) {
            float dist = DistanceToRoundedRect(px, py, g_cardRect, radius);
            if (dist <= 0.0f) {
                float t = 0.0f;
                if (g_cardRect.bottom > g_cardRect.top + 1) {
                    t = (float)(py - g_cardRect.top) / (float)(g_cardRect.bottom - g_cardRect.top - 1);
                }
                COLORREF c = BlendColor(bgTop, bgBottom, t);
                pixel[py * g_cfg.width + px] = PremultiplyColor(c, cardAlpha);
            } else if (dist <= shadow) {
                float k = 1.0f - dist / (float)shadow;
                if (k < 0.0f) k = 0.0f;
                BYTE a = (BYTE)(shadowMax * k * k);
                pixel[py * g_cfg.width + px] = PremultiplyColor(RGB(0,0,0), a);
            }
        }
    }

    SetBkMode(hdcMem, TRANSPARENT);
    RECT rcHeader = {g_cardRect.left, g_cardRect.top + g_topMargin, g_cardRect.right, g_cardRect.top + g_topMargin + g_headerHeight};
    g_rcHeader.left = g_cardRect.left;
    g_rcHeader.top = g_cardRect.top;
    g_rcHeader.right = g_cardRect.right;
    g_rcHeader.bottom = rcHeader.bottom;
    HBRUSH headerBrush = CreateSolidBrush(headerBg);
    FillRect(hdcMem, &rcHeader, headerBrush);
    DeleteObject(headerBrush);

    SYSTEMTIME st;
    GetLocalTime(&st);
    BOOL isViewingCurrent = (g_viewMonth == st.wMonth && g_viewYear == st.wYear);
    BOOL isCurrentMonth = isViewingCurrent;

    wchar_t title[160];
    if (isViewingCurrent) {
        LunarDate lunarToday = SolarToLunar(st.wDay, st.wMonth, st.wYear);
        int chiDayTitle = GetChiIndexDay(st.wDay, st.wMonth, st.wYear);
        int chiMonthTitle = (lunarToday.month + 1) % 12;
        int chiYearTitle = (lunarToday.year + 8) % 12;
        const wchar_t* emojiDay = kChiEmoji[chiDayTitle];
        const wchar_t* emojiMonth = kChiEmoji[chiMonthTitle];
        const wchar_t* emojiYear = kChiEmoji[chiYearTitle];
        swprintf(title, 160, L"Ng\u00E0y %02d %s th\u00E1ng %02d %s n\u0103m %d %s",
            st.wDay, emojiDay, st.wMonth, emojiMonth, st.wYear, emojiYear);
    } else {
        LunarDate lunarTitle = SolarToLunar(1, g_viewMonth, g_viewYear);
        int chiMonthTitle = (lunarTitle.month + 1) % 12;
        int chiYearTitle = (lunarTitle.year + 8) % 12;
        const wchar_t* emojiMonth = kChiEmoji[chiMonthTitle];
        const wchar_t* emojiYear = kChiEmoji[chiYearTitle];
        swprintf(title, 160, L"Th\u00E1ng %02d %s n\u0103m %d %s", g_viewMonth, emojiMonth, g_viewYear, emojiYear);
    }
    SetTextColor(hdcMem, g_cfg.textColor);
    HFONT oldFont = (HFONT)SelectObject(hdcMem, g_fontMonth);
    RECT rcTitle = rcHeader;
    rcTitle.left += g_sideMargin + (int)(24 * g_dpiScale);
    rcTitle.right -= g_sideMargin + (int)(24 * g_dpiScale);
    rcTitle.bottom = rcHeader.top + (int)(g_headerHeight * 0.6f);
    if (!DrawTextWithDWrite(hdcMem, &rcTitle, title, g_cfg.textColor, (float)((g_effectiveFontSize + 2) * g_dpiScale))) {
        DrawTextW(hdcMem, title, -1, &rcTitle, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }

    wchar_t sub[64];
    if (isViewingCurrent) {
        LunarDate todayLunar = SolarToLunar(st.wDay, st.wMonth, st.wYear);
        swprintf(sub, 64, L"\u00C2m l\u1ECBch h\u00F4m nay: %02d/%02d", todayLunar.day, todayLunar.month);
    } else {
        swprintf(sub, 64, L"\u0110ang xem: %02d/%d", g_viewMonth, g_viewYear);
    }
    SelectObject(hdcMem, g_fontSmall);
    RECT rcSub = rcHeader;
    rcSub.top = rcTitle.bottom - (int)(4 * g_dpiScale);
    rcSub.left += g_sideMargin;
    rcSub.right -= g_sideMargin;
    SetTextColor(hdcMem, BlendColor(g_cfg.textColor, bg, 0.35f));
    DrawTextW(hdcMem, sub, -1, &rcSub, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    int navSize = (int)(26 * g_dpiScale);
    int navY = rcHeader.top + (g_headerHeight - navSize) / 2;
    g_rcPrev.left = g_cardRect.left + g_sideMargin / 2;
    g_rcPrev.top = navY;
    g_rcPrev.right = g_rcPrev.left + navSize;
    g_rcPrev.bottom = navY + navSize;
    g_rcNext.right = g_cardRect.right - g_sideMargin / 2;
    g_rcNext.top = navY;
    g_rcNext.left = g_rcNext.right - navSize;
    g_rcNext.bottom = navY + navSize;

    // Navigation buttons are emoji only (no rounded background)

    // Emoji arrows for navigation
    {
        RECT rPrevEmoji = g_rcPrev;
        RECT rNextEmoji = g_rcNext;
        int basePad = (int)(4 * g_dpiScale);
        int padPrev = (g_hoverNav == 0) ? (int)(1 * g_dpiScale) : basePad;
        int padNext = (g_hoverNav == 1) ? (int)(1 * g_dpiScale) : basePad;
        InflateRect(&rPrevEmoji, -padPrev, -padPrev);
        InflateRect(&rNextEmoji, -padNext, -padNext);
        const wchar_t* prevEmoji = L"\u2B05\uFE0F";
        const wchar_t* nextEmoji = L"\u27A1\uFE0F";
        BOOL useDwriteEmojiNav = EnsureDWrite() && (g_dwriteEmojiFormatColor || g_dwriteEmojiFormatMono);
        if (useDwriteEmojiNav) {
            EmojiDraw navItems[2];
            ZeroMemory(navItems, sizeof(navItems));
            navItems[0].rc = rPrevEmoji;
            wcsncpy(navItems[0].emoji, prevEmoji, 7);
            navItems[0].emoji[7] = 0;
            navItems[0].color = g_cfg.textColor;
            navItems[1].rc = rNextEmoji;
            wcsncpy(navItems[1].emoji, nextEmoji, 7);
            navItems[1].emoji[7] = 0;
            navItems[1].color = g_cfg.textColor;
            DrawEmojiWithDWrite(hdcMem, navItems, 2, g_cfg.textColor, g_cfg.emojiMono);
        } else {
            SelectObject(hdcMem, g_fontEmoji);
            SetTextColor(hdcMem, g_cfg.textColor);
            DrawTextW(hdcMem, prevEmoji, -1, &rPrevEmoji, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            DrawTextW(hdcMem, nextEmoji, -1, &rNextEmoji, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        }
    }

    HPEN hSepPen = CreatePen(PS_SOLID, 1, grid);
    HPEN oldPen = (HPEN)SelectObject(hdcMem, hSepPen);
    MoveToEx(hdcMem, g_cardRect.left + g_sideMargin, rcHeader.bottom, NULL);
    LineTo(hdcMem, g_cardRect.right - g_sideMargin, rcHeader.bottom);
    SelectObject(hdcMem, oldPen);
    DeleteObject(hSepPen);

    int dayHeaderTop = rcHeader.bottom;
    int dayHeaderBottom = dayHeaderTop + g_dayHeaderHeight;
    g_rcHeader.bottom = dayHeaderBottom;
    int gridWidth = g_cellWidth * 7;
    int startX = g_cardRect.left + (g_cardWidth - gridWidth) / 2;
    int startY = dayHeaderBottom;

    SelectObject(hdcMem, g_fontSmall);
    for (int i = 0; i < 7; i++) {
        RECT r = {startX + i * g_cellWidth, dayHeaderTop, startX + (i + 1) * g_cellWidth, dayHeaderBottom};
        COLORREF nameColor = dayNameColor;
        if (i == 0) nameColor = sundayColor;
        else if (i == 6) nameColor = saturdayColor;
        SetTextColor(hdcMem, nameColor);
        DrawTextW(hdcMem, kDayNames[i], -1, &r, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    }

    HBRUSH brToday = CreateSolidBrush(todayBg);
    HBRUSH brHover = CreateSolidBrush(hoverBg);
    HBRUSH brTet = CreateSolidBrush(tetBg);

    int days = daysInMonth(g_viewMonth, g_viewYear);
    int firstDay = dayOfWeek(1, g_viewMonth, g_viewYear);
    int cellPadX = (int)(6 * g_dpiScale);
    int cellPadY = (int)(4 * g_dpiScale);
    int bigSizePx = (int)(g_effectiveFontSize * g_solarScale * g_dpiScale);
    int lunarSizePx = (int)((g_effectiveFontSize - 2) * g_lunarScale * g_dpiScale);
    int emojiSizePx = (int)((g_effectiveFontSize * g_emojiScale + 1) * g_dpiScale);
    int bandPad = (int)(4 * g_dpiScale);
    int topBand = bigSizePx + bandPad * 2;
    int bottomBand = g_showLunar ? (lunarSizePx + bandPad * 2) : (bandPad * 2);
    int emojiBand = emojiSizePx + bandPad * 2;
    int minEmojiBand = (int)(16 * g_dpiScale);
    if (emojiBand < minEmojiBand) emojiBand = minEmojiBand;
    EmojiDraw emojiItems[31];
    int emojiCount = 0;
    BOOL useDwriteEmoji = EnsureDWrite() && (g_dwriteEmojiFormatColor || g_dwriteEmojiFormatMono);

    for (int d = 1; d <= days; d++) {
        int index = firstDay + d - 1;
        int col = index % 7;
        int row = index / 7;
        int x = startX + col * g_cellWidth;
        int y = startY + row * g_cellHeight;
        RECT cell = {x, y, x + g_cellWidth, y + g_cellHeight};
        LunarDate ld = SolarToLunar(d, g_viewMonth, g_viewYear);
        BOOL isToday = isCurrentMonth && (d == st.wDay);
        BOOL isTet = (ld.day == 1 && ld.month == 1);
        BOOL isHover = (g_hoverDay == d);

        if (isToday) FillRect(hdcMem, &cell, brToday);
        else if (isTet) FillRect(hdcMem, &cell, brTet);
        else if (isHover) FillRect(hdcMem, &cell, brHover);

        COLORREF solarColor = g_cfg.textColor;
        if (col == 0) solarColor = sundayColor;
        else if (col == 6) solarColor = saturdayColor;
        if (isToday) solarColor = todayText;
        if (isTet && !isToday) solarColor = PickTextColorOn(tetBg, RGB(24, 24, 24), RGB(245, 245, 245));

        SelectObject(hdcMem, g_fontBig);
        SetTextColor(hdcMem, solarColor);
        int solarTop = y + cellPadY;
        int solarBottom = y + topBand - cellPadY;
        int emojiTop = y + topBand;
        int emojiBottom = emojiTop + emojiBand;
        int lunarTop = emojiBottom;
        int lunarBottom = y + g_cellHeight - cellPadY;

        RECT rSolar = {x + cellPadX, solarTop, x + g_cellWidth - cellPadX, solarBottom};
        wchar_t buf[16];
        swprintf(buf, 16, L"%d", d);
        DrawTextW(hdcMem, buf, -1, &rSolar, DT_LEFT | DT_TOP | DT_SINGLELINE);

        COLORREF lunarColor = g_cfg.lunarColor;
        if (ld.day == 1) lunarColor = accent;
        if (isToday) lunarColor = todayText;
        if (isTet && !isToday) lunarColor = PickTextColorOn(tetBg, RGB(24, 24, 24), RGB(245, 245, 245));
        SelectObject(hdcMem, g_fontLunar);
        SetTextColor(hdcMem, lunarColor);
        if (g_showLunar) {
            RECT rLunar = {x + cellPadX, lunarTop, x + g_cellWidth - cellPadX, lunarBottom};
            if (ld.day == 1) swprintf(buf, 16, L"%d/%d", ld.day, ld.month);
            else swprintf(buf, 16, L"%d", ld.day);
            DrawTextW(hdcMem, buf, -1, &rLunar, DT_RIGHT | DT_BOTTOM | DT_SINGLELINE);
        }

        // Zodiac emoji (center band)
        if (g_fontEmoji) {
            int chiIndex = GetChiIndexDay(d, g_viewMonth, g_viewYear);
            const wchar_t* emoji = kChiEmoji[chiIndex];
            RECT rEmoji = {x, emojiTop + (cellPadY / 2), x + g_cellWidth, emojiBottom - (cellPadY / 2)};
            COLORREF emojiColor = solarColor;
            if (isToday) emojiColor = todayText;
            if (isTet && !isToday) emojiColor = PickTextColorOn(tetBg, RGB(24, 24, 24), RGB(245, 245, 245));
            if (useDwriteEmoji && emojiCount < 31) {
                emojiItems[emojiCount].rc = rEmoji;
                wcsncpy(emojiItems[emojiCount].emoji, emoji, 7);
                emojiItems[emojiCount].emoji[7] = 0;
                emojiItems[emojiCount].color = emojiColor;
                emojiCount++;
            } else {
                SelectObject(hdcMem, g_fontEmoji);
                SetTextColor(hdcMem, emojiColor);
                DrawTextW(hdcMem, emoji, -1, &rEmoji, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
            }
        }
    }

    if (useDwriteEmoji && emojiCount > 0) {
        DrawEmojiWithDWrite(hdcMem, emojiItems, emojiCount, g_cfg.textColor, g_cfg.emojiMono);
    }

    HPEN hGridPen = CreatePen(PS_SOLID, 1, grid);
    oldPen = (HPEN)SelectObject(hdcMem, hGridPen);
    for (int i = 0; i <= 7; i++) {
        int gx = startX + i * g_cellWidth;
        MoveToEx(hdcMem, gx, startY, NULL);
        LineTo(hdcMem, gx, startY + g_cellHeight * g_visibleRows);
    }
    for (int i = 0; i <= g_visibleRows; i++) {
        int gy = startY + i * g_cellHeight;
        MoveToEx(hdcMem, startX, gy, NULL);
        LineTo(hdcMem, startX + g_cellWidth * 7, gy);
    }
    SelectObject(hdcMem, oldPen);
    DeleteObject(hGridPen);

    HBRUSH oldBrush = (HBRUSH)SelectObject(hdcMem, GetStockObject(NULL_BRUSH));
    HPEN hBorderPen = CreatePen(PS_SOLID, 1, border);
    oldPen = (HPEN)SelectObject(hdcMem, hBorderPen);
    RoundRect(hdcMem, g_cardRect.left, g_cardRect.top, g_cardRect.right, g_cardRect.bottom, radius * 2, radius * 2);
    SelectObject(hdcMem, oldPen);
    SelectObject(hdcMem, oldBrush);
    DeleteObject(hBorderPen);

    DeleteObject(brToday);
    DeleteObject(brHover);
    DeleteObject(brTet);

    SelectObject(hdcMem, oldFont);
    RestoreCardAlpha(pixel, radius);
    POINT ptSrc = {0,0};
    POINT ptDst = {g_cfg.x, g_cfg.y};
    SIZE size = {g_cfg.width, g_cfg.height};
    BLENDFUNCTION blend = {AC_SRC_OVER, 0, (BYTE)g_cfg.alpha, AC_SRC_ALPHA};
    UpdateLayeredWindow(hwnd, hdcScreen, &ptDst, &size, hdcMem, &ptSrc, 0, &blend, ULW_ALPHA);
    SelectObject(hdcMem, oldBmp);
    DeleteObject(hbm);
    DeleteDC(hdcMem);
    ReleaseDC(NULL, hdcScreen);
}

static void PickFont(HWND parent) {
    CHOOSEFONTW cf = {0};
    LOGFONTW lf = {0};
    lf.lfHeight = -MulDiv(g_cfg.fontSize, (int)(96 * g_dpiScale), 72);
    lf.lfWeight = FW_NORMAL;
    lf.lfQuality = ANTIALIASED_QUALITY;
    wcscpy(lf.lfFaceName, g_cfg.fontName);
    cf.lStructSize = sizeof(cf);
    cf.hwndOwner = parent;
    cf.lpLogFont = &lf;
    cf.Flags = CF_FORCEFONTEXIST | CF_INITTOLOGFONTSTRUCT | CF_SCREENFONTS | CF_LIMITSIZE;
    cf.nSizeMin = 8;
    cf.nSizeMax = 24;
    cf.nFontType = REGULAR_FONTTYPE;
    if (ChooseFontW(&cf)) {
        wcscpy(g_cfg.fontName, lf.lfFaceName);
        HDC hdc = GetDC(NULL);
        int dpi = GetDeviceCaps(hdc, LOGPIXELSY);
        ReleaseDC(NULL, hdc);
        g_cfg.fontSize = (int)((-lf.lfHeight * 72 + dpi/2) / dpi);
        if (g_cfg.fontSize < 8) g_cfg.fontSize = 8;
        if (g_cfg.fontSize > 24) g_cfg.fontSize = 24;
    }
}

static void PickEmojiFont(HWND parent, BOOL mono) {
    CHOOSEFONTW cf = {0};
    LOGFONTW lf = {0};
    lf.lfHeight = -MulDiv(max(g_cfg.fontSize, 8), (int)(96 * g_dpiScale), 72);
    lf.lfWeight = FW_NORMAL;
    lf.lfQuality = mono ? ANTIALIASED_QUALITY : CLEARTYPE_QUALITY;
    wcscpy(lf.lfFaceName, mono ? g_cfg.emojiFontMono : g_cfg.emojiFontColor);
    cf.lStructSize = sizeof(cf);
    cf.hwndOwner = parent;
    cf.lpLogFont = &lf;
    cf.Flags = CF_FORCEFONTEXIST | CF_INITTOLOGFONTSTRUCT | CF_SCREENFONTS | CF_NOSIZESEL;
    cf.nFontType = REGULAR_FONTTYPE;
    if (ChooseFontW(&cf)) {
        if (mono) wcscpy(g_cfg.emojiFontMono, lf.lfFaceName);
        else wcscpy(g_cfg.emojiFontColor, lf.lfFaceName);
    }
}

static void UpdateEmojiFontLabel(HWND hwnd) {
    int sel = (int)SendMessage(GetDlgItem(hwnd, ID_EMOJI_STYLE), CB_GETCURSEL, 0, 0);
    BOOL mono = (sel == 1);
    const wchar_t* face = mono ? g_cfg.emojiFontMono : g_cfg.emojiFontColor;
    SetWindowTextW(GetDlgItem(hwnd, ID_LBL_EMOJI_FONT), face);
}

static void ApplyScaleFromUI(HWND hwnd, BOOL save) {
    int solar = (int)SendMessage(GetDlgItem(hwnd, ID_SLIDER_SOLAR), TBM_GETPOS, 0, 0);
    int emoji = (int)SendMessage(GetDlgItem(hwnd, ID_SLIDER_EMOJI), TBM_GETPOS, 0, 0);
    int lunar = (int)SendMessage(GetDlgItem(hwnd, ID_SLIDER_LUNAR), TBM_GETPOS, 0, 0);
    g_cfg.solarScalePct = solar;
    g_cfg.emojiScalePct = emoji;
    g_cfg.lunarScalePct = lunar;
    SyncScaleFromConfig();
    CalculateMetrics();
    CreateFonts();
    ClampPositionToVisibleMonitor();
    SetWindowPos(g_hwnd, NULL, g_cfg.x, g_cfg.y, g_cfg.width, g_cfg.height, SWP_NOZORDER | SWP_NOACTIVATE);
    UpdateRelativePositionFromCurrent();
    DrawCalendar(g_hwnd);
    if (save) SaveConfig();
}

static void ApplySettingsFromDialog(HWND hwnd) {
    wchar_t buf[16];
    GetWindowTextW(GetDlgItem(hwnd, ID_EDIT_X), buf, 16); g_cfg.x = _wtoi(buf);
    GetWindowTextW(GetDlgItem(hwnd, ID_EDIT_Y), buf, 16); g_cfg.y = _wtoi(buf);
    g_cfg.alpha = (int)SendMessage(GetDlgItem(hwnd, ID_OPACITY_TRACK), TBM_GETPOS, 0, 0);
    g_cfg.autoRun = (BST_CHECKED == SendMessage(GetDlgItem(hwnd, ID_CHK_AUTORUN), BM_GETCHECK, 0, 0));
    g_cfg.stickDesktop = (BST_CHECKED == SendMessage(GetDlgItem(hwnd, ID_CHK_STICK), BM_GETCHECK, 0, 0));
    g_cfg.primaryScreenOnly = (BST_CHECKED == SendMessage(GetDlgItem(hwnd, ID_CHK_PRIMARY), BM_GETCHECK, 0, 0));
    g_cfg.allowMove = (BST_CHECKED == SendMessage(GetDlgItem(hwnd, ID_CHK_MOVE), BM_GETCHECK, 0, 0));
    UpdateStickTimer(g_hwnd);
    UpdateClickThroughStyle();
    g_cfg.emojiMono = ((int)SendMessage(GetDlgItem(hwnd, ID_EMOJI_STYLE), CB_GETCURSEL, 0, 0) == 1);
    int solar = (int)SendMessage(GetDlgItem(hwnd, ID_SLIDER_SOLAR), TBM_GETPOS, 0, 0);
    int emoji = (int)SendMessage(GetDlgItem(hwnd, ID_SLIDER_EMOJI), TBM_GETPOS, 0, 0);
    int lunar = (int)SendMessage(GetDlgItem(hwnd, ID_SLIDER_LUNAR), TBM_GETPOS, 0, 0);
    solar = GetClampedEditValue(hwnd, ID_EDIT_SOLAR, 90, 180, solar);
    emoji = GetClampedEditValue(hwnd, ID_EDIT_EMOJI, 70, 160, emoji);
    lunar = GetClampedEditValue(hwnd, ID_EDIT_LUNAR, 60, 120, lunar);
    SendMessage(GetDlgItem(hwnd, ID_SLIDER_SOLAR), TBM_SETPOS, TRUE, solar);
    SendMessage(GetDlgItem(hwnd, ID_SLIDER_EMOJI), TBM_SETPOS, TRUE, emoji);
    SendMessage(GetDlgItem(hwnd, ID_SLIDER_LUNAR), TBM_SETPOS, TRUE, lunar);
    g_cfg.solarScalePct = solar;
    g_cfg.emojiScalePct = emoji;
    g_cfg.lunarScalePct = lunar;
    SyncScaleFromConfig();
    SetAutoRun(g_cfg.autoRun);
    CalculateMetrics();
    CreateFonts();
    ClampPositionToVisibleMonitor();
    SetWindowPos(g_hwnd, NULL, g_cfg.x, g_cfg.y, g_cfg.width, g_cfg.height, SWP_FRAMECHANGED);
    if (!g_cfg.allowMove) {
        g_hoverDay = -1;
        if (g_hTooltip) SendMessageW(g_hTooltip, TTM_POP, 0, 0);
    }
    UpdateRelativePositionFromCurrent();
    SaveConfig();
    DrawCalendar(g_hwnd);
    SendToBack();
    wchar_t fontInfo[64];
    swprintf(fontInfo, 64, L"%s, %dpt", g_cfg.fontName, g_cfg.fontSize);
    SetWindowTextW(GetDlgItem(hwnd, ID_LBL_FONTINFO), fontInfo);
    UpdateEmojiFontLabel(hwnd);
}

static LRESULT CALLBACK SettingsWndProc(HWND hwnd, UINT msg, WPARAM wP, LPARAM lP) {
    switch(msg) {
        case WM_CREATE: {
            CreateWindowW(L"STATIC", L"Vietnamese Lunar Calendar Settings \U0001F6E0\U0000FE0F",
                WS_VISIBLE | WS_CHILD | SS_CENTER, 0, 0, 0, 0, hwnd, (HMENU)ID_LBL_TITLE, NULL, NULL);
            CreateWindowW(L"STATIC", L"Position X:", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hwnd, (HMENU)ID_LBL_POSX, NULL, NULL);
            CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_NUMBER, 0, 0, 0, 0, hwnd, (HMENU)ID_EDIT_X, NULL, NULL);
            CreateWindowW(L"STATIC", L"Position Y:", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hwnd, (HMENU)ID_LBL_POSY, NULL, NULL);
            CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_NUMBER, 0, 0, 0, 0, hwnd, (HMENU)ID_EDIT_Y, NULL, NULL);
            CreateWindowW(L"STATIC", L"Opacity:", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hwnd, (HMENU)ID_LBL_OPACITY, NULL, NULL);
            CreateWindowW(L"MSCTLS_TRACKBAR32", L"", WS_VISIBLE | WS_CHILD | TBS_HORZ, 0, 0, 0, 0, hwnd, (HMENU)ID_OPACITY_TRACK, NULL, NULL);
            SendMessage(GetDlgItem(hwnd, ID_OPACITY_TRACK), TBM_SETRANGE, TRUE, MAKELPARAM(50, 255));
            SendMessage(GetDlgItem(hwnd, ID_OPACITY_TRACK), TBM_SETPOS, TRUE, g_cfg.alpha);
            CreateWindowW(L"BUTTON", L"Background Color", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, (HMENU)ID_BTN_COLOR, NULL, NULL);
            CreateWindowW(L"BUTTON", L"Font...", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, (HMENU)ID_BTN_FONT, NULL, NULL);
            wchar_t fontInfo[64];
            swprintf(fontInfo, 64, L"%s, %dpt", g_cfg.fontName, g_cfg.fontSize);
            CreateWindowW(L"STATIC", fontInfo, WS_VISIBLE | WS_CHILD | SS_LEFT, 0, 0, 0, 0, hwnd, (HMENU)ID_LBL_FONTINFO, NULL, NULL);
            CreateWindowW(L"STATIC", L"Solar Size:", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hwnd, (HMENU)ID_LBL_SOLAR, NULL, NULL);
            HWND hSolar = CreateWindowW(L"MSCTLS_TRACKBAR32", L"", WS_VISIBLE | WS_CHILD | TBS_HORZ,
                0, 0, 0, 0, hwnd, (HMENU)ID_SLIDER_SOLAR, NULL, NULL);
            SendMessageW(hSolar, TBM_SETRANGE, TRUE, MAKELPARAM(90, 180));
            SendMessageW(hSolar, TBM_SETPOS, TRUE, g_cfg.solarScalePct);
            CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_NUMBER | ES_RIGHT,
                0, 0, 0, 0, hwnd, (HMENU)ID_EDIT_SOLAR, NULL, NULL);
            CreateWindowW(L"STATIC", L"%", WS_VISIBLE | WS_CHILD,
                0, 0, 0, 0, hwnd, (HMENU)ID_LBL_PCT_SOLAR, NULL, NULL);
            CreateWindowW(L"STATIC", L"Emoji Size:", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hwnd, (HMENU)ID_LBL_EMOJI, NULL, NULL);
            HWND hEmoji = CreateWindowW(L"MSCTLS_TRACKBAR32", L"", WS_VISIBLE | WS_CHILD | TBS_HORZ,
                0, 0, 0, 0, hwnd, (HMENU)ID_SLIDER_EMOJI, NULL, NULL);
            SendMessageW(hEmoji, TBM_SETRANGE, TRUE, MAKELPARAM(70, 160));
            SendMessageW(hEmoji, TBM_SETPOS, TRUE, g_cfg.emojiScalePct);
            CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_NUMBER | ES_RIGHT,
                0, 0, 0, 0, hwnd, (HMENU)ID_EDIT_EMOJI, NULL, NULL);
            CreateWindowW(L"STATIC", L"%", WS_VISIBLE | WS_CHILD,
                0, 0, 0, 0, hwnd, (HMENU)ID_LBL_PCT_EMOJI, NULL, NULL);
            CreateWindowW(L"STATIC", L"Lunar Size:", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hwnd, (HMENU)ID_LBL_LUNAR, NULL, NULL);
            HWND hLunar = CreateWindowW(L"MSCTLS_TRACKBAR32", L"", WS_VISIBLE | WS_CHILD | TBS_HORZ,
                0, 0, 0, 0, hwnd, (HMENU)ID_SLIDER_LUNAR, NULL, NULL);
            SendMessageW(hLunar, TBM_SETRANGE, TRUE, MAKELPARAM(60, 120));
            SendMessageW(hLunar, TBM_SETPOS, TRUE, g_cfg.lunarScalePct);
            CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_NUMBER | ES_RIGHT,
                0, 0, 0, 0, hwnd, (HMENU)ID_EDIT_LUNAR, NULL, NULL);
            CreateWindowW(L"STATIC", L"%", WS_VISIBLE | WS_CHILD,
                0, 0, 0, 0, hwnd, (HMENU)ID_LBL_PCT_LUNAR, NULL, NULL);
            CreateWindowW(L"STATIC", L"Emoji Style:", WS_VISIBLE | WS_CHILD, 0, 0, 0, 0, hwnd, (HMENU)ID_LBL_EMOJI_STYLE, NULL, NULL);
            HWND hEmojiStyle = CreateWindowW(L"COMBOBOX", L"", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST,
                0, 0, 0, 0, hwnd, (HMENU)ID_EMOJI_STYLE, NULL, NULL);
            SendMessageW(hEmojiStyle, CB_ADDSTRING, 0, (LPARAM)L"Color");
            SendMessageW(hEmojiStyle, CB_ADDSTRING, 0, (LPARAM)L"Black/White");
            SendMessageW(hEmojiStyle, CB_SETCURSEL, g_cfg.emojiMono ? 1 : 0, 0);
            CreateWindowW(L"BUTTON", L"Emoji Font...", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, (HMENU)ID_BTN_EMOJI_FONT, NULL, NULL);
            CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD | SS_LEFT, 0, 0, 0, 0, hwnd, (HMENU)ID_LBL_EMOJI_FONT, NULL, NULL);
            UpdateEmojiFontLabel(hwnd);
            CreateWindowW(L"BUTTON", L"Auto Run on Windows Start", WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)ID_CHK_AUTORUN, NULL, NULL);
            SendMessage(GetDlgItem(hwnd, ID_CHK_AUTORUN), BM_SETCHECK, g_cfg.autoRun ? BST_CHECKED : BST_UNCHECKED, 0);
            CreateWindowW(L"BUTTON", L"Stick to Desktop (Below Windows)", WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)ID_CHK_STICK, NULL, NULL);
            SendMessage(GetDlgItem(hwnd, ID_CHK_STICK), BM_SETCHECK, g_cfg.stickDesktop ? BST_CHECKED : BST_UNCHECKED, 0);
            CreateWindowW(L"BUTTON", L"Lock to Primary Screen Only", WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)ID_CHK_PRIMARY, NULL, NULL);
            SendMessage(GetDlgItem(hwnd, ID_CHK_PRIMARY), BM_SETCHECK, g_cfg.primaryScreenOnly ? BST_CHECKED : BST_UNCHECKED, 0);
            CreateWindowW(L"BUTTON", L"Enable Drag Move (Click-through when off)", WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX, 0, 0, 0, 0, hwnd, (HMENU)ID_CHK_MOVE, NULL, NULL);
            SendMessage(GetDlgItem(hwnd, ID_CHK_MOVE), BM_SETCHECK, g_cfg.allowMove ? BST_CHECKED : BST_UNCHECKED, 0);
            CreateWindowW(L"BUTTON", L"Apply", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, (HMENU)ID_APPLY, NULL, NULL);
            CreateWindowW(L"BUTTON", L"OK", WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON, 0, 0, 0, 0, hwnd, (HMENU)IDOK, NULL, NULL);
            CreateWindowW(L"BUTTON", L"Cancel", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON, 0, 0, 0, 0, hwnd, (HMENU)IDCANCEL, NULL, NULL);
            wchar_t val[16];
            swprintf(val, 16, L"%d", g_cfg.x); SetWindowTextW(GetDlgItem(hwnd, ID_EDIT_X), val);
            swprintf(val, 16, L"%d", g_cfg.y); SetWindowTextW(GetDlgItem(hwnd, ID_EDIT_Y), val);
            SyncScaleFromConfig();
            SendMessageW(GetDlgItem(hwnd, ID_EDIT_SOLAR), EM_SETLIMITTEXT, 3, 0);
            SendMessageW(GetDlgItem(hwnd, ID_EDIT_EMOJI), EM_SETLIMITTEXT, 3, 0);
            SendMessageW(GetDlgItem(hwnd, ID_EDIT_LUNAR), EM_SETLIMITTEXT, 3, 0);
            SendMessageW(GetDlgItem(hwnd, ID_SLIDER_SOLAR), TBM_SETPOS, TRUE, g_cfg.solarScalePct);
            SendMessageW(GetDlgItem(hwnd, ID_SLIDER_EMOJI), TBM_SETPOS, TRUE, g_cfg.emojiScalePct);
            SendMessageW(GetDlgItem(hwnd, ID_SLIDER_LUNAR), TBM_SETPOS, TRUE, g_cfg.lunarScalePct);
            UpdateScaleEditsFromSliders(hwnd);
            LayoutSettingsControls(hwnd, GetDpiForWindowCompat(hwnd));
            return 0;
        }
        case WM_COMMAND:
            if (LOWORD(wP) == ID_BTN_COLOR) {
                CHOOSECOLOR cc = {0};
                static COLORREF customColors[16] = {0};
                cc.lStructSize = sizeof(cc);
                cc.hwndOwner = hwnd;
                cc.lpCustColors = customColors;
                cc.rgbResult = g_cfg.bgColor;
                cc.Flags = CC_FULLOPEN | CC_RGBINIT;
                if (ChooseColor(&cc)) {
                    g_cfg.bgColor = cc.rgbResult;
                }
            }
            else if (LOWORD(wP) == ID_BTN_FONT) {
                PickFont(hwnd);
                wchar_t fontInfo[64];
                swprintf(fontInfo, 64, L"%s, %dpt", g_cfg.fontName, g_cfg.fontSize);
                SetWindowTextW(GetDlgItem(hwnd, ID_LBL_FONTINFO), fontInfo);
            }
            else if (LOWORD(wP) == ID_BTN_EMOJI_FONT) {
                int sel = (int)SendMessage(GetDlgItem(hwnd, ID_EMOJI_STYLE), CB_GETCURSEL, 0, 0);
                BOOL mono = (sel == 1);
                PickEmojiFont(hwnd, mono);
                UpdateEmojiFontLabel(hwnd);
            }
            else if (LOWORD(wP) == ID_EMOJI_STYLE && HIWORD(wP) == CBN_SELCHANGE) {
                UpdateEmojiFontLabel(hwnd);
            }
            else if (LOWORD(wP) == ID_APPLY) {
                ApplySettingsFromDialog(hwnd);
            }
            else if (LOWORD(wP) == IDOK) {
                ApplySettingsFromDialog(hwnd);
                DestroyWindow(hwnd);
                g_settingsHwnd = NULL;
            }
            else if (LOWORD(wP) == IDCANCEL) {
                DestroyWindow(hwnd);
                g_settingsHwnd = NULL;
            }
            else if (HIWORD(wP) == EN_KILLFOCUS) {
                int id = LOWORD(wP);
                if (id == ID_EDIT_SOLAR) {
                    ApplyScaleEditToSlider(hwnd, ID_EDIT_SOLAR, ID_SLIDER_SOLAR, 90, 180);
                    ApplyScaleFromUI(hwnd, TRUE);
                } else if (id == ID_EDIT_EMOJI) {
                    ApplyScaleEditToSlider(hwnd, ID_EDIT_EMOJI, ID_SLIDER_EMOJI, 70, 160);
                    ApplyScaleFromUI(hwnd, TRUE);
                } else if (id == ID_EDIT_LUNAR) {
                    ApplyScaleEditToSlider(hwnd, ID_EDIT_LUNAR, ID_SLIDER_LUNAR, 60, 120);
                    ApplyScaleFromUI(hwnd, TRUE);
                }
            }
            return 0;
        case WM_HSCROLL: {
            HWND hCtl = (HWND)lP;
            if (hCtl == GetDlgItem(hwnd, ID_SLIDER_SOLAR) ||
                hCtl == GetDlgItem(hwnd, ID_SLIDER_EMOJI) ||
                hCtl == GetDlgItem(hwnd, ID_SLIDER_LUNAR)) {
                ApplyScaleFromUI(hwnd, TRUE);
                UpdateScaleEditsFromSliders(hwnd);
            }
            return 0;
        }
        case WM_DPICHANGED: {
            UINT dpi = HIWORD(wP);
            RECT* suggested = (RECT*)lP;
            SetWindowPos(hwnd, NULL, suggested->left, suggested->top,
                suggested->right - suggested->left, suggested->bottom - suggested->top,
                SWP_NOZORDER | SWP_NOACTIVATE);
            LayoutSettingsControls(hwnd, dpi);
            return 0;
        }
        case WM_CLOSE:
            DestroyWindow(hwnd);
            g_settingsHwnd = NULL;
            return 0;
    }
    return DefWindowProcW(hwnd, msg, wP, lP);
}

static void ShowSettings() {
    if (g_settingsHwnd && IsWindow(g_settingsHwnd)) {
        SetForegroundWindow(g_settingsHwnd);
        return;
    }
    WNDCLASSEXW wc = {0};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = SettingsWndProc;
    wc.hInstance = GetModuleHandle(NULL);
    wc.hIcon = g_hIcon;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = L"SettingsClass";
    wc.hIconSm = g_hIcon;
    RegisterClassExW(&wc);
    g_settingsHwnd = CreateWindowW(L"SettingsClass", L"Calendar Settings",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_VISIBLE | WS_CLIPCHILDREN,
        CW_USEDEFAULT, CW_USEDEFAULT, 520, 420, NULL, NULL, wc.hInstance, NULL);
    if (g_settingsHwnd && g_hIcon) {
        SendMessageW(g_settingsHwnd, WM_SETICON, ICON_BIG, (LPARAM)g_hIcon);
        SendMessageW(g_settingsHwnd, WM_SETICON, ICON_SMALL, (LPARAM)g_hIcon);
    }
}

static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wP, LPARAM lP) {
    switch(msg) {
        case WM_CREATE: {
            SetTimer(hwnd, 1, 60000, NULL);
            UpdateStickTimer(hwnd);
            g_hTooltip = CreateWindowExW(WS_EX_TOPMOST, TOOLTIPS_CLASSW, NULL,
                WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
                CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
                hwnd, NULL, GetModuleHandle(NULL), NULL);
            if (g_hTooltip) {
                SetWindowPos(g_hTooltip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                TOOLINFOW ti = {0};
                ti.cbSize = sizeof(ti);
                ti.uFlags = TTF_SUBCLASS;
                ti.hwnd = hwnd;
                ti.uId = 1;
                ti.lpszText = LPSTR_TEXTCALLBACK;
                GetClientRect(hwnd, &ti.rect);
                SendMessageW(g_hTooltip, TTM_ADDTOOL, 0, (LPARAM)&ti);
                SendMessageW(g_hTooltip, TTM_SETDELAYTIME, TTDT_INITIAL, 100);
                SendMessageW(g_hTooltip, TTM_SETDELAYTIME, TTDT_AUTOPOP, 5000);
            }
            return 0;
        }
        case WM_TIMER:
            if (wP == 1) {
                MaybeStartOnlineCheck();
                SYSTEMTIME st;
                int today = GetTodayYmd(&st);
                if (today != g_lastTodayYmd) {
                    g_lastTodayYmd = today;
                    BOOL needRedraw = (g_viewMonth == st.wMonth && g_viewYear == st.wYear);
                    if (g_followToday && !needRedraw) {
                        SetViewMonthYear(st.wMonth, st.wYear, TRUE);
                        needRedraw = TRUE;
                    }
                    if (needRedraw) DrawCalendar(hwnd);
                }
            } else if (wP == 2 && g_cfg.stickDesktop) {
                if (!g_mouseInside && !g_dragging) {
                    SetWindowPos(hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                    g_stickRaised = FALSE;
                }
            } else if (wP == 3) {
                g_headerPulse += 0.12f;
                if (g_headerPulse > 6.28318f) g_headerPulse -= 6.28318f;
                if (g_hoverHeader) DrawCalendar(hwnd);
            }
            return 0;
        case WM_ONLINE_UPDATE:
            DrawCalendar(hwnd);
            return 0;
        case WM_DISPLAYCHANGE:
            Sleep(200);
            CalculateMetrics();
            CreateFonts();
            RestorePositionFromConfig();
            SetWindowPos(g_hwnd, NULL, g_cfg.x, g_cfg.y, g_cfg.width, g_cfg.height, SWP_NOZORDER | SWP_NOACTIVATE);
            UpdateRelativePositionFromCurrent();
            SaveConfig();
            DrawCalendar(hwnd);
            return 0;
        case WM_DPICHANGED: {
            int dpi = HIWORD(wP);
            g_dpiScale = dpi / 96.0f;
            if (g_dpiScale < 1.0f) g_dpiScale = 1.0f;
            CalculateMetrics();
            CreateFonts();
            RECT* suggested = (RECT*)lP;
            g_cfg.x = suggested->left;
            g_cfg.y = suggested->top;
            ClampPositionToVisibleMonitor();
            SetWindowPos(hwnd, NULL, g_cfg.x, g_cfg.y, g_cfg.width, g_cfg.height, SWP_NOZORDER | SWP_NOACTIVATE);
            UpdateRelativePositionFromCurrent();
            SaveConfig();
            DrawCalendar(hwnd);
            return 0;
        }
        case WM_NCHITTEST:
            return HTCLIENT;
        case WM_MOUSEWHEEL: {
            POINT pt;
            GetCursorPos(&pt);
            ScreenToClient(hwnd, &pt);
            if (!PtInRect(&g_rcHeader, pt)) return 0;
            short delta = GET_WHEEL_DELTA_WPARAM(wP);
            int step = (delta > 0) ? -1 : 1;
            if (GET_KEYSTATE_WPARAM(wP) & MK_SHIFT) {
                SetViewMonthYear(g_viewMonth, g_viewYear + step, FALSE);
            } else {
                SetViewMonthYear(g_viewMonth + step, g_viewYear, FALSE);
            }
            DrawCalendar(hwnd);
            return 0;
        }
        case WM_MOUSEMOVE: {
            if (g_dragging && GetCapture() == hwnd) {
                POINT pt;
                GetCursorPos(&pt);
                g_cfg.x = pt.x - g_dragOffset.x;
                g_cfg.y = pt.y - g_dragOffset.y;
                ClampPositionToVisibleMonitor();
                SetWindowPos(hwnd, NULL, g_cfg.x, g_cfg.y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
                DrawCalendar(hwnd);
            } else {
                if (!g_trackingMouse) {
                    TRACKMOUSEEVENT tme = {sizeof(tme), TME_LEAVE, hwnd, 0};
                    TrackMouseEvent(&tme);
                    g_trackingMouse = TRUE;
                }
                POINT pt;
                pt.x = LOWORD(lP);
                pt.y = HIWORD(lP);
                BOOL insideCard = PtInRect(&g_cardRect, pt);
                if (insideCard && !g_mouseInside) {
                    g_mouseInside = TRUE;
                    if (g_cfg.stickDesktop) {
                        SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                        g_stickRaised = TRUE;
                    }
                }
                BOOL overHeader = PtInRect(&g_rcHeader, pt);
                SetHeaderHover(hwnd, overHeader);
                if (overHeader) SetFocus(hwnd);
                int nav = -1;
                if (PtInRect(&g_rcPrev, pt)) nav = 0;
                else if (PtInRect(&g_rcNext, pt)) nav = 1;
                if (nav != g_hoverNav) {
                    g_hoverNav = nav;
                    DrawCalendar(hwnd);
                }
                int day = GetDayFromPoint(pt.x, pt.y);
                if (day != g_hoverDay) {
                    g_hoverDay = day;
                    if (day > 0) {
                        UpdateTooltip(day);
                    } else if (g_hTooltip) {
                        SendMessageW(g_hTooltip, TTM_POP, 0, 0);
                    }
                    DrawCalendar(hwnd);
                }
            }
            return 0;
        }
        case WM_MOUSELEAVE:
            g_trackingMouse = FALSE;
            if (g_mouseInside) {
                g_mouseInside = FALSE;
                if (g_cfg.stickDesktop && g_stickRaised) {
                    SetWindowPos(hwnd, HWND_BOTTOM, 0, 0, 0, 0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                    g_stickRaised = FALSE;
                }
            }
            SetHeaderHover(hwnd, FALSE);
            if (g_hoverNav != -1 || g_hoverDay != -1) {
                g_hoverNav = -1;
                g_hoverDay = -1;
                if (g_hTooltip) SendMessageW(g_hTooltip, TTM_POP, 0, 0);
                DrawCalendar(hwnd);
            }
            return 0;
        case WM_LBUTTONDOWN: {
            POINT pt;
            pt.x = LOWORD(lP);
            pt.y = HIWORD(lP);
            if (!g_mouseInside) {
                g_mouseInside = TRUE;
                if (g_cfg.stickDesktop) {
                    SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
                    g_stickRaised = TRUE;
                }
            }
            if (PtInRect(&g_rcPrev, pt)) {
                SetViewMonthYear(g_viewMonth - 1, g_viewYear, FALSE);
                DrawCalendar(hwnd);
                return 0;
            }
            if (PtInRect(&g_rcNext, pt)) {
                SetViewMonthYear(g_viewMonth + 1, g_viewYear, FALSE);
                DrawCalendar(hwnd);
                return 0;
            }
            if (!g_cfg.allowMove) return 0;
            SetCapture(hwnd);
            g_dragging = TRUE;
            g_dragOffset = pt;
            return 0;
        }
        case WM_LBUTTONUP:
            if (GetCapture() == hwnd) {
                ReleaseCapture();
                g_dragging = FALSE;
                UpdateRelativePositionFromCurrent();
                SaveConfig();
                if (g_cfg.stickDesktop) SendToBack();
            }
            return 0;
        case WM_LBUTTONDBLCLK: {
            POINT pt;
            pt.x = LOWORD(lP);
            pt.y = HIWORD(lP);
            if (PtInRect(&g_rcPrev, pt) || PtInRect(&g_rcNext, pt)) {
                return 0;
            }
            if (PtInRect(&g_rcHeader, pt)) {
                SetViewToToday();
                DrawCalendar(hwnd);
            }
            return 0;
        }
        case WM_TRAYICON:
            if (lP == WM_RBUTTONUP) {
                POINT pt;
                GetCursorPos(&pt);
                HMENU hMenu = CreatePopupMenu();
                InsertMenuW(hMenu, 0, MF_BYPOSITION | MF_STRING, ID_SETTINGS, L"Settings...");
                InsertMenuW(hMenu, 1, MF_BYPOSITION | MF_STRING, ID_PREV_MONTH, L"Previous Month");
                InsertMenuW(hMenu, 2, MF_BYPOSITION | MF_STRING, ID_NEXT_MONTH, L"Next Month");
                InsertMenuW(hMenu, 3, MF_BYPOSITION | MF_STRING, ID_GOTO_TODAY, L"Go to Today");
                InsertMenuW(hMenu, 4, MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
                InsertMenuW(hMenu, 5, MF_BYPOSITION | MF_STRING, ID_RESET_POS, L"Reset to Top-Right");
                InsertMenuW(hMenu, 6, MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
                wchar_t stickTxt[64];
                swprintf(stickTxt, 64, L"Stick Desktop: %s", g_cfg.stickDesktop ? L"ON" : L"OFF");
                InsertMenuW(hMenu, 7, MF_BYPOSITION | MF_STRING, 1005, stickTxt);
                wchar_t primaryTxt[64];
                swprintf(primaryTxt, 64, L"Primary Screen: %s", g_cfg.primaryScreenOnly ? L"ON" : L"OFF");
                InsertMenuW(hMenu, 8, MF_BYPOSITION | MF_STRING, 1006, primaryTxt);
                wchar_t moveTxt[64];
                swprintf(moveTxt, 64, L"Move/Click Through: %s", g_cfg.allowMove ? L"ON" : L"OFF");
                InsertMenuW(hMenu, 9, MF_BYPOSITION | MF_STRING, ID_TOGGLE_MOVE, moveTxt);
                InsertMenuW(hMenu, 10, MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
                InsertMenuW(hMenu, 11, MF_BYPOSITION | (g_cfg.autoRun ? MF_CHECKED : MF_UNCHECKED) | MF_STRING, ID_TOGGLE_AUTORUN, L"Auto Run");
                InsertMenuW(hMenu, 12, MF_BYPOSITION | MF_STRING, ID_EXIT, L"Exit");
                SetForegroundWindow(hwnd);
                int cmd = TrackPopupMenu(hMenu, TPM_RETURNCMD | TPM_RIGHTALIGN | TPM_BOTTOMALIGN, pt.x, pt.y, 0, hwnd, NULL);
                DestroyMenu(hMenu);
                if (cmd == ID_SETTINGS) ShowSettings();
                else if (cmd == ID_PREV_MONTH) {
                    SetViewMonthYear(g_viewMonth - 1, g_viewYear, FALSE);
                    DrawCalendar(hwnd);
                }
                else if (cmd == ID_NEXT_MONTH) {
                    SetViewMonthYear(g_viewMonth + 1, g_viewYear, FALSE);
                    DrawCalendar(hwnd);
                }
                else if (cmd == ID_GOTO_TODAY) {
                    SetViewToToday();
                    DrawCalendar(hwnd);
                }
                else if (cmd == ID_RESET_POS) {
                    RECT rc;
                    if (g_cfg.primaryScreenOnly) GetPrimaryScreenRect(&rc);
                    else {
                        POINT p = {g_cfg.x + g_cfg.width / 2, g_cfg.y + g_cfg.height / 2};
                        GetMonitorWorkAreaFromPoint(p, &rc, NULL, 0);
                    }
                    g_cfg.x = rc.right - g_cfg.width - (int)(10 * g_dpiScale);
                    g_cfg.y = rc.top + (int)(10 * g_dpiScale);
                    SetWindowPos(hwnd, NULL, g_cfg.x, g_cfg.y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
                    UpdateRelativePositionFromCurrent();
                    SaveConfig();
                    DrawCalendar(hwnd);
                }
                else if (cmd == 1005) {
                    g_cfg.stickDesktop = !g_cfg.stickDesktop;
                    SaveConfig();
                    UpdateStickTimer(hwnd);
                    if (g_cfg.stickDesktop) SendToBack();
                }
                else if (cmd == 1006) {
                    g_cfg.primaryScreenOnly = !g_cfg.primaryScreenOnly;
                    RestorePositionFromConfig();
                    SetWindowPos(hwnd, NULL, g_cfg.x, g_cfg.y, g_cfg.width, g_cfg.height, SWP_NOZORDER | SWP_NOACTIVATE);
                    UpdateRelativePositionFromCurrent();
                    SaveConfig();
                    DrawCalendar(hwnd);
                }
                else if (cmd == ID_TOGGLE_MOVE) {
                    g_cfg.allowMove = !g_cfg.allowMove;
                    SaveConfig();
                    UpdateClickThroughStyle();
                    g_hoverDay = -1;
                    if (g_hTooltip) SendMessageW(g_hTooltip, TTM_POP, 0, 0);
                    DrawCalendar(hwnd);
                }
                else if (cmd == ID_TOGGLE_AUTORUN) {
                    g_cfg.autoRun = !g_cfg.autoRun;
                    SaveConfig();
                    SetAutoRun(g_cfg.autoRun);
                }
                else if (cmd == ID_EXIT) DestroyWindow(hwnd);
            }
            else if (lP == WM_LBUTTONUP) {
                DrawCalendar(hwnd);
            }
            return 0;
        case WM_DESTROY:
            Shell_NotifyIconW(NIM_DELETE, &g_nid);
            if (g_hTooltip) DestroyWindow(g_hTooltip);
            PostQuitMessage(0);
            return 0;
    }
    return DefWindowProcW(hwnd, msg, wP, lP);
}

int WINAPI WinMain(HINSTANCE hInst, HINSTANCE hPrev, LPSTR cmd, int show) {
    if (!IsSupportedWindows()) {
        MessageBoxW(NULL,
            L"This app supports Windows 10 and 11 only.",
            L"Unsupported Windows",
            MB_OK | MB_ICONERROR);
        return 0;
    }
    HMODULE hUser32 = LoadLibraryW(L"user32.dll");
    if (hUser32) {
        typedef BOOL (WINAPI *SetProcessDpiAwarenessContext_t)(DPI_AWARENESS_CONTEXT);
        SetProcessDpiAwarenessContext_t SetProcessDpiAwarenessContext =
            (SetProcessDpiAwarenessContext_t)GetProcAddress(hUser32, "SetProcessDpiAwarenessContext");
        if (SetProcessDpiAwarenessContext) {
            SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
        } else {
            SetProcessDPIAware();
        }
        FreeLibrary(hUser32);
    }
    HDC hdc = GetDC(NULL);
    int dpi = GetDeviceCaps(hdc, LOGPIXELSX);
    ReleaseDC(NULL, hdc);
    g_dpiScale = dpi / 96.0f;
    if (g_dpiScale < 1.0f) g_dpiScale = 1.0f;
    LoadConfig();
    if (wcslen(g_cfg.fontName) == 0) wcscpy(g_cfg.fontName, L"Segoe UI");
    EnsureFontFallback();
    SyncScaleFromConfig();
    InitializeCriticalSection(&g_onlineLock);
    g_onlineLockInit = TRUE;
    if (g_cfg.autoRun) SetAutoRun(TRUE);
    SetViewToToday();
    g_lastTodayYmd = GetTodayYmd(NULL);
    CalculateMetrics();
    if (g_cfg.x < 0 || g_cfg.y < 0) {
        RECT rc;
        GetPrimaryScreenRect(&rc);
        g_cfg.x = rc.right - g_cfg.width - (int)(10 * g_dpiScale);
        g_cfg.y = rc.top + (int)(10 * g_dpiScale);
        UpdateRelativePositionFromCurrent();
        SaveConfig();
    } else {
        RestorePositionFromConfig();
        UpdateRelativePositionFromCurrent();
        SaveConfig();
    }
    g_hIcon = LoadIconW(hInst, MAKEINTRESOURCE(1));
    if (!g_hIcon) g_hIcon = LoadIconW(NULL, IDI_APPLICATION);
    WNDCLASSEXW wc = {sizeof(wc), CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS, WndProc, 0, 0, hInst,
        g_hIcon, LoadCursor(NULL, IDC_ARROW), NULL, NULL, APP_NAME, g_hIcon};
    RegisterClassExW(&wc);
    DWORD exStyle = WS_EX_LAYERED | WS_EX_TOOLWINDOW;
    g_hwnd = CreateWindowExW(exStyle,
        APP_NAME, L"Vietnamese Lunar Calendar Desktop Widget", WS_POPUP,
        g_cfg.x, g_cfg.y, g_cfg.width, g_cfg.height, NULL, NULL, hInst, NULL);
    g_nid.cbSize = sizeof(g_nid);
    g_nid.hWnd = g_hwnd;
    g_nid.uID = 1;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP | NIF_SHOWTIP;
    g_nid.uCallbackMessage = WM_TRAYICON;
    g_nid.hIcon = g_hIcon;
    wcscpy(g_nid.szTip, L"Vietnamese Lunar Calendar");
    Shell_NotifyIconW(NIM_ADD, &g_nid);
    MaybeStartOnlineCheck();
    CreateFonts();
    DrawCalendar(g_hwnd);
    ShowWindow(g_hwnd, SW_SHOW);
    UpdateClickThroughStyle();
    SendToBack();
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    DeleteObject(g_fontBig);
    DeleteObject(g_fontSmall);
    DeleteObject(g_fontMonth);
    DeleteObject(g_fontLunar);
    DeleteObject(g_fontEmoji);
    if (g_settingsFont) DeleteObject(g_settingsFont);
    ReleaseDWriteResources();
    if (g_hIcon) DestroyIcon(g_hIcon);
    return 0;
}

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE hPrev, PWSTR cmd, int show) {
    return WinMain(hInst, hPrev, GetCommandLineA(), show);
}

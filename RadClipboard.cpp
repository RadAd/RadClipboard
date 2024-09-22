#include "Rad/Window.h"
#include "Rad/Windowxx.h"
//#include <tchar.h>
#include <strsafe.h>
#include <memory>
#include <string>
#include <vector>
#include <set>
#include <ctime>

#include <Shlobj.h>

#include "Rad/MemoryPlus.h"
#include "Rad/WinError.h"
#include "Rad/Log.h"

const UINT CF_HTML = RegisterClipboardFormat(TEXT("HTML Format"));
const UINT CF_LINK = RegisterClipboardFormat(TEXT("Link Preview Format"));
const UINT CF_HYPERLINK = RegisterClipboardFormat(TEXT("Titled Hyperlink Format"));
const UINT CF_RICH = RegisterClipboardFormat(TEXT("Rich Text Format"));
const UINT CF_FILENAME = RegisterClipboardFormat(TEXT("Filename"));
const UINT CF_FILENAMEW = RegisterClipboardFormat(TEXT("FilenameW"));

#define HK_HIST (4)

template <class T, class U>
bool contains(const std::set<T>& s, const U& t)
{
    return s.find(t) != s.end();
}

HGLOBAL Duplicate(HGLOBAL hData)
{
    if (hData == NULL)
        return NULL;

    const SIZE_T s = GlobalSize(hData);
    if (s <= 0)
        return NULL;
    HANDLE hCopy = GlobalAlloc(GHND, s);
    if (hCopy == NULL)
        return NULL;

    {
        auto pSrc = AutoGlobalLock<void*>(hData);
        auto pDst = AutoGlobalLock<void*>(hCopy);
        memcpy(pDst.get(), pSrc.get(), s);
    }

    return hCopy;
}

inline std::tstring GetFormatName(const UINT f)
{
    switch (f)
    {
    case CF_TEXT:           return TEXT("Text");
    case CF_BITMAP:         return TEXT("Bitmap");
    case CF_METAFILEPICT:   return TEXT("Metafile Picture");
    case CF_OEMTEXT:        return TEXT("OEM Text");
    case CF_DIB:            return TEXT("Device Independent Bitmap");
    case CF_UNICODETEXT:    return TEXT("Unicode Text");
    case CF_LOCALE:         return TEXT("Locale");
    case CF_HDROP:          return TEXT("Files");
    case CF_DIBV5:          return TEXT("Device Independent Bitmap V5");
    default:
    {
        TCHAR name[1024] = TEXT("");
        if (GetClipboardFormatName(f, name, ARRAYSIZE(name)) == 0)
            StringCchPrintf(name, ARRAYSIZE(name), TEXT("Format: %u"), f);
        return name;
    }
    }
}

BOOL OpenClipboardTimeout(HWND hWnd, UINT timeout)
{
    const std::clock_t t = std::clock() + timeout * 1000 / CLOCKS_PER_SEC;
    while (!OpenClipboard(hWnd))
        if (std::clock() > t)
            return FALSE;
    return TRUE;
}

struct HistItem
{
    UINT uFormat;
    HANDLE hData;
};

class RadClipboardViewerWnd : public Window
{
    friend WindowManager<RadClipboardViewerWnd>;
public:
    static ATOM Register() { return WindowManager<RadClipboardViewerWnd>::Register(); }
    static RadClipboardViewerWnd* Create() { return WindowManager<RadClipboardViewerWnd>::Create(); }

private:
    RadClipboardViewerWnd() = default;

protected:
    static void GetWndClass(WNDCLASS& wc);
    static void GetCreateWindow(CREATESTRUCT& cs);
    LRESULT HandleMessage(UINT uMsg, WPARAM wParam, LPARAM lParam) override;

private:
    BOOL OnCreate(LPCREATESTRUCT lpCreateStruct);
    void OnDestroy();
    void OnClipboardUpdate();
    void OnContextMenu(HWND hWndContext, UINT xPos, UINT yPos);
    void OnHotKey(int idHotKey, UINT fuModifiers, UINT vk);

    virtual void OnDraw(const PAINTSTRUCT* pps) const override;

    static LPCTSTR ClassName() { return TEXT("RadClipboard"); }

    std::vector<UINT> m_formats;
    UINT m_uFormat = 0;
    mutable std::tstring m_text;

    std::vector<std::vector<HistItem>> m_history;
};

void RadClipboardViewerWnd::GetWndClass(WNDCLASS& wc)
{
    Window::GetWndClass(wc);
    wc.style |= CS_HREDRAW | CS_VREDRAW;
    wc.hbrBackground = CreateSolidBrush(RGB(31, 31, 31));
}

void RadClipboardViewerWnd::GetCreateWindow(CREATESTRUCT& cs)
{
    Window::GetCreateWindow(cs);
    cs.lpszName = TEXT("Rad Clipboard");
    cs.style = WS_OVERLAPPED | WS_CAPTION | WS_THICKFRAME;
    cs.cx = 300;
    cs.cy = 100;
    cs.dwExStyle = WS_EX_TOOLWINDOW | WS_EX_TOPMOST;
}

BOOL RadClipboardViewerWnd::OnCreate(const LPCREATESTRUCT lpCreateStruct)
{
    CHECK_LE(AddClipboardFormatListener(*this));
    OnClipboardUpdate();
    RegisterHotKey(*this, HK_HIST, MOD_CONTROL | MOD_SHIFT, 'V');
    return TRUE;
}

void RadClipboardViewerWnd::OnDestroy()
{
    CHECK_LE(RemoveClipboardFormatListener(*this));
    PostQuitMessage(0);
}

void RadClipboardViewerWnd::OnClipboardUpdate()
{
    CHECK_LE(OpenClipboardTimeout(*this, 100));
    m_formats.clear();
    {
        UINT format = 0;
        while (format = EnumClipboardFormats(format))
            m_formats.push_back(format);
    }

    m_uFormat = 0;
    for (const UINT f : m_formats)
    {
        //TCHAR name[1024] = TEXT("");
        //GetClipboardFormatName(f, name, ARRAYSIZE(name));
        if (f <= CF_MAX)
        {
            m_uFormat = f;
            break;
        }
    }
    
    std::set<UINT> format_done;
    m_history.insert(m_history.begin(), 1, {});
    std::vector<HistItem>& n = m_history.front();
    for (const UINT f : m_formats)
    {
        bool skip = false;
        // Skip according to https://learn.microsoft.com/en-us/windows/win32/dataxchg/clipboard-formats#synthesized-clipboard-formats
        switch (f)
        {
        case CF_TEXT: case CF_UNICODETEXT: case CF_OEMTEXT:
            skip = contains(format_done, CF_TEXT) || contains(format_done, CF_UNICODETEXT) || contains(format_done, CF_OEMTEXT);
            break;
        case CF_BITMAP: case CF_DIB: case CF_DIBV5:
            skip = contains(format_done, CF_BITMAP) || contains(format_done, CF_DIB) || contains(format_done, CF_DIBV5);
            break;
        case CF_PALETTE:
            skip = contains(format_done, CF_DIB) || contains(format_done, CF_DIBV5);
            break;
        case CF_ENHMETAFILE: case CF_METAFILEPICT:
            skip = contains(format_done, CF_ENHMETAFILE) || contains(format_done, CF_METAFILEPICT);
            break;
        }
        if (skip)
            continue;

        const HANDLE hData = GetClipboardData(f);
        HANDLE hCopy = NULL;
        switch (f)
        {   // Copy according to https://learn.microsoft.com/en-gb/windows/win32/dataxchg/standard-clipboard-formats#constants
        case CF_BITMAP:
            hCopy = CopyImage(hData, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE);
            break;
        default:
            hCopy = Duplicate(hData);
            break;
        }
        if (hCopy == NULL)
            continue;

        format_done.insert(f);
        n.push_back({ f, hCopy});
    }

    CHECK_LE(CloseClipboard());

    m_text.clear();
    CHECK_LE(InvalidateRect(*this, nullptr, TRUE));
}

void RadClipboardViewerWnd::OnContextMenu(HWND hWndContext, UINT xPos, UINT yPos)
{
    auto hMenu = MakeUniqueHandle(CreatePopupMenu(), DestroyMenu);
    static const UINT CommandBegin = 0x100;
    for (const UINT f : m_formats)
        CHECK_LE(AppendMenu(hMenu.get(), MF_STRING | (f == m_uFormat ? MF_CHECKED : MF_UNCHECKED), CommandBegin + f, GetFormatName(f).c_str()));

    const int Command = TrackPopupMenu(hMenu.get(), TPM_LEFTBUTTON | TPM_LEFTALIGN | TPM_RETURNCMD, xPos, yPos, 0, *this, nullptr);
    if (Command > 0)
    {
        m_uFormat = Command - CommandBegin;
        m_text.clear();
        CHECK_LE(InvalidateRect(*this, nullptr, TRUE));
    }
}

void RadClipboardViewerWnd::OnHotKey(int idHotKey, UINT fuModifiers, UINT vk)
{
    if (idHotKey == HK_HIST)
    {
        SetForegroundWindow(*this);
        POINT pt;
        CHECK_LE(GetCursorPos(&pt));

        auto hMenu = MakeUniqueHandle(CreatePopupMenu(), DestroyMenu);

        static const UINT CommandBegin = 0x100;
        UINT id = CommandBegin;
        for (const auto& i : m_history)
        {
            TCHAR name[1024] = TEXT("");
            for (const auto& e : i)
            {
                bool dobreak = true;
                switch (e.uFormat)
                {
                case CF_TEXT:
                {
                    auto pStr = AutoGlobalLock<LPCSTR>(e.hData);
                    StringCchPrintf(name, ARRAYSIZE(name), TEXT("%.50S"), pStr.get());
                    break;
                }
                case CF_UNICODETEXT:
                {
                    auto pStr = AutoGlobalLock<LPWSTR>(e.hData);
                    StringCchPrintf(name, ARRAYSIZE(name), TEXT("%.50s"), pStr.get());
                    break;
                }
                case CF_BITMAP:
                {
                    const HBITMAP hbmp = (HBITMAP) e.hData;
                    BITMAP bm;
                    GetObject(hbmp, sizeof(bm), &bm);
                    StringCchPrintf(name, ARRAYSIZE(name), TEXT("Image %d x %d x %d"), bm.bmWidth, bm.bmHeight, bm.bmBitsPixel);
                    break;
                }
                case CF_DIB:
                {
                    auto pBitmapInfo = AutoGlobalLock<LPBITMAPINFO>(e.hData);
                    StringCchPrintf(name, ARRAYSIZE(name), TEXT("Image %d x %d x %d"), pBitmapInfo->bmiHeader.biWidth, pBitmapInfo->bmiHeader.biHeight, pBitmapInfo->bmiHeader.biBitCount);
                    break;
                }
                case CF_HDROP:
                {
                    int count = 0;
                    auto pDropFiles = AutoGlobalLock<LPDROPFILES>(e.hData);
                    if (pDropFiles->fWide)
                    {
                        LPWSTR pStr = (LPWSTR)((BYTE*)pDropFiles.get() + pDropFiles->pFiles);
                        while (*pStr)
                        {
                            size_t len = 0;
                            CHECK_HR(StringCchLengthW(pStr, MAX_PATH, &len));

                            ++count;

                            pStr += len + 1;
                        }
                    }
                    else
                    {
                        LPSTR pStr = (LPSTR)((BYTE*)pDropFiles.get() + pDropFiles->pFiles);
                        while (*pStr)
                        {
                            size_t len = 0;
                            CHECK_HR(StringCchLengthA(pStr, MAX_PATH, &len));

                            ++count;

                            pStr += len + 1;
                        }
                    }
                    StringCchPrintf(name, ARRAYSIZE(name), TEXT("%d Files"), count);
                    break;
                }
                default:
                    dobreak = false;
                    break;
                }

                if (dobreak)
                    break;
            }

            if (name[0] == TEXT('\0'))
                StringCchPrintf(name, ARRAYSIZE(name), TEXT("Item: %u"), i.front().uFormat);
            AppendMenu(hMenu.get(), MF_STRING, id++, name);
        }

        const int Command = TrackPopupMenu(hMenu.get(), TPM_LEFTBUTTON | TPM_LEFTALIGN | TPM_RETURNCMD, pt.x, pt.y, 0, *this, nullptr);
        if (Command > CommandBegin)
        {
            const auto& i = m_history[Command - CommandBegin];
            CHECK_LE(OpenClipboardTimeout(*this, 100));
            CHECK_LE(EmptyClipboard());

            for (const auto& e : i)
                //SetClipboardData(e.uFormat, Duplicate(e.hData));
                SetClipboardData(e.uFormat, e.hData);

            CHECK_LE(CloseClipboard());

            m_history.erase(std::next(m_history.begin(), Command - CommandBegin));
        }
    }
}

LRESULT RadClipboardViewerWnd::HandleMessage(const UINT uMsg, const WPARAM wParam, const LPARAM lParam)
{
    LRESULT ret = 0;
    switch (uMsg)
    {
        HANDLE_MSG(WM_CREATE, OnCreate);
        HANDLE_MSG(WM_DESTROY, OnDestroy);
        HANDLE_MSG(WM_CLIPBOARDUPDATE, OnClipboardUpdate);
        HANDLE_MSG(WM_CONTEXTMENU, OnContextMenu);
        HANDLE_MSG(WM_HOTKEY, OnHotKey);
    }

    if (!IsHandled())
        ret = Window::HandleMessage(uMsg, wParam, lParam);

    return ret;
}

void RadClipboardViewerWnd::OnDraw(const PAINTSTRUCT* pps) const
{
    RECT rc;
    CHECK_LE(GetClientRect(*this, &rc));

    SetBkMode(pps->hdc, TRANSPARENT);
    SetTextColor(pps->hdc, RGB(250, 250, 223));
    if (m_uFormat == 0)
    {
        DrawText(pps->hdc, TEXT("The clipboard is empty."), -1, &rc, DT_CENTER | DT_VCENTER | DT_NOPREFIX | DT_WORDBREAK);
    }
    else
    {
        CHECK_LE(OpenClipboardTimeout(*this, 100));
        const HANDLE hData = GetClipboardData(m_uFormat);
        if (hData != NULL)
        {
            UINT uFormat = m_uFormat;
            if (uFormat == CF_HTML || uFormat == CF_LINK || uFormat == CF_HYPERLINK || uFormat == CF_RICH || uFormat == CF_FILENAME)
                uFormat = CF_TEXT;  // UTF-8
            else if (uFormat == CF_FILENAMEW)
                uFormat = CF_UNICODETEXT;
            auto hOldFont = AutoSelectObject(pps->hdc, GetStockObject(SYSTEM_FONT));
            switch (uFormat)
            {
            case CF_OEMTEXT:
                SelectObject(pps->hdc, GetStockObject(OEM_FIXED_FONT));
                // Fallthrough
            case CF_TEXT:
            {
                auto pStr = AutoGlobalLock<LPCSTR>(hData);
                if (pStr)
                    DrawTextA(pps->hdc, pStr.get(), -1, &rc, DT_TOP | DT_LEFT | DT_NOPREFIX | DT_WORDBREAK | DT_EXPANDTABS);
                break;
            }
            case CF_UNICODETEXT:
            {
                auto pStr = AutoGlobalLock<LPWSTR>(hData);
                if (pStr)
                    DrawTextW(pps->hdc, pStr.get(), -1, &rc, DT_TOP | DT_LEFT | DT_NOPREFIX | DT_WORDBREAK | DT_EXPANDTABS);
                break;
            }
            case CF_BITMAP:
            {
                const HBITMAP hbm = (HBITMAP) hData;
                const auto hdcMem = MakeUniqueHandle(CreateCompatibleDC(pps->hdc), DeleteDC);
                if (hdcMem)
                {
                    SelectObject(hdcMem.get(), hbm);
                    BitBlt(pps->hdc, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, hdcMem.get(), 0, 0, SRCCOPY);
                }
                break;
            }
            case CF_DIB:
            {
                auto pBitmapInfo = AutoGlobalLock<LPBITMAPINFO>(hData);
                size_t offset = pBitmapInfo->bmiHeader.biClrUsed * sizeof(RGBQUAD);
                if (pBitmapInfo->bmiHeader.biCompression == BI_BITFIELDS
                    and (pBitmapInfo->bmiHeader.biBitCount == 16 or pBitmapInfo->bmiHeader.biBitCount == 32)
                    and pBitmapInfo->bmiHeader.biSize == 40)
                    offset += 12; // mask-bytes
                //offset += sizeof(RGBQUAD);
                void *pDIBBits = (BYTE*) ((LPBITMAPINFOHEADER) pBitmapInfo.get() + 1) + offset;
                ::StretchDIBits(pps->hdc,
                    rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top,
                    //0, 0, pBitmapInfo->bmiHeader.biWidth, pBitmapInfo->bmiHeader.biHeight,
                    0, pBitmapInfo->bmiHeader.biHeight - (rc.bottom - rc.top), rc.right - rc.left, rc.bottom - rc.top,
                    pDIBBits, pBitmapInfo.get(), DIB_RGB_COLORS, SRCCOPY);
                break;
            }
            case CF_DIBV5:
            {
                auto pBitmapInfo = AutoGlobalLock<LPBITMAPV5HEADER>(hData);
                size_t offset = pBitmapInfo->bV5ClrUsed * sizeof(RGBQUAD);
                if (pBitmapInfo->bV5Compression == BI_BITFIELDS
                    and (pBitmapInfo->bV5BitCount == 16 or pBitmapInfo->bV5BitCount == 32)
                    and pBitmapInfo->bV5Size == 40)
                    offset += 12; // mask-bytes
                offset += sizeof(RGBQUAD) * 3; // TODO Not sure why this is here?
                void* pDIBBits = (BYTE*) (pBitmapInfo.get() + 1) + offset;
                ::StretchDIBits(pps->hdc,
                    rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top,
                    //0, 0, pBitmapInfo->bmiHeader.biWidth, pBitmapInfo->bmiHeader.biHeight,
                    0, pBitmapInfo->bV5Height - (rc.bottom - rc.top), rc.right - rc.left, rc.bottom - rc.top,
                    pDIBBits, (LPBITMAPINFO) pBitmapInfo.get(), DIB_RGB_COLORS, SRCCOPY);
                break;
            }
            case CF_ENHMETAFILE:
            {
                const HENHMETAFILE hEmf = (HENHMETAFILE) hData;
                PlayEnhMetaFile(pps->hdc, hEmf, &rc);
                break;
            }
            case CF_OWNERDISPLAY:
            {
                const HWND hWndOwner = GetClipboardOwner();
                const auto hglb = MakeUniqueHandle(GlobalAlloc(GMEM_MOVEABLE, sizeof(PAINTSTRUCT)), &GlobalFree);
                if (hglb)
                {
                    {
                        auto lpps = AutoGlobalLock<LPPAINTSTRUCT>(hglb.get());
                        if (lpps)
                            memcpy(lpps.get(), pps, sizeof(PAINTSTRUCT));
                    }

                    SendMessage(hWndOwner, WM_PAINTCLIPBOARD, (WPARAM) (HWND) *this, (LPARAM) hglb.get());
                }
                break;
            }
            case CF_LOCALE:
            {
                auto pLCID = AutoGlobalLock<LCID*>(hData);
                TCHAR name[100];
                LCIDToLocaleName(*(pLCID.get()), name, ARRAYSIZE(name), LOCALE_ALLOW_NEUTRAL_NAMES);
                DrawText(pps->hdc, name, -1, &rc, DT_TOP | DT_LEFT);
                break;
            }
            case CF_HDROP:
            {
                auto pDropFiles = AutoGlobalLock<LPDROPFILES>(hData);
                if (pDropFiles->fWide)
                {
                    LPWSTR pStr = (LPWSTR) ((BYTE*) pDropFiles.get() + pDropFiles->pFiles);
                    while (*pStr)
                    {
                        size_t len = 0;
                        CHECK_HR(StringCchLengthW(pStr, MAX_PATH, &len));

                        RECT src = rc;
                        DrawTextW(pps->hdc, pStr, int(len), &src, DT_TOP | DT_LEFT | DT_NOPREFIX | DT_PATH_ELLIPSIS | DT_EXPANDTABS);
                        DrawTextW(pps->hdc, pStr, int(len), &src, DT_TOP | DT_LEFT | DT_NOPREFIX | DT_PATH_ELLIPSIS | DT_EXPANDTABS | DT_CALCRECT);
                        rc.top += src.bottom - src.top;

                        pStr += len + 1;
                    }
                }
                else
                {
                    LPSTR pStr = (LPSTR) ((BYTE*) pDropFiles.get() + pDropFiles->pFiles);
                    while (*pStr)
                    {
                        size_t len = 0;
                        CHECK_HR(StringCchLengthA(pStr, MAX_PATH, &len));

                        RECT src = rc;
                        DrawTextA(pps->hdc, pStr, int(len), &src, DT_TOP | DT_LEFT | DT_NOPREFIX | DT_PATH_ELLIPSIS | DT_EXPANDTABS);
                        DrawTextA(pps->hdc, pStr, int(len), &src, DT_TOP | DT_LEFT | DT_NOPREFIX | DT_PATH_ELLIPSIS | DT_EXPANDTABS | DT_CALCRECT);
                        rc.top += src.bottom - src.top;

                        pStr += len + 1;
                    }
                }
                break;
            }
            default:
            {
                if (m_text.empty())
                {
                    m_text += TEXT("Unknown format: ");
                    m_text += GetFormatName(m_uFormat);
                    const SIZE_T sz = GlobalSize(hData);
                    auto pData = AutoGlobalLock<BYTE*>(hData);
                    for (SIZE_T i = 0; i < sz; i += 16)
                    {
                        m_text += TEXT('\n');
                        for (SIZE_T j = 0; j < 16 && (i + j) < sz; ++j)
                        {
                            const BYTE b = pData.get()[i + j];
                            TCHAR tmp[100];
                            CHECK_HR(StringCchPrintf(tmp, ARRAYSIZE(tmp), TEXT(" %02X"), b));
                            m_text += tmp;
                        }
                        m_text += TEXT('\t');
                        for (SIZE_T j = 0; j < 16 && (i + j) < sz; ++j)
                        {
                            const BYTE b = pData.get()[i + j];
                            m_text += std::isprint(b) ? b : TEXT('.');
                        }
                    }
                }
                SelectObject(pps->hdc, GetStockObject(SYSTEM_FIXED_FONT));
                DrawText(pps->hdc, m_text.data(), int(m_text.length()), &rc, DT_TOP | DT_LEFT | DT_EXPANDTABS);
                break;
            }
            }
        }
        CHECK_LE(CloseClipboard());
    }
}

bool Run(_In_ const LPCTSTR lpCmdLine, _In_ const int nShowCmd)
{
    RadLogInitWnd(NULL, "Rad Clipboard", L"Rad Clipboard");

    if (RadClipboardViewerWnd::Register() == 0)
    {
        RadLog(LOG_ERROR, TEXT("Error registering window class"), SRC_LOC);
        return false;
    }
    RadClipboardViewerWnd* prw = RadClipboardViewerWnd::Create();
    if (prw == nullptr)
    {
        RadLog(LOG_ERROR, TEXT("Error creating root window"), SRC_LOC);
        return false;
    }

    RadLogInitWnd(*prw, "Rad Clipboard", L"Rad Clipboard");

    ShowWindow(*prw, nShowCmd);
    return true;
}

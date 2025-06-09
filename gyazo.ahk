#Requires AutoHotkey v2.0
#Include ImagePut.ahk

global gdipToken := 0

SetTimer(WatchClipboard, 1000)
return

WatchClipboard() {
    static lastClip := ""

    if !ClipWait(1)
        return

    clip := A_Clipboard
    m := []

    if (clip != lastClip && RegExMatch(clip, "^https://gyazo\.com/([a-zA-Z0-9]+)", &m)) {
        lastClip := clip
        id := m[1]
        imageURL := "https://i.gyazo.com/" id ".png"

        try {
            ImagePutClipboard(imageURL)
        } catch as e {
            MsgBox("Failed to copy image to clipboard: " e.Message)
        }
    }
}

DownloadFile(URL, SaveTo) {
    if !DirExist(A_Temp)
        DirCreate(A_Temp)
    http := ComObject("WinHttp.WinHttpRequest.5.1")
    http.Open("GET", URL, false)
    http.Send()
    if (http.Status != 200)
        return false

    stream := ComObject("ADODB.Stream")
    stream.Type := 1  ; Binary
    stream.Open()
    stream.Write(http.ResponseBody)
    try {
        stream.SaveToFile(SaveTo, 2)  ; Overwrite
    } catch as e {
        MsgBox("Failed to save file: " SaveTo "`nError: " e.Message)
        stream.Close()
        return false
    }
    stream.Close()
    return true
}

CopyImageToClipboard(FilePath) {
    Gdip_Startup()
    hBitmap := LoadImageAsBitmap(FilePath)
    if !hBitmap {
        MsgBox("Failed to load image as bitmap.")
        return false
    }

    if !OpenClipboard(0) {
        MsgBox("Failed to open clipboard.")
        DeleteObject(hBitmap)
        return false
    }
    try {
        EmptyClipboard()
        hDIB := BitmapToDIB(hBitmap)
        if hDIB {
            SetClipboardData(8, hDIB) ; CF_DIB = 8
            result := true
        } else {
            MsgBox("Failed to convert bitmap to DIB. The image may not be 24/32bpp or is not supported.")
            result := false
        }
    } finally {
        CloseClipboard()
        DeleteObject(hBitmap)
    }
    return result
}

; Helper: Load image file as HBITMAP
LoadImageAsBitmap(FilePath) {
    pBitmap := 0
    hr := DllCall("gdiplus\GdipCreateBitmapFromFile", "WStr", FilePath, "Ptr*", &pBitmap)
    if hr != 0 || !pBitmap
        return 0
    hBitmap := 0
    hr := DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "Ptr*", &hBitmap, "UInt", 0)
    DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
    if hr != 0
        return 0
    return hBitmap
}

; Helper: Convert HBITMAP to DIB section (returns handle to DIB)
BitmapToDIB(hBitmap) {
    bi := Buffer(40, 0)
    NumPut(40, bi, 0, "UInt")
    DllCall("gdi32\GetObjectW", "Ptr", hBitmap, "Int", 40, "Ptr", bi.Ptr)
    width := NumGet(bi, 4, "Int")
    height := NumGet(bi, 8, "Int")
    bits := NumGet(bi, 18, "UShort")
    if (bits != 24 && bits != 32)
        return 0

    bi2 := Buffer(40, 0)
    NumPut(40, bi2, 0, "UInt")
    NumPut(width, bi2, 4, "Int")
    NumPut(height, bi2, 8, "Int")
    NumPut(1, bi2, 12, "UShort")
    NumPut(bits, bi2, 14, "UShort")
    NumPut(0, bi2, 16, "UInt")
    hdc := DllCall("user32\GetDC", "Ptr", 0, "Ptr")
    pBits := 0
    hDIB := DllCall("gdi32\CreateDIBSection", "Ptr", hdc, "Ptr", bi2.Ptr, "UInt", 0, "Ptr*", &pBits, "Ptr", 0, "UInt",
        0, "Ptr")
    DllCall("user32\ReleaseDC", "Ptr", 0, "Ptr", hdc)
    if !hDIB
        return 0

    hdcSrc := DllCall("gdi32\CreateCompatibleDC", "Ptr", 0, "Ptr")
    hdcDst := DllCall("gdi32\CreateCompatibleDC", "Ptr", 0, "Ptr")
    obmSrc := DllCall("gdi32\SelectObject", "Ptr", hdcSrc, "Ptr", hBitmap, "Ptr")
    obmDst := DllCall("gdi32\SelectObject", "Ptr", hdcDst, "Ptr", hDIB, "Ptr")
    DllCall("gdi32\BitBlt", "Ptr", hdcDst, "Int", 0, "Int", 0, "Int", width, "Int", height, "Ptr", hdcSrc, "Int", 0,
        "Int", 0, "UInt", 0x00CC0020)
    DllCall("gdi32\SelectObject", "Ptr", hdcSrc, "Ptr", obmSrc)
    DllCall("gdi32\SelectObject", "Ptr", hdcDst, "Ptr", obmDst)
    DllCall("gdi32\DeleteDC", "Ptr", hdcSrc)
    DllCall("gdi32\DeleteDC", "Ptr", hdcDst)
    return hDIB
}

; Helper: Delete GDI object
DeleteObject(hObj) {
    return DllCall("gdi32\DeleteObject", "Ptr", hObj)
}

Gdip_Startup() {
    global gdipToken
    if gdipToken
        return gdipToken

    GdiplusStartupInput := Buffer(16, 0)
    NumPut("UInt", 1, GdiplusStartupInput)
    DllCall("gdiplus\GdiplusStartup", "Ptr*", &gdipToken, "Ptr", GdiplusStartupInput, "Ptr", 0)
    return gdipToken
}

Gdip_Shutdown(pToken) {
    static shutdownTokens := Map()
    if pToken && !shutdownTokens.Has(pToken) {
        DllCall("gdiplus\GdiplusShutdown", "Ptr", pToken)
        shutdownTokens[pToken] := true
    }
}

Gdip_CreateBitmapFromFile(FilePath) {
    pBitmap := 0
    hr := DllCall("gdiplus\GdipCreateBitmapFromFile", "WStr", FilePath, "Ptr*", &pBitmap)
    if hr != 0
        return 0
    return pBitmap
}

Gdip_GetHBITMAPFromBitmap(pBitmap) {
    hBitmap := 0
    hr := DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "Ptr*", &hBitmap, "UInt", 0)
    if hr != 0
        return 0
    return hBitmap
}

Gdip_DisposeImage(pBitmap) {
    if pBitmap
        DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
}

OpenClipboard(hWnd := 0) => DllCall("user32\OpenClipboard", "Ptr", hWnd)
EmptyClipboard() => DllCall("user32\EmptyClipboard")
SetClipboardData(f, h) => DllCall("user32\SetClipboardData", "UInt", f, "Ptr", h)
CloseClipboard() => DllCall("user32\CloseClipboard", "Ptr")

; Ensure cleanup on exit
OnExit(ShutdownGDI)

ShutdownGDI(*) {
    global gdipToken
    Gdip_Shutdown(gdipToken)
}

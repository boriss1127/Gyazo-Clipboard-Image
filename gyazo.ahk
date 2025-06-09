#Requires AutoHotkey v2.0
#Include ImagePut.ahk

global gdipToken := 0

; Set timer once, no Return after
SetTimer(WatchClipboard, 35)

WatchClipboard() {
    static lastClip := ""

    clip := A_Clipboard
    if (clip = lastClip)
        return

    m := []
    if (RegExMatch(clip, "^https://gyazo\.com/([a-zA-Z0-9]+)", &m)) {
        lastClip := clip
        id := m[1]
        imageURL := "https://i.gyazo.com/" id ".png"

        ; Try multiple approaches simultaneously
        ; Approach 1: Try immediately
        SetTimer(() => TryGetImage(imageURL), -1)

        ; Approach 2: Try after very short delay
        SetTimer(() => TryGetImage(imageURL), -100)

        ; Approach 3: Try after slightly longer delay
        SetTimer(() => TryGetImage(imageURL), -250)

        ; Approach 4: Try after moderate delay
        SetTimer(() => TryGetImage(imageURL), -500)

        ; Approach 5: Final attempt after 1 second
        SetTimer(() => TryGetImage(imageURL), -1000)
    }
}

TryGetImage(imageURL) {
    static successfulURLs := Map()

    ; Skip if we already successfully processed this URL
    if (successfulURLs.Has(imageURL))
        return

    try {
        ; Try with different timeout settings for ImagePut
        ImagePutClipboard(imageURL)
        successfulURLs[imageURL] := true
        return  ; Success - stop trying
    } catch Error as e {
        ; If ImagePut failed, try alternative methods
        TryAlternativeMethod(imageURL)
    }
}

TryAlternativeMethod(imageURL) {
    static successfulURLs := Map()

    if (successfulURLs.Has(imageURL))
        return

    try {
        ; Method 1: Download to temp file first, then copy to clipboard
        tempFile := A_Temp . "\gyazo_temp_" . A_TickCount . ".png"

        if (DownloadFile(imageURL, tempFile)) {
            if (CopyImageToClipboard(tempFile)) {
                successfulURLs[imageURL] := true
                FileDelete(tempFile)  ; Clean up
                return
            }
            FileDelete(tempFile)  ; Clean up on failure
        }
    } catch {
        ; Method 2: Try with different HTTP approach
        TryHTTPMethod(imageURL)
    }
}

TryHTTPMethod(imageURL) {
    static successfulURLs := Map()

    if (successfulURLs.Has(imageURL))
        return

    try {
        ; Create HTTP request with specific headers that might help
        http := ComObject("WinHttp.WinHttpRequest.5.1")
        http.Open("GET", imageURL, false)

        ; Set headers to mimic a browser request
        http.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
        http.SetRequestHeader("Accept", "image/png,image/*,*/*;q=0.8")
        http.SetRequestHeader("Cache-Control", "no-cache")

        ; Short timeout for quick attempts
        http.SetTimeouts(2000, 2000, 3000, 5000)

        http.Send()

        if (http.Status = 200) {
            ; Save to temp file and copy to clipboard
            tempFile := A_Temp . "\gyazo_http_" . A_TickCount . ".png"

            stream := ComObject("ADODB.Stream")
            stream.Type := 1  ; Binary
            stream.Open()
            stream.Write(http.ResponseBody)
            stream.SaveToFile(tempFile, 2)  ; Overwrite
            stream.Close()

            if (CopyImageToClipboard(tempFile)) {
                successfulURLs[imageURL] := true
                FileDelete(tempFile)
                return
            }
            FileDelete(tempFile)
        }
    } catch {
        ; All methods failed, but don't show errors since we're trying multiple times
    }
}

DownloadFile(URL, SaveTo) {
    try {
        if !DirExist(A_Temp)
            DirCreate(A_Temp)

        http := ComObject("WinHttp.WinHttpRequest.5.1")
        http.Open("GET", URL, false)

        ; Set aggressive timeouts for quick attempts
        http.SetTimeouts(1000, 1000, 2000, 3000)

        http.Send()
        if (http.Status != 200)
            return false

        stream := ComObject("ADODB.Stream")
        stream.Type := 1  ; Binary
        stream.Open()
        stream.Write(http.ResponseBody)
        stream.SaveToFile(SaveTo, 2)  ; Overwrite
        stream.Close()
        return true
    } catch {
        return false
    }
}

CopyImageToClipboard(FilePath) {
    try {
        Gdip_Startup()
        hBitmap := LoadImageAsBitmap(FilePath)
        if !hBitmap {
            return false
        }

        if !OpenClipboard(0) {
            DeleteObject(hBitmap)
            return false
        }

        EmptyClipboard()
        hDIB := BitmapToDIB(hBitmap)
        if hDIB {
            SetClipboardData(8, hDIB) ; CF_DIB = 8
            result := true
        } else {
            result := false
        }

        CloseClipboard()
        DeleteObject(hBitmap)
        return result
    } catch {
        return false
    }
}

; Helper: Load image file as HBITMAP
LoadImageAsBitmap(FilePath) {
    try {
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
    } catch {
        return 0
    }
}

; Helper: Convert HBITMAP to DIB section (returns handle to DIB)
BitmapToDIB(hBitmap) {
    try {
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
        hDIB := DllCall("gdi32\CreateDIBSection", "Ptr", hdc, "Ptr", bi2.Ptr, "UInt", 0, "Ptr*", &pBits, "Ptr", 0,
            "UInt",
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
    } catch {
        return 0
    }
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

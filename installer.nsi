; ─────────────────────────────────────────────────────────────────────────────
; QA Screenshot Tool Portable — PortableApps.com Format Installer
; Built with NSIS (Nullsoft Scriptable Install System)
; RequestExecutionLevel user  →  NO admin rights required
; ─────────────────────────────────────────────────────────────────────────────

Unicode true

!define APP_NAME    "QA Screenshot Tool Portable"
!define APP_ID      "QAScreenshotToolPortable"
!define APP_VERSION "1.0"
!define PUBLISHER   "EmilsinghSamuel"

Name    "${APP_NAME} ${APP_VERSION}"
OutFile "${APP_ID}_${APP_VERSION}.paf.exe"

; No admin needed — fully portable
RequestExecutionLevel user

; Solid LZMA gives smallest file
SetCompressor /SOLID lzma

; Default install location: user's Documents\PortableApps folder
InstallDir "$DOCUMENTS\PortableApps\${APP_ID}"

; Pages
Page directory
Page instfiles

; ── Install section ───────────────────────────────────────────────────────────
Section "Install" SEC_MAIN

    ; Launcher in root (PortableApps Platform looks here)
    SetOutPath "$INSTDIR"
    File "dist\${APP_ID}.exe"

    ; Main application exe
    SetOutPath "$INSTDIR\App\QAScreenshotTool"
    File "dist\QA_Screenshot_Tool.exe"

    ; App metadata (PortableApps Platform reads this for its menu)
    SetOutPath "$INSTDIR\App\AppInfo"
    File "App\AppInfo\appinfo.ini"

    ; Ensure data directory exists on first install
    CreateDirectory "$INSTDIR\Data\settings"

    ; Write uninstaller
    WriteUninstaller "$INSTDIR\Uninstall.exe"

SectionEnd

; ── Uninstall section ────────────────────────────────────────────────────────
Section "Uninstall"

    Delete "$INSTDIR\${APP_ID}.exe"
    Delete "$INSTDIR\Uninstall.exe"
    RMDir /r "$INSTDIR\App"
    ; Data directory is intentionally NOT removed — preserve tester's settings
    RMDir "$INSTDIR"   ; removes root only if empty

SectionEnd

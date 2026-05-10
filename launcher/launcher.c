/*
 * QAScreenshotToolPortable.exe — PortableApps launcher
 *
 * Compiled with GCC (no runtime dependency).
 * Sets PORTABLE_DATA_DIR so the app stores config.json in
 * Data\settings\ (inside the portable root) instead of next to its own exe.
 * Then launches the real app and waits for it to exit.
 */
#include <windows.h>

int WINAPI WinMain(HINSTANCE hInst, HINSTANCE hPrev, LPSTR lpCmd, int nShow)
{
    char root[MAX_PATH];   /* portable app root  (folder of this launcher) */
    char exe[MAX_PATH];    /* path to main exe   */
    char data[MAX_PATH];   /* path to Data\settings\ */

    /* Resolve the launcher's own directory as the portable root */
    GetModuleFileNameA(NULL, root, MAX_PATH);
    *strrchr(root, '\\') = '\0';

    /* Let the Python app know where to write config.json */
    snprintf(data, MAX_PATH, "%s\\Data\\settings", root);
    SetEnvironmentVariableA("PORTABLE_DATA_DIR", data);
    CreateDirectoryA(data, NULL);   /* create on first run if absent */

    /* Path to the real executable */
    snprintf(exe, MAX_PATH,
             "%s\\App\\QAScreenshotTool\\QA_Screenshot_Tool.exe", root);

    STARTUPINFOA        si = { sizeof(si) };
    PROCESS_INFORMATION pi;

    if (!CreateProcessA(exe, NULL, NULL, NULL, FALSE, 0, NULL, root, &si, &pi)) {
        char msg[600];
        snprintf(msg, sizeof(msg),
                 "Could not start the application.\n\nExpected:\n%s\n\nError code: %lu",
                 exe, GetLastError());
        MessageBoxA(NULL, msg, "QA Screenshot Tool — Launch Error", MB_ICONERROR);
        return 1;
    }

    WaitForSingleObject(pi.hProcess, INFINITE);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);
    return 0;
}

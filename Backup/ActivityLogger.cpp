// ActivityLogger.cpp
#include <windows.h>
#include <psapi.h>
#include <tlhelp32.h>
#include <shlobj.h>
#include <shellapi.h>
#include <commctrl.h>
#include <commdlg.h>
#include <fstream>
#include <sstream>
#include <string>
#include <vector>
#include <thread>
#include <chrono>
#include <mutex>
#include <map>
#include <memory>
#include <iostream>
#include <iomanip>
#include <algorithm>

#pragma comment(lib, "user32.lib")
#pragma comment(lib, "psapi.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "comdlg32.lib")

class ActivityLogger {
private:
    std::string logPath;
    bool running;
    std::thread loggerThread;
    std::mutex dataMutex;
    
    // Window tracking
    std::string prevWindow;
    std::string prevProcess;
    std::string prevDetails;
    std::string prevCategory;
    std::chrono::system_clock::time_point startTime;
    std::chrono::system_clock::time_point appStartTime;
    
    // Idle detection
    bool wasIdle;
    std::chrono::system_clock::time_point idleStart;
    
    // Tray icon
    NOTIFYICONDATA nid;
    HWND hwnd;
    HMENU hMenu;
    
    // Viewer window
    HWND viewerHwnd;
    bool viewerOpen;

public:
    ActivityLogger() : running(false), wasIdle(false), viewerOpen(false), viewerHwnd(nullptr) {
        appStartTime = std::chrono::system_clock::now();
        logPath = getLogPath();
    }

    ~ActivityLogger() {
        stop();
        if (viewerOpen && viewerHwnd) {
            DestroyWindow(viewerHwnd);
        }
    }

    std::string getLogPath() {
        char path[MAX_PATH];
        char computerName[MAX_COMPUTERNAME_LENGTH + 1];
        DWORD size = MAX_COMPUTERNAME_LENGTH + 1;
        GetComputerNameA(computerName, &size);
        
        // Try OneDrive paths first
        std::vector<std::string> onedrivePaths = {
            "\\OneDrive\\Documents",
            "\\OneDrive - Personal\\Documents", 
            "\\OneDrive - GE HealthCare\\Documents"
        };
        
        if (SHGetFolderPathA(NULL, CSIDL_PROFILE, NULL, 0, path) == S_OK) {
            for (const auto& onedrivePath : onedrivePaths) {
                std::string fullPath = std::string(path) + onedrivePath + "\\ActivityLogger";
                if (GetFileAttributesA(fullPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
                    CreateDirectoryA(fullPath.c_str(), NULL);
                    return fullPath + "\\" + std::string(computerName) + "_ActivityLog.csv";
                }
            }
        }
        
        // Fallback to AppData\Local
        if (SHGetFolderPathA(NULL, CSIDL_LOCAL_APPDATA, NULL, 0, path) == S_OK) {
            std::string appDataPath = std::string(path) + "\\ActivityLogger";
            CreateDirectoryA(appDataPath.c_str(), NULL);
            return appDataPath + "\\" + std::string(computerName) + "_ActivityLog.csv";
        }
        
        return "ActivityLog.csv";
    }

    std::string getActiveWindowTitle() {
        HWND hwnd = GetForegroundWindow();
        if (!hwnd) return "";
        
        char title[256];
        GetWindowTextA(hwnd, title, sizeof(title));
        return std::string(title);
    }

    std::string getActiveProcessName() {
        HWND hwnd = GetForegroundWindow();
        if (!hwnd) return "";
        
        DWORD processId;
        GetWindowThreadProcessId(hwnd, &processId);
        
        HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, processId);
        if (!hProcess) return "";
        
        char processName[MAX_PATH];
        DWORD size = MAX_PATH;
        if (QueryFullProcessImageNameA(hProcess, 0, processName, &size)) {
            CloseHandle(hProcess);
            std::string fullPath(processName);
            size_t pos = fullPath.find_last_of("\\/");
            return (pos != std::string::npos) ? fullPath.substr(pos + 1) : fullPath;
        }
        
        CloseHandle(hProcess);
        return "";
    }

    std::string getWindowDetails(const std::string& windowTitle, const std::string& processName) {
        std::string title = windowTitle;
        std::string proc = processName;
        std::transform(proc.begin(), proc.end(), proc.begin(), ::tolower);
        
        // Remove common application suffixes
        if (proc == "excel.exe" && title.find(" - Excel") != std::string::npos) {
            size_t pos = title.find(" - Excel");
            return title.substr(0, pos);
        } else if (proc == "winword.exe" && title.find(" - Word") != std::string::npos) {
            size_t pos = title.find(" - Word");
            return title.substr(0, pos);
        } else if (proc == "chrome.exe" && title.find(" - Google Chrome") != std::string::npos) {
            size_t pos = title.find(" - Google Chrome");
            return title.substr(0, pos);
        } else if (title.find(" - ") != std::string::npos) {
            size_t pos = title.rfind(" - ");
            return title.substr(0, pos);
        }
        
        return title;
    }

    std::string getCategory(const std::string& windowTitle, const std::string& processName, const std::string& windowDetails) {
        std::string proc = processName;
        std::string title = windowTitle;
        std::string details = windowDetails;
        
        std::transform(proc.begin(), proc.end(), proc.begin(), ::tolower);
        std::transform(title.begin(), title.end(), title.begin(), ::tolower);
        std::transform(details.begin(), details.end(), details.begin(), ::tolower);
        
        std::map<std::string, std::string> categories = {
            {"excel", "Work - Office"},
            {"winword", "Work - Office"},
            {"powerpnt", "Work - Office"},
            {"outlook", "Email"},
            {"chrome", "Web Browsing"},
            {"firefox", "Web Browsing"},
            {"msedge", "Web Browsing"},
            {"teams", "Meetings"},
            {"slack", "Communication"},
            {"zoom", "Meetings"},
            {"notepad", "Notes"},
            {"code", "Development"},
            {"cmd", "Terminal"},
            {"powershell", "Terminal"}
        };
        
        for (const auto& pair : categories) {
            if (proc.find(pair.first) != std::string::npos ||
                title.find(pair.first) != std::string::npos ||
                details.find(pair.first) != std::string::npos) {
                return pair.second;
            }
        }
        
        return "Uncategorized";
    }

    int getIdleSeconds() {
        LASTINPUTINFO lii;
        lii.cbSize = sizeof(LASTINPUTINFO);
        if (GetLastInputInfo(&lii)) {
            DWORD tickCount = GetTickCount();
            return (tickCount - lii.dwTime) / 1000;
        }
        return 0;
    }

    void logActivity(const std::chrono::system_clock::time_point& start,
                    const std::chrono::system_clock::time_point& end,
                    const std::string& window,
                    const std::string& process,
                    const std::string& details,
                    const std::string& category) {
        
        std::lock_guard<std::mutex> lock(dataMutex);
        
        auto duration = std::chrono::duration_cast<std::chrono::seconds>(end - start).count();
        if (duration <= 0) return;
        
        bool fileExists = (GetFileAttributesA(logPath.c_str()) != INVALID_FILE_ATTRIBUTES);
        std::ofstream file(logPath, std::ios::app);
        
        if (!file.is_open()) return;
        
        if (!fileExists) {
            file << "StartTime,EndTime,DurationSeconds,WindowTitle,WindowDetails,ProcessName,Category\n";
        }
        
        auto startTime_t = std::chrono::system_clock::to_time_t(start);
        auto endTime_t = std::chrono::system_clock::to_time_t(end);
        
        struct tm startTm, endTm;
        localtime_s(&startTm, &startTime_t);
        localtime_s(&endTm, &endTime_t);
        
        file << std::put_time(&startTm, "%Y-%m-%d %H:%M:%S") << ","
             << std::put_time(&endTm, "%Y-%m-%d %H:%M:%S") << ","
             << duration << ","
             << "\"" << window << "\","
             << "\"" << details << "\","
             << "\"" << process << "\","
             << "\"" << category << "\"\n";
        
        file.close();
    }

    void pollingLoop() {
        std::cout << "Starting polling method\n";
        
        startTime = std::chrono::system_clock::now();
        prevWindow = getActiveWindowTitle();
        prevProcess = getActiveProcessName();
        prevDetails = getWindowDetails(prevWindow, prevProcess);
        prevCategory = getCategory(prevWindow, prevProcess, prevDetails);
        wasIdle = false;
        
        int idleCheckCounter = 0;
        const int idleCheckFrequency = 10; // Check idle every 5 seconds (10 * 0.5s)
        
        while (running) {
            try {
                // Get current window info
                std::string currentWindow = getActiveWindowTitle();
                std::string currentProcess = getActiveProcessName();
                std::string currentDetails = getWindowDetails(currentWindow, currentProcess);
                std::string currentCategory = getCategory(currentWindow, currentProcess, currentDetails);
                
                // Check if window changed
                if (currentWindow != prevWindow || 
                    currentDetails != prevDetails || 
                    currentCategory != prevCategory) {
                    
                    // Log previous activity
                    if (!prevWindow.empty() && !wasIdle) {
                        auto endTime = std::chrono::system_clock::now();
                        auto duration = std::chrono::duration_cast<std::chrono::seconds>(endTime - startTime).count();
                        if (duration >= 1) {
                            logActivity(startTime, endTime, prevWindow, prevProcess, prevDetails, prevCategory);
                        }
                    }
                    
                    // Update to new window
                    prevWindow = currentWindow;
                    prevProcess = currentProcess;
                    prevDetails = currentDetails;
                    prevCategory = currentCategory;
                    startTime = std::chrono::system_clock::now();
                }
                
                // Check idle status periodically
                idleCheckCounter++;
                if (idleCheckCounter >= idleCheckFrequency) {
                    idleCheckCounter = 0;
                    
                    int idleSeconds = getIdleSeconds();
                    int idleThreshold = (prevCategory == "Meetings") ? 3600 : 300; // 1 hour for meetings, 5 min for others
                    bool isIdle = idleSeconds >= idleThreshold;
                    
                    if (isIdle && !wasIdle) {
                        // Just went idle
                        idleStart = std::chrono::system_clock::now();
                        
                        // Log current activity before going idle
                        if (!prevWindow.empty()) {
                            auto duration = std::chrono::duration_cast<std::chrono::seconds>(idleStart - startTime).count();
                            if (duration >= 1) {
                                logActivity(startTime, idleStart, prevWindow, prevProcess, prevDetails, prevCategory);
                            }
                        }
                        wasIdle = true;
                        
                    } else if (!isIdle && wasIdle) {
                        // Just became active
                        auto now = std::chrono::system_clock::now();
                        auto idleDuration = std::chrono::duration_cast<std::chrono::seconds>(now - idleStart).count();
                        
                        if (idleDuration >= 300) { // Log idle periods longer than 5 minutes
                            logActivity(idleStart, now, "Inactive", "", "", "Inactive");
                        }
                        
                        // Reset for new activity
                        startTime = now;
                        wasIdle = false;
                        prevWindow = currentWindow;
                        prevProcess = currentProcess;
                        prevDetails = currentDetails;
                        prevCategory = currentCategory;
                    }
                }
                
                std::this_thread::sleep_for(std::chrono::milliseconds(500));
                
            } catch (const std::exception& e) {
                std::cerr << "Error in polling loop: " << e.what() << std::endl;
                std::this_thread::sleep_for(std::chrono::milliseconds(500));
            }
        }
    }

    void start() {
        if (!running) {
            running = true;
            loggerThread = std::thread(&ActivityLogger::pollingLoop, this);
        }
    }

    void stop() {
        if (running) {
            running = false;
            if (loggerThread.joinable()) {
                loggerThread.join();
            }
        }
    }

    void restart() {
        stop();
        std::this_thread::sleep_for(std::chrono::milliseconds(500));
        start();
    }

    void openLogFolder() {
        std::string folder = logPath.substr(0, logPath.find_last_of("\\/"));
        ShellExecuteA(NULL, "open", folder.c_str(), NULL, NULL, SW_SHOWNORMAL);
    }

    void openLogViewer() {
        if (!viewerOpen) {
            createLogViewer();
        } else if (viewerHwnd) {
            SetForegroundWindow(viewerHwnd);
        }
    }

    void createLogViewer() {
        // This would require a full Windows GUI implementation
        // For brevity, just open the CSV file in default application
        ShellExecuteA(NULL, "open", logPath.c_str(), NULL, NULL, SW_SHOWNORMAL);
    }

    void showHelp() {
        std::string helpText = 
            "Activity Logger - Help\n\n"
            "MENU ITEMS:\n\n"
            "Start Logging: Begins monitoring active windows\n"
            "Stop Logging: Pauses activity monitoring\n" 
            "Restart Logging: Stops and restarts logging\n"
            "Open Log File: Opens the activity log\n"
            "Open Folder: Opens log file location\n"
            "Help: Shows this help\n"
            "Exit: Closes the application\n\n"
            "Log Location: " + logPath + "\n\n"
            "The CSV can be analyzed with Excel or Power BI.";
            
        MessageBoxA(NULL, helpText.c_str(), "Activity Logger Help", MB_OK | MB_ICONINFORMATION);
    }

    // Tray icon management
    void createTrayIcon(HWND hwnd) {
        this->hwnd = hwnd;
        
        ZeroMemory(&nid, sizeof(NOTIFYICONDATA));
        nid.cbSize = sizeof(NOTIFYICONDATA);
        nid.hWnd = hwnd;
        nid.uID = 1;
        nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
        nid.uCallbackMessage = WM_USER + 1;
        nid.hIcon = LoadIcon(NULL, IDI_APPLICATION);
        
        // Fix: Use lstrcpynA for ANSI or lstrcpynW for Unicode
        lstrcpynW(nid.szTip, L"Activity Logger", sizeof(nid.szTip) / sizeof(wchar_t));
        
        Shell_NotifyIcon(NIM_ADD, &nid);
        
        // Create context menu - use ANSI versions
        hMenu = CreatePopupMenu();
        AppendMenuA(hMenu, MF_STRING, 1001, "Start Logging");
        AppendMenuA(hMenu, MF_STRING, 1002, "Stop Logging"); 
        AppendMenuA(hMenu, MF_STRING, 1003, "Restart Logging");
        AppendMenuA(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenuA(hMenu, MF_STRING, 1004, "Open Log File");
        AppendMenuA(hMenu, MF_STRING, 1005, "Open Folder");
        AppendMenuA(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenuA(hMenu, MF_STRING, 1006, "Help");
        AppendMenuA(hMenu, MF_STRING, 1007, "Exit");
    }

    void destroyTrayIcon() {
        Shell_NotifyIcon(NIM_DELETE, &nid);
        if (hMenu) {
            DestroyMenu(hMenu);
        }
    }

    void handleTrayMessage(WPARAM wParam, LPARAM lParam) {
        if (lParam == WM_RBUTTONUP) {
            POINT pt;
            GetCursorPos(&pt);
            SetForegroundWindow(hwnd);
            
            // Update menu items based on current state
            EnableMenuItem(hMenu, 1001, running ? MF_GRAYED : MF_ENABLED);
            EnableMenuItem(hMenu, 1002, running ? MF_ENABLED : MF_GRAYED);
            
            TrackPopupMenu(hMenu, TPM_RIGHTBUTTON, pt.x, pt.y, 0, hwnd, NULL);
            PostMessage(hwnd, WM_NULL, 0, 0);
        }
    }

    void handleMenuCommand(WPARAM wParam) {
        switch (LOWORD(wParam)) {
            case 1001: start(); break;
            case 1002: stop(); break;
            case 1003: restart(); break;
            case 1004: openLogViewer(); break;
            case 1005: openLogFolder(); break;
            case 1006: showHelp(); break;
            case 1007: 
                stop();
                PostQuitMessage(0);
                break;
        }
    }

    bool isRunning() const { return running; }
};

// Global instance
std::unique_ptr<ActivityLogger> g_logger;

// Window procedure
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
        case WM_USER + 1: // Tray icon message
            if (g_logger) {
                g_logger->handleTrayMessage(wParam, lParam);
            }
            return 0;
            
        case WM_COMMAND:
            if (g_logger) {
                g_logger->handleMenuCommand(wParam);
            }
            return 0;
            
        case WM_DESTROY:
            if (g_logger) {
                g_logger->destroyTrayIcon();
                g_logger->stop();
            }
            PostQuitMessage(0);
            return 0;
    }
    
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    // Register window class
    WNDCLASSEX wc = {};
    wc.cbSize = sizeof(WNDCLASSEX);
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"ActivityLoggerClass";  // Use wide string for Unicode
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    
    RegisterClassEx(&wc);  // Use ANSI version
    
    // Create hidden window for message handling
    HWND hwnd = CreateWindowEx(0, L"ActivityLoggerClass", L"Activity Logger", 
                               0, 0, 0, 0, 0, NULL, NULL, hInstance, NULL);
    
    if (!hwnd) {
        MessageBoxW(NULL, L"Failed to create window", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }
    
    // Create logger instance
    g_logger = std::make_unique<ActivityLogger>();
    g_logger->createTrayIcon(hwnd);
    g_logger->start();
    
    // Message loop
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    
    // Cleanup
    g_logger.reset();
    
    return 0;
}
#include "MainWindow.h"
#include "ConfigManager.h"
#include "Resource.h"
#include "Localization.h"
#include <tchar.h>
#include <string>
#include <algorithm>
#include <commctrl.h>
#pragma comment(lib, "comctl32.lib")

MainWindow::MainWindow() : m_hwnd(NULL)
{
}

MainWindow::~MainWindow()
{
}

bool MainWindow::Create(HINSTANCE hInstance, int nCmdShow)
{
    WNDCLASSEX wcex = {0};
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.style          = CS_HREDRAW | CS_VREDRAW | CS_DBLCLKS;
    wcex.lpfnWndProc    = MainWindow::WindowProc;
    wcex.hInstance      = hInstance;
    wcex.hCursor        = LoadCursor(NULL, IDC_ARROW);
    wcex.lpszClassName  = _T("CSVEditorWindow");

    RegisterClassEx(&wcex);

    // Load Config
    m_rowHeight = ConfigManager::Instance().GetFloat(L"Layout", L"RowHeight", 20.0f);
    m_defaultColWidth = ConfigManager::Instance().GetFloat(L"Layout", L"DefaultColWidth", 100.0f);
    m_maxCellLines = ConfigManager::Instance().GetInt(L"Layout", L"MaxCellLines", 1);
    // Apply m_maxCellLines to row height if desired?
    // For now independent.

    m_hwnd = CreateWindow(
        _T("CSVEditorWindow"),
        _T("CSV Editor"),
        WS_OVERLAPPEDWINDOW | WS_VSCROLL | WS_HSCROLL,
        CW_USEDEFAULT, CW_USEDEFAULT,
        1024, 768,
        NULL,
        NULL,
        hInstance,
        this
    );

    if (!m_hwnd) return false;
    
    // Formula Bar
    // Label "A1"
    m_hPosLabel = CreateWindow(
        _T("STATIC"), _T(""),
        WS_CHILD | WS_VISIBLE | SS_CENTER | SS_CENTERIMAGE, 
        0, 0, 40, (int)m_formulaBarHeight,
        m_hwnd, NULL, hInstance, NULL
    );
    SendMessage(m_hPosLabel, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);

    // Edit Field
    m_hFormulaEdit = CreateWindow(
        _T("EDIT"),
        NULL,
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
        40, 0, 100, (int)m_formulaBarHeight, // x starts after label
        m_hwnd,
        (HMENU)ID_FORMULA_EDIT,
        hInstance,
        NULL
    );
    // Set font for edit?
    SendMessage(m_hFormulaEdit, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);

    // Progress Bar
    m_hProgressBar = CreateWindowEx(
        0, PROGRESS_CLASS, NULL,
        WS_CHILD | PBS_SMOOTH, // Not visible initially
        0, 0, 100, 20,
        m_hwnd, NULL, hInstance, NULL
    );
    SendMessage(m_hProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, 100));

    InitializeMenu(m_hwnd);

    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_WIN95_CLASSES | ICC_STANDARD_CLASSES;
    InitCommonControlsEx(&icex);

    
    if (FAILED(m_dxResources.CreateDeviceIndependentResources())) {
        return false;
    }
    
    m_hCursorArrow = LoadCursor(NULL, IDC_ARROW);
    m_hCursorWE = LoadCursor(NULL, IDC_SIZEWE);

    ShowWindow(m_hwnd, nCmdShow);
    UpdateWindow(m_hwnd);

    return true;
}

bool MainWindow::OpenFile(const std::wstring& filePath)
{
    if (m_document.Load(filePath)) {
        m_currentFilePath = filePath;
        SetWindowText(m_hwnd, filePath.c_str());
        m_state.SetScrollRow(0);
        UpdateScrollBars();
        InvalidateRect(m_hwnd, NULL, FALSE);
        return true;
    }
    return false;
}

LRESULT CALLBACK MainWindow::WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    MainWindow* pThis = NULL;

    if (uMsg == WM_NCCREATE) {
        CREATESTRUCT* pCreate = (CREATESTRUCT*)lParam;
    pThis = (MainWindow*)pCreate->lpCreateParams;
        SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR)pThis);
        pThis->m_hwnd = hwnd;
    } else {
        pThis = (MainWindow*)GetWindowLongPtr(hwnd, GWLP_USERDATA);
    } // No change here, just context

    if (pThis) {
        return pThis->HandleMessage(hwnd, uMsg, wParam, lParam);
    } else {
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
}

LRESULT MainWindow::HandleMessage(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg) {
    case WM_SIZE:
        OnResize(LOWORD(lParam), HIWORD(lParam));
        return 0;

    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hwnd, &ps);
            OnPaint(hwnd);
            EndPaint(hwnd, &ps);
        }
        return 0;

    case WM_COMMAND:
        OnCommand(hwnd, LOWORD(wParam), (HWND)lParam, HIWORD(wParam));
        return 0;

    case WM_LBUTTONDOWN:
        OnLButtonDown(LOWORD(lParam), HIWORD(lParam), (UINT)wParam);
        return 0;

    case WM_LBUTTONUP:
        OnLButtonUp(LOWORD(lParam), HIWORD(lParam), (UINT)wParam);
        return 0;

    case WM_CONTEXTMENU:
        OnContextMenu((HWND)wParam, LOWORD(lParam), HIWORD(lParam));
        return 0;

    case WM_LBUTTONDBLCLK:
        OnLButtonDblClk(LOWORD(lParam), HIWORD(lParam), (UINT)wParam);
        return 0;
    
    case WM_KEYDOWN:
        if (GetKeyState(VK_CONTROL) & 0x8000) {
            if (wParam == 'Z') {
                SendMessage(hwnd, WM_COMMAND, IDM_EDIT_UNDO, 0);
            } else if (wParam == 'Y') {
                SendMessage(hwnd, WM_COMMAND, IDM_EDIT_REDO, 0);
            } else if (wParam == 'S') { // Save
                SendMessage(hwnd, WM_COMMAND, IDM_FILE_SAVE, 0);
            } else if (wParam == 'F') { // Find
                m_showSearch = true;
                UpdateSearchBar();
            } else if (wParam == 'H') { // Replace
                OnSearchReplace();
            }
        } else {
             // Navigation
            bool shift = (GetKeyState(VK_SHIFT) & 0x8000);
            size_t r = 0, c = 0;
            auto sels = m_state.GetSelections();
            if(!sels.empty()) { r = sels[0].start.row; c = sels[0].start.col; }
            
            bool changed = false;
            if (wParam == VK_UP) {
                if(r > 0) { r--; changed = true; }
            } else if (wParam == VK_DOWN) {
                if(r + 1 < m_document.GetRowCount()) { r++; changed = true; }
            } else if (wParam == VK_LEFT) {
                if(c > 0) { c--; changed = true; }
            } else if (wParam == VK_RIGHT) {
                c++; changed = true; 
            }
            
            if (changed) {
                if (shift) m_state.DragTo(r, c);
                else m_state.SelectCell(r, c, false);
                
                UpdateFormulaBar();
                InvalidateRect(hwnd, NULL, FALSE);
                UpdateScrollBars();
            }
        }
        return 0;

    case WM_MOUSEMOVE:
        OnMouseMove(LOWORD(lParam), HIWORD(lParam), (UINT)wParam);
        return 0;

    case WM_VSCROLL:
        OnVScroll(wParam);
        return 0;

    case WM_HSCROLL:
        OnHScroll(wParam);
        return 0;

    case WM_MOUSEWHEEL:
        OnMouseWheel(GET_WHEEL_DELTA_WPARAM(wParam));
        return 0; 

    case WM_ERASEBKGND:
        return 1; // Prevent flickering

    case WM_DESTROY:
        ConfigManager::Instance().SetFloat(L"Layout", L"RowHeight", m_rowHeight);
        ConfigManager::Instance().SetFloat(L"Layout", L"DefaultColWidth", m_defaultColWidth);
        ConfigManager::Instance().SetInt(L"Layout", L"MaxCellLines", m_maxCellLines);
        ConfigManager::Instance().Save(); // Implicit in Set, but good for clarity if we had batching
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

void MainWindow::InitializeMenu(HWND hwnd)
{
    // Destroy old menu if exists
    HMENU hOldMenu = GetMenu(hwnd);
    if (hOldMenu) DestroyMenu(hOldMenu);

    HMENU hMenu = CreateMenu();
    
    // File Menu
    HMENU hFileMenu = CreatePopupMenu();
    AppendMenu(hFileMenu, MF_STRING, IDM_FILE_OPEN, Localization::GetString(StringId::Menu_File_Open));
    AppendMenu(hFileMenu, MF_STRING | MF_GRAYED, IDM_FILE_REOPEN, Localization::GetString(StringId::Menu_File_Reopen));
    AppendMenu(hFileMenu, MF_STRING, IDM_FILE_IMPORT, Localization::GetString(StringId::Menu_File_Import));
    AppendMenu(hFileMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hFileMenu, MF_STRING, IDM_FILE_SAVE, Localization::GetString(StringId::Menu_File_Save));
    AppendMenu(hFileMenu, MF_STRING, IDM_FILE_SAVEAS, Localization::GetString(StringId::Menu_File_SaveAs));
    AppendMenu(hFileMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hFileMenu, MF_STRING, IDM_FILE_EXPORT_HTML, Localization::GetString(StringId::Menu_File_Export_HTML));
    AppendMenu(hFileMenu, MF_STRING, IDM_FILE_EXPORT_MD, Localization::GetString(StringId::Menu_File_Export_MD));
    AppendMenu(hFileMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hFileMenu, MF_STRING, IDM_FILE_EXIT, Localization::GetString(StringId::Menu_File_Exit));
    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hFileMenu, Localization::GetString(StringId::Menu_File));

    // Edit Menu
    HMENU hEditMenu = CreatePopupMenu();
    AppendMenu(hEditMenu, MF_STRING, IDM_EDIT_UNDO, Localization::GetString(StringId::Menu_Edit_Undo));
    AppendMenu(hEditMenu, MF_STRING, IDM_EDIT_REDO, Localization::GetString(StringId::Menu_Edit_Redo));
    AppendMenu(hEditMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hEditMenu, MF_STRING, IDM_EDIT_CUT, Localization::GetString(StringId::Menu_Edit_Cut));
    AppendMenu(hEditMenu, MF_STRING, IDM_EDIT_COPY, Localization::GetString(StringId::Menu_Edit_Copy));
    AppendMenu(hEditMenu, MF_STRING, IDM_EDIT_PASTE, Localization::GetString(StringId::Menu_Edit_Paste));
    AppendMenu(hEditMenu, MF_STRING, IDM_EDIT_DELETE, Localization::GetString(StringId::Menu_Edit_Delete));
    AppendMenu(hEditMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hEditMenu, MF_STRING, IDM_EDIT_INSERT_ROW, Localization::GetString(StringId::Menu_Edit_InsertRow));
    AppendMenu(hEditMenu, MF_STRING, IDM_EDIT_INSERT_COL, Localization::GetString(StringId::Menu_Edit_InsertCol));
    AppendMenu(hEditMenu, MF_STRING, IDM_EDIT_CLEAR, Localization::GetString(StringId::Menu_Edit_Clear));
    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hEditMenu, Localization::GetString(StringId::Menu_Edit));

    // Search Menu
    HMENU hSearchMenu = CreatePopupMenu();
    AppendMenu(hSearchMenu, MF_STRING, ID_SEARCH_CLOSE, Localization::GetString(StringId::Menu_Search_Find));
    AppendMenu(hSearchMenu, MF_STRING, ID_SEARCH_NEXT, Localization::GetString(StringId::Menu_Search_FindNext));
    AppendMenu(hSearchMenu, MF_STRING, ID_SEARCH_PREV, Localization::GetString(StringId::Menu_Search_FindPrev));
    AppendMenu(hSearchMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hSearchMenu, MF_STRING, ID_SEARCH_REPLACE, Localization::GetString(StringId::Menu_Search_Replace));
    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hSearchMenu, Localization::GetString(StringId::Menu_Search));

    // Config Menu
    HMENU hConfigMenu = CreatePopupMenu();
    
    // Delimiter Submenu
    HMENU hDelimMenu = CreatePopupMenu();
    AppendMenu(hDelimMenu, MF_STRING, IDM_CONFIG_DELIM_COMMA, Localization::GetString(StringId::Menu_Config_Delim_Comma));
    AppendMenu(hDelimMenu, MF_STRING, IDM_CONFIG_DELIM_TAB, Localization::GetString(StringId::Menu_Config_Delim_Tab));
    AppendMenu(hDelimMenu, MF_STRING, IDM_CONFIG_DELIM_SEMI, Localization::GetString(StringId::Menu_Config_Delim_Semi));
    AppendMenu(hDelimMenu, MF_STRING, IDM_CONFIG_DELIM_PIPE, Localization::GetString(StringId::Menu_Config_Delim_Pipe));
    AppendMenu(hConfigMenu, MF_POPUP, (UINT_PTR)hDelimMenu, Localization::GetString(StringId::Menu_Config_Delimiter));

    // Encoding Submenu
    HMENU hEncMenu = CreatePopupMenu();
    AppendMenu(hEncMenu, MF_STRING, IDM_CONFIG_ENC_UTF8, Localization::GetString(StringId::Menu_Config_Enc_UTF8));
    AppendMenu(hEncMenu, MF_STRING, IDM_CONFIG_ENC_ANSI, Localization::GetString(StringId::Menu_Config_Enc_ANSI));
    AppendMenu(hConfigMenu, MF_POPUP, (UINT_PTR)hEncMenu, Localization::GetString(StringId::Menu_Config_Encoding));

    // Newline Submenu
    HMENU hNlMenu = CreatePopupMenu();
    AppendMenu(hNlMenu, MF_STRING, IDM_CONFIG_NL_CRLF, Localization::GetString(StringId::Menu_Config_NL_CRLF));
    AppendMenu(hNlMenu, MF_STRING, IDM_CONFIG_NL_LF, Localization::GetString(StringId::Menu_Config_NL_LF));
    AppendMenu(hNlMenu, MF_STRING, IDM_CONFIG_NL_LF, Localization::GetString(StringId::Menu_Config_NL_LF));
    AppendMenu(hConfigMenu, MF_POPUP, (UINT_PTR)hNlMenu, Localization::GetString(StringId::Menu_Config_Newline));
    
    // Font
    AppendMenu(hConfigMenu, MF_STRING, IDM_CONFIG_FONT, Localization::GetString(StringId::Menu_Config_Font));
    AppendMenu(hConfigMenu, MF_STRING | (m_maxCellLines > 1 ? MF_CHECKED : 0), IDM_CONFIG_FOLDING, Localization::GetString(StringId::Menu_Config_Folding));

    // Language Submenu
    HMENU hLangMenu = CreatePopupMenu();
    AppendMenu(hLangMenu, MF_STRING, IDM_CONFIG_LANG_ENGLISH, Localization::GetString(StringId::Menu_Config_Lang_English));
    AppendMenu(hLangMenu, MF_STRING, IDM_CONFIG_LANG_JAPANESE, Localization::GetString(StringId::Menu_Config_Lang_Japanese));
    AppendMenu(hConfigMenu, MF_POPUP, (UINT_PTR)hLangMenu, Localization::GetString(StringId::Menu_Config_Language));

    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hConfigMenu, Localization::GetString(StringId::Menu_Config));

    // Help Menu
    HMENU hHelpMenu = CreatePopupMenu();
    AppendMenu(hHelpMenu, MF_STRING, IDM_HELP_View, Localization::GetString(StringId::Menu_Help_View));
    AppendMenu(hHelpMenu, MF_STRING, IDM_HELP_ABOUT, Localization::GetString(StringId::Menu_Help_About));
    AppendMenu(hMenu, MF_POPUP, (UINT_PTR)hHelpMenu, Localization::GetString(StringId::Menu_Help));

    SetMenu(hwnd, hMenu);
    DrawMenuBar(hwnd); // Force redraw
}

void MainWindow::OnCommand(HWND hwnd, int id, HWND hwndCtl, UINT codeNotify)
{
    switch (id) {
    case ID_FORMULA_EDIT:
        if (codeNotify == EN_CHANGE) {
            if (m_isUpdatingFormula) break;
            
            // Sync Edit -> Cell
            int len = GetWindowTextLength(m_hFormulaEdit);
            std::wstring text; text.resize(len+1);
            GetWindowText(m_hFormulaEdit, &text[0], len+1);
            text.resize(len);
            
            auto selections = m_state.GetSelections();
            if (!selections.empty()) {
                size_t r = selections[0].start.row;
                size_t c = selections[0].start.col;
                m_document.UpdateCell(r, c, text);
                InvalidateRect(m_hwnd, NULL, FALSE);
            }
        }
        break;
        
    case IDM_FILE_OPEN:
        OnFileOpen();
        break;
    case IDM_FILE_IMPORT:
        OnFileImport();
        break;
    case IDM_FILE_EXPORT_HTML:
        OnFileExport(CsvDocument::ExportFormat::HTML);
        break;
    case IDM_FILE_EXPORT_MD:
        OnFileExport(CsvDocument::ExportFormat::Markdown);
        break;
    case IDM_FILE_EXIT:
        OnExit();
        break;
    case IDM_EDIT_COPY:
        OnEditCopy();
        break;
    case IDM_EDIT_PASTE:
        OnEditPaste();
        break;
    case IDM_EDIT_CUT:
        OnEditCut();
        break;
    case IDM_EDIT_INSERT_ROW:
        OnEditInsertRow();
        break;
    case IDM_EDIT_INSERT_COL:
        OnEditInsertColumn();
        break;
        
    case IDM_EDIT_INSERT_ROW_UP:
        {
            WaitCursor wait;
            size_t row = 0;
            const auto& selections = m_state.GetSelections();
            if (!selections.empty()) row = selections[0].start.row;
            m_document.InsertRow(row, {});
            InvalidateRect(m_hwnd, NULL, FALSE);
            UpdateScrollBars();
        }
        break;
    case IDM_EDIT_INSERT_ROW_DOWN:
        {
            WaitCursor wait;
            size_t row = 0;
            const auto& selections = m_state.GetSelections();
            if (!selections.empty()) row = selections[0].start.row + 1;
            else row = m_document.GetRowCount();
            m_document.InsertRow(row, {});
            InvalidateRect(m_hwnd, NULL, FALSE);
            UpdateScrollBars();
        }
        break;
    case IDM_EDIT_INSERT_COL_LEFT:
        {
            WaitCursor wait;
            size_t col = 0;
            const auto& selections = m_state.GetSelections();
            if (!selections.empty()) col = selections[0].start.col;
            m_document.InsertColumn(col, L"");
            InvalidateRect(m_hwnd, NULL, FALSE);
            UpdateScrollBars();
        }
        break;
    case IDM_EDIT_INSERT_COL_RIGHT:
        {
            WaitCursor wait;
            size_t col = 0;
            const auto& selections = m_state.GetSelections();
            if (!selections.empty()) col = selections[0].start.col + 1;
            else col = 1; // Default? Or 0?
             // Actually if empty selection, cannot determine right of what.
             // Default to appending? Or just 0?
             // If empty, start.col is 0 usually?
             // Let's assume 0 if empty.
            if (selections.empty()) col = 0; 
            
            // Wait, "Right" of selection means index + 1.
            m_document.InsertColumn(col, L"");
            InvalidateRect(m_hwnd, NULL, FALSE);
            UpdateScrollBars();
        }
        break;
        
    case IDM_EDIT_UNDO:
        if (m_document.CanUndo()) {
            WaitCursor wait;
            m_document.Undo();
            UpdateScrollBars();
            InvalidateRect(m_hwnd, NULL, FALSE);
        }
        break;
    case IDM_EDIT_REDO:
        if (m_document.CanRedo()) {
            WaitCursor wait;
            m_document.Redo();
            UpdateScrollBars();
            InvalidateRect(m_hwnd, NULL, FALSE);
        }
        break;
        
    case IDM_CONFIG_FONT:
        OnConfigFont();
        break;
        
    case IDM_CONFIG_FOLDING:
        m_maxCellLines = (m_maxCellLines > 1) ? 1 : 5; // Toggle between 1 implementation and 5 (wrapped)
        InitializeMenu(m_hwnd); // Update checkmark
        InvalidateRect(m_hwnd, NULL, FALSE);
        break;

    case ID_SEARCH_CLOSE:
        m_showSearch = false;
        m_showReplace = false;
        UpdateSearchBar();
        break;
    case ID_SEARCH_NEXT:
        OnSearchNext();
        break;
    case ID_SEARCH_PREV:
        OnSearchPrev();
        break;

    case ID_SEARCH_REPLACE:
        OnSearchReplace();
        break;
    case ID_REPLACE_BTN:
        OnReplace();
        break;
    case ID_REPLACE_ALL_BTN:
        OnReplaceAll();
        break;
    case IDM_CONFIG_DELIM_COMMA: m_document.SetDelimiter(L','); m_document.RebuildRowIndex(); InvalidateRect(m_hwnd, NULL, FALSE); break;
    case IDM_CONFIG_DELIM_TAB:   m_document.SetDelimiter(L'\t'); m_document.RebuildRowIndex(); InvalidateRect(m_hwnd, NULL, FALSE); break;
    case IDM_CONFIG_DELIM_SEMI:  m_document.SetDelimiter(L';'); m_document.RebuildRowIndex(); InvalidateRect(m_hwnd, NULL, FALSE); break;
    case IDM_CONFIG_DELIM_PIPE:  m_document.SetDelimiter(L'|'); m_document.RebuildRowIndex(); InvalidateRect(m_hwnd, NULL, FALSE); break;

    case IDM_CONFIG_ENC_UTF8:    m_document.SetEncoding(FileEncoding::UTF8); break;
    
    case IDM_CONFIG_NL_CRLF:     /* m_document.SetNewline(L"\r\n"); */ break;
    case IDM_CONFIG_NL_LF:       /* m_document.SetNewline(L"\n"); */ break;

    case IDM_CONFIG_LANG_ENGLISH:
        Localization::SetLanguage(Language::English);
        InitializeMenu(m_hwnd);
        break;
    case IDM_CONFIG_LANG_JAPANESE:
        Localization::SetLanguage(Language::Japanese);
        InitializeMenu(m_hwnd);
        break;

    case IDM_EDIT_DELETE:
        OnEditDelete();
        break;
    case IDM_EDIT_CLEAR:
        OnEditClear();
        break;
    case IDM_FILE_SAVEAS:
        OnFileSaveAs();
        break;
    case IDM_HELP_ABOUT:
        MessageBox(hwnd, Localization::GetString(StringId::Msg_AppTitle), Localization::GetString(StringId::Msg_About), MB_OK);
        break;
    // Add other handlers...
    }
}

void MainWindow::OnFileOpen()
{
    OPENFILENAME ofn;
    TCHAR szFile[260] = {0};

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = m_hwnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    
    // Construct filter: "CSV Files\0*.csv\0TSV Files\0*.tsv\0All Files\0*.*\0"
    static std::vector<TCHAR> filter; 
    filter.clear();
    auto append = [&](const TCHAR* s) { while(*s) filter.push_back(*s++); filter.push_back('\0'); };
    append(Localization::GetString(StringId::File_Filter_CSV)); append(_T("*.csv"));
    append(Localization::GetString(StringId::File_Filter_TSV)); append(_T("*.tsv"));
    append(Localization::GetString(StringId::File_Filter_All)); append(_T("*.*"));
    filter.push_back('\0');

    ofn.lpstrFilter = filter.data();
    ofn.nFilterIndex = 1;
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.lpstrInitialDir = NULL;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

    if (GetOpenFileName(&ofn) == TRUE) {
        WaitCursor wait;
        if (OpenFile(szFile)) {
            // Update window title or status
            // SetWindowText(m_hwnd, szFile);
        } else {
            MessageBox(m_hwnd, Localization::GetString(StringId::Msg_OpenFailed), Localization::GetString(StringId::Msg_Error), MB_ICONERROR);
        }
    }
}

void MainWindow::OnFileSave()
{
    // If we have a current file path, save to it.
    // If not (new file), call SaveAs.
    // Currently OpenFile sets m_document internal path? No, CsvDocument doesn't store path publicly?
    // MainWindow should track current file path.
    // For now, I'll rely on a member variable I need to add to MainWindow: m_currentFilePath.
    if (!m_currentFilePath.empty()) {
        WaitCursor wait;
        if (m_document.Save(m_currentFilePath)) {
            MessageBox(m_hwnd, Localization::GetString(StringId::Msg_SaveSuccess), Localization::GetString(StringId::Msg_Info), MB_OK);
        } else {
            MessageBox(m_hwnd, Localization::GetString(StringId::Msg_SaveFailed), Localization::GetString(StringId::Msg_Error), MB_ICONERROR);
        }
    } else {
        OnFileSaveAs();
    }
}

void MainWindow::OnFileSaveAs()
{
    OPENFILENAME ofn;
    TCHAR szFile[260] = {0};

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = m_hwnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    
    // Construct filter: "CSV Files\0*.csv\0TSV Files\0*.tsv\0All Files\0*.*\0"
    static std::vector<TCHAR> filter; 
    filter.clear();
    auto append = [&](const TCHAR* s) { while(*s) filter.push_back(*s++); filter.push_back('\0'); };
    append(Localization::GetString(StringId::File_Filter_CSV)); append(_T("*.csv"));
    append(Localization::GetString(StringId::File_Filter_TSV)); append(_T("*.tsv"));
    append(Localization::GetString(StringId::File_Filter_All)); append(_T("*.*"));
    filter.push_back('\0');

    ofn.lpstrFilter = filter.data();
    ofn.nFilterIndex = 1;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;

    if (GetSaveFileName(&ofn) == TRUE) {
        WaitCursor wait;
        if (m_document.Save(szFile)) {
            m_currentFilePath = szFile;
            // Update title
            SetWindowText(m_hwnd, szFile);
        } else {
            MessageBox(m_hwnd, Localization::GetString(StringId::Msg_SaveFailed), Localization::GetString(StringId::Msg_Error), MB_ICONERROR);
        }
    }
}

void MainWindow::OnFileExport(CsvDocument::ExportFormat format)
{
    OPENFILENAME ofn;
    TCHAR szFile[260] = {0};

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = m_hwnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    
    if (format == CsvDocument::ExportFormat::HTML) {
        ofn.lpstrFilter = _T("HTML Files\0*.html;*.htm\0All Files\0*.*\0");
        ofn.lpstrDefExt = _T("html");
    } else {
        ofn.lpstrFilter = _T("Markdown Files\0*.md;*.markdown\0All Files\0*.*\0");
        ofn.lpstrDefExt = _T("md");
    }

    ofn.nFilterIndex = 1;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;

    if (GetSaveFileName(&ofn) == TRUE) {
        // TCHAR to wstring
        std::wstring path = szFile;
        // Check if we need to call Localization for success/fail messages?
        // Using existing SaveSuccess/Fail msg is probably fine or generic.
        WaitCursor wait;
        if (m_document.Export(path, format)) {
             MessageBox(m_hwnd, Localization::GetString(StringId::Msg_SaveSuccess), Localization::GetString(StringId::Msg_Info), MB_OK);
        } else {
             MessageBox(m_hwnd, Localization::GetString(StringId::Msg_SaveFailed), Localization::GetString(StringId::Msg_Error), MB_ICONERROR);
        }
    }
}

void MainWindow::OnFileImport()
{
    OPENFILENAME ofn;
    TCHAR szFile[260] = {0};

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = m_hwnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    
    // Construct filter: "CSV Files\0*.csv\0TSV Files\0*.tsv\0All Files\0*.*\0"
    static std::vector<TCHAR> filter; 
    filter.clear();
    auto append = [&](const TCHAR* s) { while(*s) filter.push_back(*s++); filter.push_back('\0'); };
    append(Localization::GetString(StringId::File_Filter_CSV)); append(_T("*.csv"));
    append(Localization::GetString(StringId::File_Filter_TSV)); append(_T("*.tsv"));
    append(Localization::GetString(StringId::File_Filter_All)); append(_T("*.*"));
    filter.push_back('\0');

    ofn.lpstrFilter = filter.data();
    ofn.nFilterIndex = 1;
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.lpstrInitialDir = NULL;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

    if (GetOpenFileName(&ofn) == TRUE) {
        WaitCursor wait;
        if (m_document.Import(szFile)) {
             // Invalidate entire window to refresh grid
             InvalidateRect(m_hwnd, NULL, TRUE);
             UpdateScrollBars(); 
             MessageBox(m_hwnd, Localization::GetString(StringId::Msg_Info), Localization::GetString(StringId::Msg_Info), MB_OK); // "Info" / "Imported" (reuse info string for now)
        } else {
             MessageBox(m_hwnd, Localization::GetString(StringId::Msg_OpenFailed), Localization::GetString(StringId::Msg_Error), MB_ICONERROR);
        }
    }
}

void MainWindow::OnEditDelete()
{
    // Delete selected rows or columns
    
    const auto& selections = m_state.GetSelections();
    if (selections.empty()) return;

    // Handle multiple selections?
    // For now, take first selection range
    const auto& range = selections[0];
    
    WaitCursor wait;

    if (range.mode == SelectionMode::Row) {
        // Delete rows from range.start.row to range.end.row
        // Delete in reverse order to preserve indices!
        size_t start = (std::min)(range.start.row, range.end.row);
        size_t end = (std::max)(range.start.row, range.end.row);
        
        for (size_t i = end; ; --i) {
            m_document.DeleteRow(i);
            if (i == start) break;
        }
        // m_state.ClearSelection(); // Keep selection? Or clear because rows are gone.
        // Better clear to avoid invalid indices.
        m_state.ClearSelection();
        InvalidateRect(m_hwnd, NULL, FALSE);
    } else if (range.mode == SelectionMode::Column) {
        size_t start = (std::min)(range.start.col, range.end.col);
        size_t end = (std::max)(range.start.col, range.end.col);

        for (size_t i = end; ; --i) {
            m_document.DeleteColumn(i);
            if (i == start) break;
        }
        m_state.ClearSelection();
        InvalidateRect(m_hwnd, NULL, FALSE);
    }
}


void MainWindow::OnEditInsertColumn()
{
    // Insert before current selection
    const auto& selections = m_state.GetSelections();
    size_t col = 0;
    if (!selections.empty()) {
        col = selections[0].start.col;
    }
    
    m_document.InsertColumn(col, L"");
    
    InvalidateRect(m_hwnd, NULL, FALSE);
    UpdateScrollBars();
}

void MainWindow::OnEditClear()
{
    // Clear content of selected cells
    const auto& selections = m_state.GetSelections();
    if (selections.empty()) return;
    
    const auto& range = selections[0];
    
    size_t startRow = (std::min)(range.start.row, range.end.row);
    size_t endRow = (std::max)(range.start.row, range.end.row);
    size_t startCol = (std::min)(range.start.col, range.end.col);
    size_t endCol = (std::max)(range.start.col, range.end.col);
    
    for (size_t r = startRow; r <= endRow; ++r) {
        for (size_t c = startCol; c <= endCol; ++c) {
            m_document.UpdateCell(r, c, L"");
        }
    }
    
    UpdateFormulaBar();
    InvalidateRect(m_hwnd, NULL, FALSE);
}

void MainWindow::OnEditCopy()
{
    const auto& selections = m_state.GetSelections();
    if (selections.empty()) return;
    const auto& range = selections[0]; // Multi-selection copy is complex, use primary
    
    // Normalize range
    size_t startRow = (std::min)(range.start.row, range.end.row);
    size_t endRow = (std::max)(range.start.row, range.end.row);
    size_t startCol = (std::min)(range.start.col, range.end.col);
    size_t endCol = (std::max)(range.start.col, range.end.col);

    // If Row selection, handle max col
    if (range.mode == SelectionMode::Row) {
         endCol = 100; // HACK: Should ideally find max cols in rows
    }
    // If Column selection, handle max row
    else if (range.mode == SelectionMode::Column) {
         endRow = m_document.GetRowCount() > 0 ? m_document.GetRowCount() - 1 : 0;
    }
    // If All selection
    else if (range.mode == SelectionMode::All) {
         endRow = m_document.GetRowCount() > 0 ? m_document.GetRowCount() - 1 : 0;
         endCol = 100; // HACK: Max cols
    }

    std::wstring text = m_document.GetRangeAsText(startRow, startCol, endRow, endCol);

    if (OpenClipboard(m_hwnd)) {
        EmptyClipboard();
        HGLOBAL hGlob = GlobalAlloc(GMEM_MOVEABLE, (text.length() + 1) * sizeof(wchar_t));
        if (hGlob) {
            memcpy(GlobalLock(hGlob), text.c_str(), (text.length() + 1) * sizeof(wchar_t));
            GlobalUnlock(hGlob);
            SetClipboardData(CF_UNICODETEXT, hGlob);
        }
        CloseClipboard();
    }
}

void MainWindow::OnEditPaste()
{
    // Basic Paste: Insert text at start of selection
    // Parsing the pasted CSV/TSV data and inserting into table.
    CsvDocument pastedDoc;
    
    if (!OpenClipboard(m_hwnd)) return;
    
    // ScopedWaitCursor for potential long parse
    WaitCursor wait;
    
    if (IsClipboardFormatAvailable(CF_UNICODETEXT)) {
        HGLOBAL hGlob = GetClipboardData(CF_UNICODETEXT);
        if (hGlob) {
            wchar_t* pText = (wchar_t*)GlobalLock(hGlob);
            if (pText) {
                std::wstring pasteData = pText;
                GlobalUnlock(hGlob); // Unlock early as we copied data
                
                const auto& selections = m_state.GetSelections();
                if (!selections.empty()) {
                    size_t r = selections[0].start.row;
                    size_t c = selections[0].start.col;
                    m_document.PasteCells(r, c, pasteData);
                    InvalidateRect(m_hwnd, NULL, FALSE);
                }
            }
        }
    }
    CloseClipboard();
}

void MainWindow::OnEditCut()
{
    OnEditCopy();
    OnEditClear(); // Or Delete? Excel clears content on cut usually, moving it. Text editors delete.
    // Excel style: "Cut" mode, then move. 
    // Text Editor style: Copy + Delete.
    // Let's go with Clear for cells, Delete for Row mode.
    const auto& selections = m_state.GetSelections();
    if (!selections.empty() && selections[0].mode == SelectionMode::Row) {
         OnEditDelete();
    } else {
         OnEditClear();
    }
}

void MainWindow::OnEditInsertRow()
{
    const auto& selections = m_state.GetSelections();
    size_t insertRow = 0;
    
    if (!selections.empty()) {
        insertRow = selections[0].start.row;
    } else {
        insertRow = m_document.GetRowCount();
    }
    
    // Insert empty row
    std::vector<std::wstring> emptyValues; // No columns? Or default columns?
    // CsvDocument InsertRow handles empty vector -> just delimiter+newline?
    // Let's insert "" to ensure at least one cell if needed or just newline.
    // Actually InsertRow logic iterates values. If empty, it adds just \n.
    m_document.InsertRow(insertRow, emptyValues);
    
    InvalidateRect(m_hwnd, NULL, FALSE);
    UpdateScrollBars();
}

void MainWindow::OnExit()
{
    DestroyWindow(m_hwnd);
}

void MainWindow::UpdateScrollBars()
{
    if (!m_hwnd) return;

    RECT rc;
    GetClientRect(m_hwnd, &rc);
    int clientHeight = rc.bottom - rc.top;
    int clientWidth = rc.right - rc.left;

    // Vertical Scrollbar
    // Range: 0 to RowCount
    // Page: VisibleRows
    
    size_t rowCount = m_document.GetRowCount();
    int visibleRows = (int)((clientHeight - m_headerHeight) / m_rowHeight);
    if (visibleRows < 0) visibleRows = 0;

    SCROLLINFO si = {0};
    si.cbSize = sizeof(SCROLLINFO);
    si.fMask = SIF_RANGE | SIF_PAGE | SIF_POS;
    si.nMin = 0;
    si.nMax = (int)rowCount + visibleRows - 1; // +visibleRows to allow scrolling to bottom
    // Actually standard is nMax. 
    // If nMax = 100, nPage = 10. Max Pos = 91. 
    // We want to be able to scroll until the last row is visible. 
    
    si.nMax = (int)rowCount - 1; 
    // If we want "past end" scrolling, increase max.
    // Let's stick to standard map: Pos 0 = Row 0. Pos Max = Last Row at top? 
    // Usually we want Last Row at bottom. => Max Pos = RowCount - VisibleRows.
    // Set nMax = RowCount - 1. nPage = VisibleRows.
    // The scrollbar logic handles nMax - nPage + 1 automatically.
    
    si.nMax = (int)rowCount > 0 ? (int)rowCount - 1 : 0;
    si.nPage = visibleRows;
    si.nPos = (int)m_state.GetScrollRow();
    
    SetScrollInfo(m_hwnd, SB_VERT, &si, TRUE);

    // Horizontal Scrollbar
    // Range: Pixels
    float totalWidth = GetColumnX(m_document.GetMaxColumnCount());
    // Ensure at least some width
    if (totalWidth < clientWidth) totalWidth = (float)clientWidth;

    si.nMax = (int)totalWidth;
    si.nPage = clientWidth;
    si.nPos = (int)m_state.GetScrollX();
    
    SetScrollInfo(m_hwnd, SB_HORZ, &si, TRUE);
    SetScrollInfo(m_hwnd, SB_HORZ, &si, TRUE);
}

void MainWindow::OnVScroll(WPARAM wParam)
{
    SCROLLINFO si = {0};
    si.cbSize = sizeof(SCROLLINFO);
    si.fMask = SIF_ALL;
    GetScrollInfo(m_hwnd, SB_VERT, &si);

    int newPos = si.nPos;
    int action = LOWORD(wParam);

    switch (action) {
    case SB_TOP: newPos = si.nMin; break;
    case SB_BOTTOM: newPos = si.nMax; break;
    case SB_LINEUP: newPos -= 1; break;
    case SB_LINEDOWN: newPos += 1; break;
    case SB_PAGEUP: newPos -= si.nPage; break;
    case SB_PAGEDOWN: newPos += si.nPage; break;
    case SB_THUMBTRACK: newPos = si.nTrackPos; break;
    }

    if (newPos < si.nMin) newPos = si.nMin;
    if (newPos > (int)(si.nMax - si.nPage + 1)) newPos = (si.nMax - si.nPage + 1);
    // Ensure we don't go negative if page > max
    if (newPos < 0) newPos = 0;

    m_state.SetScrollRow((size_t)newPos);
    UpdateScrollBars();
    InvalidateRect(m_hwnd, NULL, FALSE);
}

void MainWindow::OnHScroll(WPARAM wParam)
{
    SCROLLINFO si = {0};
    si.cbSize = sizeof(SCROLLINFO);
    si.fMask = SIF_ALL;
    GetScrollInfo(m_hwnd, SB_HORZ, &si);

    int newPos = si.nPos;
    int action = LOWORD(wParam);

    switch (action) {
    case SB_LINELEFT: newPos -= 10; break;
    case SB_LINERIGHT: newPos += 10; break;
    case SB_PAGELEFT: newPos -= si.nPage; break;
    case SB_PAGERIGHT: newPos += si.nPage; break;
    case SB_THUMBTRACK: newPos = si.nTrackPos; break;
    }
    
    if (newPos < si.nMin) newPos = si.nMin;
    if (newPos > (int)(si.nMax - si.nPage + 1)) newPos = (si.nMax - si.nPage + 1);
    if (newPos < 0) newPos = 0;

    m_state.SetScrollX((float)newPos);
    UpdateScrollBars();
    InvalidateRect(m_hwnd, NULL, FALSE);
}

void MainWindow::OnMouseWheel(short delta)
{
    // Delta / 120 = clicks. Each click = 3 lines usually.
    int lines = - (delta / WHEEL_DELTA) * 3;
    
    int currentRow = (int)m_state.GetScrollRow();
    int newRow = currentRow + lines;
    if (newRow < 0) newRow = 0;
    
    // Check max
    size_t rowCount = m_document.GetRowCount();
    if (newRow >= (int)rowCount) newRow = (int)rowCount - 1; // Basic clamp
    
    m_state.SetScrollRow((size_t)newRow);
    UpdateScrollBars();
    InvalidateRect(m_hwnd, NULL, FALSE);
}

bool MainWindow::HitTest(int x, int y, size_t& outRow, size_t& outCol, bool& outIsRowHeader, bool& outIsColHeader)
{
    float fx = (float)x;
    float fy = (float)y;
    
    // Formula Bar Offset
    fy -= m_formulaBarHeight;
    if (fy < 0) return false; // Clicked in formula bar or above

    outIsRowHeader = false;
    outIsColHeader = false;

    // Check Headers
    if (fx < m_headerWidth && fy < m_headerHeight) {
        // Top-left corner (Select all)
        outIsRowHeader = true;
        outIsColHeader = true;
        return true; 
    }

    if (fx < m_headerWidth) {
        outIsRowHeader = true;
        // Calculate row
        float relativeY = fy - m_headerHeight;
        if (relativeY < 0) return false;
        
        outRow = m_state.GetScrollRow() + (size_t)(relativeY / m_rowHeight);
        return true;
    }

    if (fy < m_headerHeight) {
        outIsColHeader = true;
        // Calculate col
        float relativeX = fx - m_headerWidth + m_state.GetScrollX(); 
        
        // Find column index based on cumulative width

        outCol = GetColumnAtX(relativeX);
        return true;
    }

    // Grid Area
    float relativeY = fy - m_headerHeight;
    float relativeX = fx - m_headerWidth + m_state.GetScrollX(); 
    
    outRow = m_state.GetScrollRow() + (size_t)(relativeY / m_rowHeight);
    outCol = GetColumnAtX(relativeX);
    
    return true;
}

void MainWindow::OnLButtonDown(int x, int y, UINT keyFlags)
{
    SetCapture(m_hwnd);
    
    // Check Resize
    float fy = (float)y - m_formulaBarHeight;
    bool inHeader = (fy >= 0 && fy < m_headerHeight);
    if (inHeader && x > m_headerWidth) {
        float gridX = (float)x - m_headerWidth;
        float scrollX = m_state.GetScrollX();
        float contentX = gridX + scrollX;
        size_t col = GetColumnAtX(contentX);
        float colRight = GetColumnX(col) + GetColumnWidth(col);
        
        if (abs(contentX - colRight) < 8.0f) {
            m_isResizingCol = true;
            m_resizingColIndex = col;
            m_resizeStartX = (float)x;
            m_resizeStartWidth = GetColumnWidth(col);
            return; // Consume event
        }
    }

    size_t row, col;
    bool isRowHeader, isColHeader;
    
    if (HitTest(x, y, row, col, isRowHeader, isColHeader)) {
        bool ctrl = (keyFlags & MK_CONTROL); // Logical AND for check usually or GetKeyState

        if (isRowHeader && isColHeader) {
             m_state.SelectAll();
        } else if (isRowHeader) {
            m_state.SelectRow(row, ctrl);
        } else if (isColHeader) {
            m_state.SelectColumn(col, ctrl);
        } else {
            m_state.SelectCell(row, col, ctrl);
        }
        
        UpdateFormulaBar();
        InvalidateRect(m_hwnd, NULL, FALSE);
    }
}

void MainWindow::OnLButtonUp(int x, int y, UINT keyFlags)
{
    if (m_isResizingCol) {
        m_isResizingCol = false;
    }
    ReleaseCapture();
}

void MainWindow::OnMouseMove(int x, int y, UINT keyFlags)
{
    // Adjust y for formula bar
    // Standard Grid Area check for Resize Cursor
    // Only allow resize in header area? Or entire column line? Usually Header + Grid.
    // Let's stick to Header area for triggering, or Header area for visual feedback? 
    // Excel allows resizing by dragging separator in Header. Dragging in grid usually selects.
    // User said "dragging with holding lines dividing cells".
    // "holding lines" -> grabbing lines?
    // Usually standard UI allows resize from Header. 
    // If I allow resize from Grid lines, it conflicts with Selection.
    // I will enable it in Header area (like Excel).
    
    // Check if in Header vertical range
    float fy = (float)y - m_formulaBarHeight;
    bool inHeader = (fy >= 0 && fy < m_headerHeight);
    
    if (m_isResizingCol) {
        float delta = (float)x - m_resizeStartX;
        float newW = m_resizeStartWidth + delta;
        if (newW < 20.0f) newW = 20.0f;
        
        m_colWidths[m_resizingColIndex] = newW;
        InvalidateRect(m_hwnd, NULL, FALSE);
        UpdateScrollBars(); // Width changed
        return;
    }

    // Check for resize cursor
    if (inHeader && x > m_headerWidth) {
        float gridX = (float)x - m_headerWidth;
        float scrollX = m_state.GetScrollX();
        float contentX = gridX + scrollX;
        
        size_t col = GetColumnAtX(contentX);
        float colX = GetColumnX(col); // start of col
        float colW = GetColumnWidth(col);
        float colRight = colX + colW;
        
        if (abs(contentX - colRight) < 8.0f) { // Tolerance
            SetCursor(m_hCursorWE);
            return; 
        }
    }
    SetCursor(m_hCursorArrow);

    if (keyFlags & MK_LBUTTON) {
        size_t row, col;
        bool isRowHeader, isColHeader;
        
        if (HitTest(x, y, row, col, isRowHeader, isColHeader)) {
            m_state.DragTo(row, col);
            InvalidateRect(m_hwnd, NULL, FALSE);
            UpdateFormulaBar();
        }
    }
}


void MainWindow::OnContextMenu(HWND hwnd, int x, int y)
{
    // If handled by keyboard (x=-1, y=-1), get cursor pos
    if (x == -1 && y == -1) {
        POINT pt;
        GetCursorPos(&pt);
        x = pt.x;
        y = pt.y;
    }

    // Convert to client for HitTest?
    // HitTest expects Client coordinates.
    POINT ptClient = { x, y };
    ScreenToClient(m_hwnd, &ptClient);

    size_t row, col;
    bool isRowHeader, isColHeader;
    
    // Only show if we hit something valid or generally inside the grid/header area
    // HitTest returns true for headers and cells
    if (HitTest(ptClient.x, ptClient.y, row, col, isRowHeader, isColHeader)) {
        
        // If it's a header or cell, we show the menu.
        // We might want to ensure selection is correct? 
        // Standard behavior: Right click selects the item if not already selected.
        // But for multiple selection, right click on *selection* keeps selection. 
        // Right click *outside* selection selects that single item.
        
        bool isSelected = false;
        if (isRowHeader) isSelected = m_state.IsRowSelected(row);
        else if (isColHeader) isSelected = m_state.IsColumnSelected(col);
        else isSelected = m_state.IsSelected(row, col);

        if (!isSelected) {
            // Select it (singlular)
            if (isRowHeader) m_state.SelectRow(row, false);
            else if (isColHeader) m_state.SelectColumn(col, false);
            else m_state.SelectCell(row, col, false);
            
            InvalidateRect(m_hwnd, NULL, FALSE);
        }

        HMENU hMenu = CreatePopupMenu();
        AppendMenu(hMenu, MF_STRING, IDM_EDIT_CUT, Localization::GetString(StringId::Menu_Edit_Cut));
        AppendMenu(hMenu, MF_STRING, IDM_EDIT_COPY, Localization::GetString(StringId::Menu_Edit_Copy));
        AppendMenu(hMenu, MF_STRING, IDM_EDIT_PASTE, Localization::GetString(StringId::Menu_Edit_Paste));
        
        // For Delete, context sensitive string? Or just "Delete"
        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);
        AppendMenu(hMenu, MF_STRING, IDM_EDIT_DELETE, Localization::GetString(StringId::Menu_Edit_Delete));

        AppendMenu(hMenu, MF_SEPARATOR, 0, NULL);

        if (isRowHeader || (!isRowHeader && !isColHeader)) {
            AppendMenu(hMenu, MF_STRING, IDM_EDIT_INSERT_ROW_UP, Localization::GetString(StringId::Menu_Edit_InsertRowUp));
            AppendMenu(hMenu, MF_STRING, IDM_EDIT_INSERT_ROW_DOWN, Localization::GetString(StringId::Menu_Edit_InsertRowDown));
        }

        if (isColHeader || (!isRowHeader && !isColHeader)) {
            AppendMenu(hMenu, MF_STRING, IDM_EDIT_INSERT_COL_LEFT, Localization::GetString(StringId::Menu_Edit_InsertColLeft));
            AppendMenu(hMenu, MF_STRING, IDM_EDIT_INSERT_COL_RIGHT, Localization::GetString(StringId::Menu_Edit_InsertColRight));
        }

        TrackPopupMenu(hMenu, TPM_LEFTALIGN | TPM_RIGHTBUTTON, x, y, 0, m_hwnd, NULL);
        DestroyMenu(hMenu);
    }
}

void MainWindow::OnResize(UINT width, UINT height)
{
    if (m_hPosLabel) {
        MoveWindow(m_hPosLabel, 0, 0, 40, (int)m_formulaBarHeight, TRUE);
    }
    if (m_hFormulaEdit) {
        MoveWindow(m_hFormulaEdit, 40, 0, width - 40, (int)m_formulaBarHeight, TRUE);
    }
    
    if (m_hProgressBar) {
        MoveWindow(m_hProgressBar, 0, height - (int)m_progressBarHeight, width, (int)m_progressBarHeight, TRUE);
    }

    if (m_dxResources.GetRenderTarget()) {
        // Render target is entire client area.
        m_dxResources.GetRenderTarget()->Resize(D2D1::SizeU(width, height));
    }
    UpdateScrollBars();
}

void MainWindow::OnConfigFont()
{
    CHOOSEFONT cf = {0};
    LOGFONT lf = {0};
    cf.lStructSize = sizeof(cf);
    cf.hwndOwner = m_hwnd;
    cf.lpLogFont = &lf;
    cf.Flags = CF_SCREENFONTS | CF_EFFECTS | CF_INITTOLOGFONTSTRUCT;
    
    // Set current? Not tracking current yet.
    
    if (ChooseFont(&cf)) {
        // Point size is in 1/10 pts
        float fontSize = (float)cf.iPointSize / 10.0f; 
        
        // Update Resources
        m_dxResources.UpdateTextFormat(lf.lfFaceName, fontSize);
        
        // Recalc row height based on font?
        // Simple heuristic: font size * 1.5
        // Or measure text.
        // For now, let's just create format. 
        // Row height needs update maybe?
        m_rowHeight = fontSize * 1.8f; 
        if (m_rowHeight < 20.0f) m_rowHeight = 20.0f;
        
        UpdateScrollBars();
        InvalidateRect(m_hwnd, NULL, FALSE);
    }
}

// Subclass procedure for the Edit control to handle Enter/Esc
LRESULT CALLBACK MainWindow::EditSubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData)
{
    MainWindow* pThis = (MainWindow*)dwRefData;

    switch (uMsg) {
    case WM_GETDLGCODE:
        // Ensure we receive Enter and Tab
        return DLGC_WANTALLKEYS;
        
    case WM_KEYDOWN:
        if (wParam == VK_RETURN) {
            pThis->HideCellEditor(true); // Save
            return 0;
        }
        else if (wParam == VK_ESCAPE) {
            pThis->HideCellEditor(false); // Cancel
            return 0;
        }
        break;
    
    case WM_KILLFOCUS:
        // If we lose focus, save? Usually yes in Excel.
        // But be careful if lose focus because we are destroying it.
        // Let main window handle it via notification if possible? 
        // Or just trigger save here.
        pThis->HideCellEditor(true);
        break;
        
    case WM_NCDESTROY:
        RemoveWindowSubclass(hwnd, EditSubclassProc, uIdSubclass);
        return DefSubclassProc(hwnd, uMsg, wParam, lParam);
    }
    
    return DefSubclassProc(hwnd, uMsg, wParam, lParam);
}

void MainWindow::ShowCellEditor(size_t row, size_t col)
{
    if (m_isEditing) HideCellEditor(true);

    m_editingRow = row;
    m_editingCol = col;
    m_isEditing = true;

    // Calculate Rect
    // Need to account for scroll
    size_t scrollRow = m_state.GetScrollRow();
    
    if (row < scrollRow) return; // Not visible
    
    float y = m_formulaBarHeight + m_headerHeight + (row - scrollRow) * m_rowHeight;
    float x = m_headerWidth + GetColumnX(col) - m_state.GetScrollX(); 
    float w = GetColumnWidth(col);
    
    RECT rc;
    rc.left = (LONG)x;
    rc.top = (LONG)y;
    rc.right = (LONG)(x + w);
    rc.bottom = (LONG)(y + m_rowHeight);
    
    // Get Cell Value
    std::wstring text;
    auto cells = m_document.GetRowCells(row);
    if (col < cells.size()) text = cells[col];
    
    m_hEdit = CreateWindowEx(
        0, _T("EDIT"), text.c_str(),
        WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | ES_LEFT,
        rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top,
        m_hwnd, NULL, GetModuleHandle(NULL), NULL
    );
    
    if (m_hEdit) {
        // Set Font
        // SendMessage(m_hEdit, WM_SETFONT, (WPARAM)hFont, TRUE); 
        // Need to expose font or create simple one. System font for now.
        
        SetWindowSubclass(m_hEdit, EditSubclassProc, 0, (DWORD_PTR)this);
        SetFocus(m_hEdit);
        SendMessage(m_hEdit, EM_SETSEL, 0, -1); // Select all text
    }
}

void MainWindow::HideCellEditor(bool save)
{
    if (!m_isEditing || !m_hEdit) return;

    if (save) {
        // Get text
        int len = GetWindowTextLength(m_hEdit);
        std::wstring text;
        text.resize(len + 1);
        GetWindowText(m_hEdit, &text[0], len + 1);
        text.resize(len); // Remove null terminator
        
        m_document.UpdateCell(m_editingRow, m_editingCol, text);
    }

    m_isEditing = false;
    DestroyWindow(m_hEdit);
    m_hEdit = NULL;
    
    // Restore focus to main window logic handled by OS usually but good to ensure
    // SetFocus(m_hwnd); // Might trigger recursion if called from KillFocus?
    
    InvalidateRect(m_hwnd, NULL, FALSE);
}

void MainWindow::CreateSearchBar()
{
    if (m_hSearchEdit) return;

    // Create simple controls at top
    int y = (int)m_formulaBarHeight;
    int h = 25;
    
    m_hSearchEdit = CreateWindowEx(0, _T("EDIT"), _T(""), WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL, 
        0, y, 200, h, m_hwnd, (HMENU)ID_SEARCH_EDIT, GetModuleHandle(NULL), NULL);
        
    m_hNextBtn = CreateWindowEx(0, _T("BUTTON"), _T("Next"), WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        205, y, 50, h, m_hwnd, (HMENU)ID_SEARCH_NEXT, GetModuleHandle(NULL), NULL);

    m_hPrevBtn = CreateWindowEx(0, _T("BUTTON"), _T("Prev"), WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        260, y, 50, h, m_hwnd, (HMENU)ID_SEARCH_PREV, GetModuleHandle(NULL), NULL);

    m_hCloseSearchBtn = CreateWindowEx(0, _T("BUTTON"), _T("X"), WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
        315, y, 25, h, m_hwnd, (HMENU)ID_SEARCH_CLOSE, GetModuleHandle(NULL), NULL);

    // Replace Controls (Row 2)
    int y2 = y + h;
    m_hReplaceEdit = CreateWindowEx(0, _T("EDIT"), _T(""), WS_CHILD | WS_BORDER | ES_AUTOHSCROLL, 
        0, y2, 200, h, m_hwnd, (HMENU)ID_REPLACE_EDIT, GetModuleHandle(NULL), NULL);

    m_hReplaceBtn = CreateWindowEx(0, _T("BUTTON"), _T("Replace"), WS_CHILD | BS_PUSHBUTTON,
        205, y2, 60, h, m_hwnd, (HMENU)ID_REPLACE_BTN, GetModuleHandle(NULL), NULL);
        
    m_hReplaceAllBtn = CreateWindowEx(0, _T("BUTTON"), _T("All"), WS_CHILD | BS_PUSHBUTTON,
        270, y2, 40, h, m_hwnd, (HMENU)ID_REPLACE_ALL_BTN, GetModuleHandle(NULL), NULL);
        
    // Subclass Edit for Enter key? (Optional)
}

void MainWindow::UpdateSearchBar()
{
    if (m_showSearch) {
        CreateSearchBar();
        ShowWindow(m_hSearchEdit, SW_SHOW);
        ShowWindow(m_hNextBtn, SW_SHOW);
        ShowWindow(m_hPrevBtn, SW_SHOW);
        ShowWindow(m_hCloseSearchBtn, SW_SHOW);
        
        if (m_showReplace) {
             ShowWindow(m_hReplaceEdit, SW_SHOW);
             ShowWindow(m_hReplaceBtn, SW_SHOW);
             ShowWindow(m_hReplaceAllBtn, SW_SHOW);
        } else {
             ShowWindow(m_hReplaceEdit, SW_HIDE);
             ShowWindow(m_hReplaceBtn, SW_HIDE);
             ShowWindow(m_hReplaceAllBtn, SW_HIDE);
        }

        SetFocus(m_hSearchEdit);
    } else {
        if (m_hSearchEdit) ShowWindow(m_hSearchEdit, SW_HIDE);
        if (m_hNextBtn) ShowWindow(m_hNextBtn, SW_HIDE);
        if (m_hPrevBtn) ShowWindow(m_hPrevBtn, SW_HIDE);
        if (m_hCloseSearchBtn) ShowWindow(m_hCloseSearchBtn, SW_HIDE);
        
        if (m_hReplaceEdit) ShowWindow(m_hReplaceEdit, SW_HIDE);
        if (m_hReplaceBtn) ShowWindow(m_hReplaceBtn, SW_HIDE);
        if (m_hReplaceAllBtn) ShowWindow(m_hReplaceAllBtn, SW_HIDE);
        
        // Restore focus to grid/main
        if (m_hwnd) SetFocus(m_hwnd);
    }
    
    // Adjust layout
    m_headerHeight = 25.0f; // Base
    
    if (m_showSearch) {
        m_headerHeight += 25.0f; // Search row
        if (m_showReplace) m_headerHeight += 25.0f; // Replace row
    }
    
    InvalidateRect(m_hwnd, NULL, TRUE);
}

// Replace Handlers
void MainWindow::OnSearchReplace()
{
    m_showSearch = true;
    m_showReplace = true;
    UpdateSearchBar();
}

void MainWindow::OnReplace()
{
    if (!m_hSearchEdit || !m_hReplaceEdit) return;
    
    int len = GetWindowTextLength(m_hSearchEdit);
    if (len == 0) return;
    std::wstring query; query.resize(len+1);
    GetWindowText(m_hSearchEdit, &query[0], len+1);
    query.resize(len);

    int rLen = GetWindowTextLength(m_hReplaceEdit);
    std::wstring replace; replace.resize(rLen+1);
    GetWindowText(m_hReplaceEdit, &replace[0], rLen+1);
    replace.resize(rLen);

    // Get Selection
    auto selections = m_state.GetSelections();
    size_t r = 0, c = 0;
    if (!selections.empty()) {
        r = selections[0].start.row;
        c = selections[0].start.col;
    }
    
    // Search Options
    CsvDocument::SearchOptions opts;
    opts.forward = true;
    opts.mode = CsvDocument::SearchMode::Contains;
    
    // Try Replace
    bool replaced = false;
    {
        WaitCursor wait;
        replaced = m_document.Replace(query, replace, r, c, opts);
    }
    
    if (replaced) {
        // Successful replace at r, c.
        // Move selection? Or just stay?
        // Usually jump to next.
        // Let's Search Next from current position + 1?
        // Actually Search() handles finding next if called again.
        // But we modified the cell. 
        // Let's just find next.
        
        // Find next *after* this cell?
        // If we replaced "foo" with "bar", we are still at r,c.
        // Search should start from r,c?
        // If "bar" matches "foo", we might loop? "Contains" mode "foo" -> "foofoo" -> loop.
        // Search takes start pos.
        // Let's increment c temporarily to skip? 
        
        // Actually, just calling Search(..., forward=true) will find *next* match if we assume we are at current match.
        // But Replace updates the cell.
        
        // Let's just update view first.
        InvalidateRect(m_hwnd, NULL, FALSE);
        
        // Find Next
        if (m_document.Search(query, r, c, opts)) {
             m_state.SelectCell(r, c, false);
             m_state.SetScrollRow(r > 5 ? r - 5 : 0);
             UpdateScrollBars();
        }
    } else {
        // Maybe try to Find first if not currently on match?
        if (m_document.Search(query, r, c, opts)) {
             m_state.SelectCell(r, c, false);
             m_state.SetScrollRow(r > 5 ? r - 5 : 0);
             UpdateScrollBars();
             // Found it, now user clicks Replace again to replace.
        } else {
             MessageBox(m_hwnd, _T("Not found"), _T("Replace"), MB_OK);
        }
    }
}

void MainWindow::OnReplaceAll()
{
    if (!m_hSearchEdit || !m_hReplaceEdit) return;
    
    int len = GetWindowTextLength(m_hSearchEdit);
    if (len == 0) return;
    std::wstring query; query.resize(len+1);
    GetWindowText(m_hSearchEdit, &query[0], len+1);
    query.resize(len);

    int rLen = GetWindowTextLength(m_hReplaceEdit);
    std::wstring replace; replace.resize(rLen+1);
    GetWindowText(m_hReplaceEdit, &replace[0], rLen+1);
    replace.resize(rLen);

    CsvDocument::SearchOptions opts;
    opts.mode = CsvDocument::SearchMode::Contains;
    
    int count = 0;
    {
        WaitCursor wait;
        count = m_document.ReplaceAll(query, replace, opts);
    }
    
    std::wstring msg = _T("Replaced ") + std::to_wstring(count) + _T(" occurrences.");
    MessageBox(m_hwnd, msg.c_str(), _T("Replace All"), MB_OK);
    
    InvalidateRect(m_hwnd, NULL, FALSE);
}

void MainWindow::OnSearchNext()
{
    if (!m_hSearchEdit) return;
    int len = GetWindowTextLength(m_hSearchEdit);
    if (len == 0) return;
    std::wstring query;
    query.resize(len+1);
    GetWindowText(m_hSearchEdit, &query[0], len+1);
    query.resize(len);
    
    auto selections = m_state.GetSelections();
    size_t r = 0, c = 0;
    if (!selections.empty()) {
        r = selections[0].start.row;
        c = selections[0].start.col;
    }
    
    CsvDocument::SearchOptions opts;
    opts.forward = true;
    opts.matchCase = false; 
    opts.mode = CsvDocument::SearchMode::Contains;

    bool found = false;
    {
         WaitCursor wait;
         found = m_document.Search(query, r, c, opts);
    }

    if (found) {
        m_state.SelectCell(r, c, false);
        // Ensure visible
        m_state.SetScrollRow(r > 5 ? r - 5 : 0); // Center roughly
        UpdateScrollBars();
        InvalidateRect(m_hwnd, NULL, FALSE);
    } else {
        MessageBox(m_hwnd, _T("Not found"), _T("Search"), MB_OK);
    }
}

void MainWindow::OnSearchPrev()
{
    if (!m_hSearchEdit) return;
    int len = GetWindowTextLength(m_hSearchEdit);
    if (len == 0) return;
    std::wstring query;
    query.resize(len+1);
    GetWindowText(m_hSearchEdit, &query[0], len+1);
    query.resize(len);
    
    auto selections = m_state.GetSelections();
    size_t r = 0, c = 0;
    if (!selections.empty()) {
        r = selections[0].start.row;
        c = selections[0].start.col;
    }
    
    CsvDocument::SearchOptions opts;
    opts.forward = false;
    opts.matchCase = false;
    opts.mode = CsvDocument::SearchMode::Contains;

    bool found = false;
    {
         WaitCursor wait;
         found = m_document.Search(query, r, c, opts);
    }
    
    if (found) {
        m_state.SelectCell(r, c, false);
        m_state.SetScrollRow(r > 5 ? r - 5 : 0);
        UpdateScrollBars();
        InvalidateRect(m_hwnd, NULL, FALSE);
    } else {
         MessageBox(m_hwnd, _T("Not found"), _T("Search"), MB_OK);
    }
}

void MainWindow::OnLButtonDblClk(int x, int y, UINT keyFlags)
{
    size_t row, col;
    bool isRH, isCH;
    if (HitTest(x, y, row, col, isRH, isCH)) {
        if (!isRH && !isCH) {
            ShowCellEditor(row, col);
        }
    }
}

void MainWindow::UpdateFormulaBar()
{
    if (!m_hFormulaEdit) return;
    
    m_isUpdatingFormula = true;
    
    // Get Active Cell
    auto selections = m_state.GetSelections();
    if (selections.empty()) {
        SetWindowText(m_hFormulaEdit, L"");
        if (m_hPosLabel) SetWindowText(m_hPosLabel, L"");
        m_isUpdatingFormula = false;
        return;
    }
    
    // Check mode
    if (selections[0].mode == SelectionMode::Row || selections[0].mode == SelectionMode::Column || selections[0].mode == SelectionMode::All) {
         SetWindowText(m_hFormulaEdit, L"");
         if (m_hPosLabel) SetWindowText(m_hPosLabel, L"");
         m_isUpdatingFormula = false;
         return;
    }
    
    size_t r = selections[0].start.row;
    size_t c = selections[0].start.col;
    
    // Update Label (e.g. A1)
    if (m_hPosLabel) {
        std::wstring label;
        // Col to string: 0->A, ... 26->AA
        size_t tempC = c + 1;
        std::wstring colStr;
        while (tempC > 0) {
            size_t rem = (tempC - 1) % 26;
            colStr += (wchar_t)(L'A' + rem);
            tempC = (tempC - 1) / 26;
        }
        std::reverse(colStr.begin(), colStr.end());
        label = colStr + std::to_wstring(r + 1);
        SetWindowText(m_hPosLabel, label.c_str());
    }
    
    if (r < m_document.GetRowCount()) {
        auto cells = m_document.GetRowCells(r);
        if (c < cells.size()) {
            SetWindowText(m_hFormulaEdit, cells[c].c_str());
            m_isUpdatingFormula = false;
            return;
        }
    }
    SetWindowText(m_hFormulaEdit, L"");
    m_isUpdatingFormula = false;
}

void MainWindow::OnPaint(HWND hwnd)
{
    HRESULT hr = m_dxResources.CreateDeviceResources(hwnd);
    if (FAILED(hr)) return;

    auto pRT = m_dxResources.GetRenderTarget();
    auto pTF = m_dxResources.GetTextFormat();

    pRT->BeginDraw();
    pRT->Clear(D2D1::ColorF(D2D1::ColorF::White));

    CComPtr<ID2D1SolidColorBrush> pBrushBlack;
    pRT->CreateSolidColorBrush(D2D1::ColorF(D2D1::ColorF::Black), &pBrushBlack);
    
    CComPtr<ID2D1SolidColorBrush> pBrushGrid;
    pRT->CreateSolidColorBrush(D2D1::ColorF(D2D1::ColorF::LightGray), &pBrushGrid);

    CComPtr<ID2D1SolidColorBrush> pBrushSel;
    pRT->CreateSolidColorBrush(D2D1::ColorF(0.2f, 0.6f, 1.0f, 0.3f), &pBrushSel); // Translucent Blue
    
    CComPtr<ID2D1SolidColorBrush> pBrushHeader;
    pRT->CreateSolidColorBrush(D2D1::ColorF(D2D1::ColorF::WhiteSmoke), &pBrushHeader);

    if (pBrushBlack && pBrushGrid && pTF) {
        D2D1_SIZE_F size = pRT->GetSize();
        
        // Formula Bar Offset
        float gridY = m_formulaBarHeight; 
        
        size_t rowCount = m_document.GetRowCount();
        size_t visibleRows = (size_t)((size.height - m_headerHeight - gridY) / m_rowHeight) + 2;
        
        float startY = m_headerHeight + gridY;
        
        size_t startRow = m_state.GetScrollRow();

        // Draw Row Headers
        // Match valid rows
        bool match = true;
        // For simplicity: Iterate same rows as grid
        
        // Draw Grid and Selection
        for (size_t i = startRow; i < rowCount && i < startRow + visibleRows; ++i) {
            std::vector<std::wstring> cells = m_document.GetRowCells(i);
            float y = startY + (i - startRow) * m_rowHeight;
            
            // Draw Row Header
            D2D1_RECT_F headerRect = D2D1::RectF(0, y, m_headerWidth, y + m_rowHeight);
            pRT->FillRectangle(headerRect, pBrushHeader);
            pRT->DrawRectangle(headerRect, pBrushGrid);
            
            // Row Number
            std::wstring rowNum = std::to_wstring(i + 1);
            pRT->DrawText(rowNum.c_str(), (UINT32)rowNum.length(), pTF, headerRect, pBrushBlack);

            float scrollX = m_state.GetScrollX();
            size_t startCol = GetColumnAtX(scrollX);
            
            float x = m_headerWidth + GetColumnX(startCol) - scrollX;
            
            size_t colCount = cells.size(); 

            for (size_t col = startCol; col < colCount; ++col) {
                const auto& cell = cells[col];
                float colW = GetColumnWidth(col);
                
                D2D1_RECT_F cellRect = D2D1::RectF(x, y, x + colW, y + m_rowHeight);
                
                // Selection
                if (m_state.IsSelected(i, col)) {
                    pRT->FillRectangle(cellRect, pBrushSel);
                }

                // Draw text with Layout for Folding support
                CComPtr<IDWriteTextLayout> pLayout;
                HRESULT hrLayout = m_dxResources.GetWriteFactory()->CreateTextLayout(
                    cell.c_str(),
                    (UINT32)cell.length(),
                    pTF,
                    colW,         // Max Width
                    m_rowHeight,  // Max Height
                    &pLayout
                );

                if (SUCCEEDED(hrLayout)) {
                    if (m_maxCellLines > 1) {
                         pLayout->SetWordWrapping(DWRITE_WORD_WRAPPING_WRAP);
                    } else {
                         pLayout->SetWordWrapping(DWRITE_WORD_WRAPPING_NO_WRAP);
                    }
                    
                    // Clip visuals
                    pRT->PushAxisAlignedClip(cellRect, D2D1_ANTIALIAS_MODE_ALIASED);
                    pRT->DrawTextLayout(D2D1::Point2F(x, y), pLayout, pBrushBlack);
                    pRT->PopAxisAlignedClip();
                }

                // Draw cell border
                pRT->DrawRectangle(cellRect, pBrushGrid);
                
                x += colW;
                if (x > size.width) break;
            }
        }
        
        // Draw Column Headers
        float scrollX = m_state.GetScrollX();
        size_t startCol = GetColumnAtX(scrollX);
        float hx = m_headerWidth + GetColumnX(startCol) - scrollX;
        
        // Estimate max cols to draw
        // We iterate until off-screen
        size_t c = startCol;
        
        while (hx < size.width) {
            float colW = GetColumnWidth(c);
            
            D2D1_RECT_F colRect = D2D1::RectF(hx, gridY, hx + colW, gridY + m_headerHeight);
             pRT->FillRectangle(colRect, pBrushHeader);
             pRT->DrawRectangle(colRect, pBrushGrid);
             
             // A, B, C...
             std::wstring colStr;
             size_t tempC = c + 1;
             while (tempC > 0) {
                 size_t rem = (tempC - 1) % 26;
                 colStr += (wchar_t)(L'A' + rem);
                 tempC = (tempC - 1) / 26;
             }
             std::reverse(colStr.begin(), colStr.end());

             pRT->DrawText(colStr.c_str(), (UINT32)colStr.length(), pTF, colRect, pBrushBlack);
             
             hx += colW;
             c++;
             
             // Safety cap
             if (c > m_document.GetMaxColumnCount() + 100) break; 
        }

        // Corner Header
         D2D1_RECT_F cornerRect = D2D1::RectF(0, gridY, m_headerWidth, gridY + m_headerHeight);
         pRT->FillRectangle(cornerRect, pBrushHeader);
         pRT->DrawRectangle(cornerRect, pBrushGrid);

    }

    hr = pRT->EndDraw();
    if (hr == D2DERR_RECREATE_TARGET) {
        m_dxResources.DiscardDeviceResources();
    }
}
float MainWindow::GetColumnWidth(size_t col) const
{
    auto it = m_colWidths.find(col);
    if (it != m_colWidths.end()) return it->second;
    return m_defaultColWidth;
}

float MainWindow::GetColumnX(size_t col) const
{
    float x = 0;
    if (m_colWidths.empty()) {
        return (float)col * m_defaultColWidth;
    }
    
    for (size_t i = 0; i < col; ++i) {
        x += GetColumnWidth(i);
    }
    return x;
}

size_t MainWindow::GetColumnAtX(float targetX) const
{
    if (targetX < 0) return 0;
    
    if (m_colWidths.empty()) {
         if (m_defaultColWidth == 0) return 0;
         return (size_t)(targetX / m_defaultColWidth);
    }
    
    float x = 0;
    size_t c = 0;
    size_t maxCols = 10000; // Safety break
    
    while (x <= targetX) {
        float w = GetColumnWidth(c);
        if (x + w > targetX) return c;
        x += w;
        c++;
        if (c > maxCols) return c; 
    }
    return c;
}
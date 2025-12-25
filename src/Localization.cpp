#include "Localization.h"
#include <map>

Language Localization::s_currentLanguage = Language::English;

// Using \u escapes to avoid code page issues
static std::map<StringId, std::wstring> s_stringsJP = {
    { StringId::Menu_File, L"\u30D5\u30A1\u30A4\u30EB(&F)" }, // File
    { StringId::Menu_File_Open, L"\u958B\u304F(&O)...\tCtrl+O" }, // Open
    { StringId::Menu_File_Reopen, L"\u518D\u8AAD\u307F\u8FBC\u307F(&R)" }, // Reopen
    { StringId::Menu_File_Import, L"\u30A4\u30F3\u30DD\u30FC\u30C8(\u8FFD\u52A0)(&I)..." }, // Import (Append)...
    { StringId::Menu_File_Save, L"\u4FDD\u5B58(&S)\tCtrl+S" }, // Save
    { StringId::Menu_File_SaveAs, L"\u540D\u524D\u3092\u4ED8\u3051\u3066\u4FDD\u5B58(&A)..." }, // Save As
    { StringId::Menu_File_Export_HTML, L"HTML \u306B\u30A8\u30AF\u30B9\u30DD\u30FC\u30C8" }, // Export to HTML
    { StringId::Menu_File_Export_MD, L"Markdown \u306B\u30A8\u30AF\u30B9\u30DD\u30FC\u30C8" }, // Export to Markdown
    { StringId::Menu_File_Exit, L"\u7D42\u4E86(&X)\tAlt+F4" }, // Exit

    { StringId::Menu_Edit, L"\u7DE8\u96C6(&E)" }, // Edit
    { StringId::Menu_Edit_Undo, L"\u5143\u306B\u623B\u3059(&U)\tCtrl+Z" }, // Undo
    { StringId::Menu_Edit_Redo, L"\u3084\u308A\u76F4\u3057(&R)\tCtrl+Y" }, // Redo
    { StringId::Menu_Edit_Cut, L"\u5207\u308A\u53D6\u308A(&T)\tCtrl+X" }, // Cut
    { StringId::Menu_Edit_Copy, L"\u30B3\u30D4\u30FC(&C)\tCtrl+C" }, // Copy
    { StringId::Menu_Edit_Paste, L"\u88BC\u308A\u4ED8\u3051(&P)\tCtrl+V" }, // Paste
    { StringId::Menu_Edit_Delete, L"\u524A\u9664(&D)\tDel" }, // Delete
    { StringId::Menu_Edit_InsertRow, L"\u884C\u3092\u633F\u5165(&R)" }, // Insert Row
    { StringId::Menu_Edit_InsertCol, L"\u5217\u3092\u633F\u5165(&C)" }, // Insert Column
    { StringId::Menu_Edit_InsertRowUp, L"\u4E0A\u306B\u884C\u3092\u633F\u5165" }, // Insert Row Above
    { StringId::Menu_Edit_InsertRowDown, L"\u4E0B\u306B\u884C\u3092\u633F\u5165" }, // Insert Row Below
    { StringId::Menu_Edit_InsertColLeft, L"\u5DE6\u306B\u5217\u3092\u633F\u5165" }, // Insert Column Left
    { StringId::Menu_Edit_InsertColRight, L"\u53F3\u306B\u5217\u3092\u633F\u5165" }, // Insert Column Right
    { StringId::Menu_Edit_Clear, L"\u5185\u5BB9\u3092\u30AF\u30EA\u30A2" }, // Clear

    { StringId::Menu_Search, L"\u691C\u7D22(&S)" }, // Search
    { StringId::Menu_Search_Find, L"\u691C\u7D22(&F)...\tCtrl+F" }, // Find
    { StringId::Menu_Search_FindNext, L"\u6B21\u3092\u691C\u7D22(&N)\tF3" }, // Find Next
    { StringId::Menu_Search_FindPrev, L"\u524D\u3092\u691C\u7D22(&P)\tShift+F3" }, // Find Prev
    { StringId::Menu_Search_Replace, L"\u7F6E\u63DB(&R)...\tCtrl+H" }, // Replace

    { StringId::Menu_Config, L"\u8A2D\u5B9A(&C)" }, // Config
    { StringId::Menu_Config_Delimiter, L"\u533A\u5207\u308A\u6587\u5B57(&D)" }, // Delimiter
    { StringId::Menu_Config_Delim_Comma, L"\u30AB\u30F3\u30DE (,)" }, // Comma
    { StringId::Menu_Config_Delim_Tab, L"\u30BF\u30D6 (\\t)" }, // Tab
    { StringId::Menu_Config_Delim_Semi, L"\u30BB\u30DF\u30B3\u30ED\u30F3 (;)" }, // Semi
    { StringId::Menu_Config_Delim_Pipe, L"\u30D1\u30A4\u30D7 (|)" }, // Pipe
    { StringId::Menu_Config_Encoding, L"\u6587\u5B57\u30B3\u30FC\u30C9(&E)" }, // Encoding
    { StringId::Menu_Config_Enc_UTF8, L"UTF-8" },
    { StringId::Menu_Config_Enc_ANSI, L"ANSI" },
    { StringId::Menu_Config_Newline, L"\u6539\u884C\u30B3\u30FC\u30C9(&N)" }, // Newline
    { StringId::Menu_Config_NL_CRLF, L"CRLF (Windows)" },
    { StringId::Menu_Config_NL_LF, L"LF (Unix)" },
    { StringId::Menu_Config_Font, L"\u30D5\u30A9\u30F3\u30C8(&F)..." }, // Font
    { StringId::Menu_Config_Language, L"\u8A00\u8A9E(&L)" }, // Language
    { StringId::Menu_Config_Lang_English, L"English" },
    { StringId::Menu_Config_Lang_Japanese, L"\u65E5\u672C\u8A9E" }, // Japanese

    { StringId::Menu_Help, L"\u30D8\u30EB\u30D7(&H)" }, // Help
    { StringId::Menu_Help_View, L"\u30D8\u30EB\u30D7\u306E\u8868\u793A(&V)" }, // View Help
    { StringId::Menu_Help_About, L"\u30D0\u30FC\u30B8\u30E7\u30F3\u60C5\u5831(&A)" }, // About

    { StringId::Msg_Error, L"\u30A8\u30E9\u30FC" }, // Error
    { StringId::Msg_Info, L"\u60C5\u5831" }, // Info
    { StringId::Msg_About, L"\u30D0\u30FC\u30B8\u30E7\u30F3\u60C5\u5831" }, // About
    { StringId::Msg_AppTitle, L"CSV Editor v0.1" },
    { StringId::Msg_OpenFailed, L"\u30D5\u30A1\u30A4\u30EB\u3092\u958B\u3051\u307E\u305B\u3093\u3067\u3057\u305F\u3002" }, // Failed to open
    { StringId::Msg_SaveSuccess, L"\u4FDD\u5B58\u3057\u307E\u3057\u305F\u3002" }, // Saved
    { StringId::Msg_SaveFailed, L"\u4FDD\u5B58\u306B\u5931\u6557\u3057\u307E\u3057\u305F\u3002" }, // Save failed

    { StringId::File_Filter_CSV, L"CSV \u30D5\u30A1\u30A4\u30EB" }, // CSV Files
    { StringId::File_Filter_TSV, L"TSV \u30D5\u30A1\u30A4\u30EB" }, // TSV Files
    { StringId::File_Filter_All, L"\u3059\u3079\u3066\u306E\u30D5\u30A1\u30A4\u30EB" } // All Files
};

static std::map<StringId, std::wstring> s_stringsEN = {
    { StringId::Menu_File, L"&File" },
    { StringId::Menu_File_Open, L"&Open...\tCtrl+O" },
    { StringId::Menu_File_Reopen, L"Re&open" },
    { StringId::Menu_File_Import, L"&Import (Append)..." },
    { StringId::Menu_File_Save, L"&Save\tCtrl+S" },
    { StringId::Menu_File_SaveAs, L"Save &As..." },
    { StringId::Menu_File_Export_HTML, L"Export to HTML..." },
    { StringId::Menu_File_Export_MD, L"Export to Markdown..." },
    { StringId::Menu_File_Exit, L"E&xit\tAlt+F4" },

    { StringId::Menu_Edit, L"&Edit" },
    { StringId::Menu_Edit_Undo, L"&Undo\tCtrl+Z" },
    { StringId::Menu_Edit_Redo, L"&Redo\tCtrl+Y" },
    { StringId::Menu_Edit_Cut, L"Cu&t\tCtrl+X" },
    { StringId::Menu_Edit_Copy, L"&Copy\tCtrl+C" },
    { StringId::Menu_Edit_Paste, L"&Paste\tCtrl+V" },
    { StringId::Menu_Edit_Delete, L"&Delete\tDel" },
    { StringId::Menu_Edit_InsertRow, L"Insert &Row" },
    { StringId::Menu_Edit_InsertCol, L"Insert &Column" },
    { StringId::Menu_Edit_InsertRowUp, L"Insert Row Above" },
    { StringId::Menu_Edit_InsertRowDown, L"Insert Row Below" },
    { StringId::Menu_Edit_InsertColLeft, L"Insert Column Left" },
    { StringId::Menu_Edit_InsertColRight, L"Insert Column Right" },
    { StringId::Menu_Config_Folding, L"Toggle Text Wrapping" },
    { StringId::Menu_Edit_Clear, L"Cle&ar" },

    { StringId::Menu_Search, L"&Search" },
    { StringId::Menu_Search_Find, L"&Find...\tCtrl+F" },
    { StringId::Menu_Search_FindNext, L"Find &Next\tF3" },
    { StringId::Menu_Search_FindPrev, L"Find &Previous\tShift+F3" },
    { StringId::Menu_Search_Replace, L"&Replace...\tCtrl+H" },

    { StringId::Menu_Config, L"&Config" },
    { StringId::Menu_Config_Delimiter, L"&Delimiter" },
    { StringId::Menu_Config_Delim_Comma, L"Comma (,)" },
    { StringId::Menu_Config_Delim_Tab, L"Tab (\\t)" },
    { StringId::Menu_Config_Delim_Semi, L"Semicolon (;)" },
    { StringId::Menu_Config_Delim_Pipe, L"Pipe (|)" },
    { StringId::Menu_Config_Encoding, L"&Encoding" },
    { StringId::Menu_Config_Enc_UTF8, L"UTF-8" },
    { StringId::Menu_Config_Enc_ANSI, L"ANSI" },
    { StringId::Menu_Config_Newline, L"&Newline" },
    { StringId::Menu_Config_NL_CRLF, L"CRLF (Windows)" },
    { StringId::Menu_Config_NL_LF, L"LF (Unix)" },
    { StringId::Menu_Config_Font, L"&Font..." },
    { StringId::Menu_Config_Language, L"&Language" },
    { StringId::Menu_Config_Lang_English, L"English" },
    { StringId::Menu_Config_Lang_Japanese, L"\u65E5\u672C\u8A9E" }, // Japanese

    { StringId::Menu_Help, L"&Help" },
    { StringId::Menu_Help_View, L"&View Help" },
    { StringId::Menu_Help_About, L"&About" },

    { StringId::Msg_Error, L"Error" },
    { StringId::Msg_Info, L"Info" },
    { StringId::Msg_About, L"About" },
    { StringId::Msg_AppTitle, L"CSV Editor v0.1" },
    { StringId::Msg_OpenFailed, L"Failed to open file." },
    { StringId::Msg_SaveSuccess, L"File saved." },
    { StringId::Msg_SaveFailed, L"Failed to save file." },

    { StringId::File_Filter_CSV, L"CSV Files" },
    { StringId::File_Filter_TSV, L"TSV Files" },
    { StringId::File_Filter_All, L"All Files" }
};

void Localization::SetLanguage(Language lang)
{
    s_currentLanguage = lang;
}

Language Localization::GetLanguage()
{
    return s_currentLanguage;
}

const TCHAR* Localization::GetString(StringId id)
{
    if (s_currentLanguage == Language::Japanese) {
        if (s_stringsJP.count(id)) return s_stringsJP[id].c_str();
    }
    // Default to English
    if (s_stringsEN.count(id)) return s_stringsEN[id].c_str();
    return _T("?");
}

#include "ConfigManager.h"
#include <shlwapi.h>
#include <cstdio>

#pragma comment(lib, "Shlwapi.lib")

ConfigManager& ConfigManager::Instance() {
    static ConfigManager instance;
    return instance;
}

ConfigManager::ConfigManager() {
    m_configPath = GetConfigPath();
}

std::wstring ConfigManager::GetConfigPath() {
    wchar_t path[MAX_PATH];
    if (SUCCEEDED(SHGetFolderPath(NULL, CSIDL_MYDOCUMENTS, NULL, 0, path))) {
        std::wstring configDir = std::wstring(path) + L"\\CSVEditor";
        CreateDirectory(configDir.c_str(), NULL);
        return configDir + L"\\config.ini";
    }
    return L"config.ini"; // Fallback to current dir
}

void ConfigManager::Load() {
    // Nothing to do explicitly, GetPrivateProfileString reads from file on demand or cached by OS
}

void ConfigManager::Save() {
    // Nothing to do explicitly, WritePrivateProfileString writes to file
}

std::wstring ConfigManager::GetString(const std::wstring& section, const std::wstring& key, const std::wstring& defaultVal) {
    wchar_t buf[1024];
    GetPrivateProfileString(section.c_str(), key.c_str(), defaultVal.c_str(), buf, 1024, m_configPath.c_str());
    return buf;
}

int ConfigManager::GetInt(const std::wstring& section, const std::wstring& key, int defaultVal) {
    return GetPrivateProfileInt(section.c_str(), key.c_str(), defaultVal, m_configPath.c_str());
}

float ConfigManager::GetFloat(const std::wstring& section, const std::wstring& key, float defaultVal) {
    std::wstring val = GetString(section, key, L"");
    if (val.empty()) return defaultVal;
    try {
        return std::stof(val);
    } catch (...) {
        return defaultVal;
    }
}

void ConfigManager::SetString(const std::wstring& section, const std::wstring& key, const std::wstring& val) {
    WritePrivateProfileString(section.c_str(), key.c_str(), val.c_str(), m_configPath.c_str());
}

void ConfigManager::SetInt(const std::wstring& section, const std::wstring& key, int val) {
    SetString(section, key, std::to_wstring(val));
}

void ConfigManager::SetFloat(const std::wstring& section, const std::wstring& key, float val) {
    SetString(section, key, std::to_wstring(val));
}

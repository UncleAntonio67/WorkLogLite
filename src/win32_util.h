#pragma once

#ifndef UNICODE
#  define UNICODE
#endif

#include <windows.h>

#include <string>

std::wstring GetExeDir();
std::wstring GetLocalAppDataDir();
std::wstring JoinPath(const std::wstring& a, const std::wstring& b);
bool EnsureDirExists(const std::wstring& dir);
bool FileExists(const std::wstring& path);

std::wstring Utf8ToWide(const std::string& s);
std::string WideToUtf8(const std::wstring& s);

// Base64 helpers using Windows crypt32.
std::string Base64Encode(const std::string& bytes);
bool Base64Decode(const std::string& b64, std::string* out_bytes);

std::wstring FormatDateYYYYMMDD(const SYSTEMTIME& st);
std::wstring FormatDateYYYYMMDD(int year, int month, int day);
bool ParseYYYYMMDD(const std::wstring& s, SYSTEMTIME* out);

int DaysInMonth(int year, int month);
bool IsLeapYear(int year);
SYSTEMTIME AddDays(const SYSTEMTIME& st, int days);

int GetWindowDpi(HWND hwnd);
int ScalePx(HWND hwnd, int px96);

void SetControlFont(HWND hwnd, HFONT font);
HFONT CreateUiFont(HWND hwnd);

void ShowLastErrorBox(HWND parent, const wchar_t* context);
void ShowInfoBox(HWND parent, const wchar_t* text, const wchar_t* title);

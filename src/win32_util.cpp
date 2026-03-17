#include "win32_util.h"

#include <wincrypt.h>
#include <shlobj.h>
#include <shlwapi.h>

#include <vector>

#pragma comment(lib, "shlwapi.lib")

std::wstring GetExeDir() {
  wchar_t path[MAX_PATH]{};
  DWORD len = GetModuleFileNameW(nullptr, path, MAX_PATH);
  if (len == 0 || len >= MAX_PATH) return L".";
  PathRemoveFileSpecW(path);
  return std::wstring(path);
}

std::wstring GetLocalAppDataDir() {
  PWSTR p = nullptr;
  HRESULT hr = SHGetKnownFolderPath(FOLDERID_LocalAppData, KF_FLAG_DEFAULT, nullptr, &p);
  if (FAILED(hr) || !p) return L".";
  std::wstring out(p);
  CoTaskMemFree(p);
  return out;
}

std::wstring JoinPath(const std::wstring& a, const std::wstring& b) {
  if (a.empty()) return b;
  if (b.empty()) return a;
  wchar_t buf[MAX_PATH]{};
  if (PathCombineW(buf, a.c_str(), b.c_str()) == nullptr) {
    return a + L"\\" + b;
  }
  return std::wstring(buf);
}

bool EnsureDirExists(const std::wstring& dir) {
  if (dir.empty()) return false;
  int r = SHCreateDirectoryExW(nullptr, dir.c_str(), nullptr);
  return (r == ERROR_SUCCESS || r == ERROR_FILE_EXISTS || r == ERROR_ALREADY_EXISTS);
}

bool FileExists(const std::wstring& path) {
  DWORD attr = GetFileAttributesW(path.c_str());
  return (attr != INVALID_FILE_ATTRIBUTES) && ((attr & FILE_ATTRIBUTE_DIRECTORY) == 0);
}

std::wstring Utf8ToWide(const std::string& s) {
  if (s.empty()) return L"";
  int needed = MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0);
  if (needed <= 0) return L"";
  std::wstring out;
  out.resize((size_t)needed);
  MultiByteToWideChar(CP_UTF8, 0, s.data(), (int)s.size(), out.data(), needed);
  return out;
}

std::string WideToUtf8(const std::wstring& s) {
  if (s.empty()) return "";
  int needed = WideCharToMultiByte(CP_UTF8, 0, s.data(), (int)s.size(), nullptr, 0, nullptr, nullptr);
  if (needed <= 0) return "";
  std::string out;
  out.resize((size_t)needed);
  WideCharToMultiByte(CP_UTF8, 0, s.data(), (int)s.size(), out.data(), needed, nullptr, nullptr);
  return out;
}

std::wstring FormatDateYYYYMMDD(int year, int month, int day) {
  wchar_t buf[16]{};
  wsprintfW(buf, L"%04d-%02d-%02d", year, month, day);
  return std::wstring(buf);
}

std::wstring FormatDateYYYYMMDD(const SYSTEMTIME& st) {
  return FormatDateYYYYMMDD((int)st.wYear, (int)st.wMonth, (int)st.wDay);
}

bool ParseYYYYMMDD(const std::wstring& s, SYSTEMTIME* out) {
  if (!out) return false;
  if (s.size() != 10) return false;
  int y = 0, m = 0, d = 0;
  if (swscanf_s(s.c_str(), L"%d-%d-%d", &y, &m, &d) != 3) return false;
  if (y < 1601 || m < 1 || m > 12) return false;
  int dim = DaysInMonth(y, m);
  if (d < 1 || d > dim) return false;
  SYSTEMTIME st{};
  st.wYear = (WORD)y;
  st.wMonth = (WORD)m;
  st.wDay = (WORD)d;
  *out = st;
  return true;
}

bool IsLeapYear(int year) {
  if (year % 400 == 0) return true;
  if (year % 100 == 0) return false;
  return (year % 4 == 0);
}

int DaysInMonth(int year, int month) {
  static const int days[12] = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
  if (month < 1 || month > 12) return 30;
  if (month == 2) return IsLeapYear(year) ? 29 : 28;
  return days[month - 1];
}

SYSTEMTIME AddDays(const SYSTEMTIME& st, int days) {
  FILETIME ft{};
  SYSTEMTIME tmp = st;
  SystemTimeToFileTime(&tmp, &ft);
  ULARGE_INTEGER ui{};
  ui.LowPart = ft.dwLowDateTime;
  ui.HighPart = ft.dwHighDateTime;
  // 1 day = 86400 sec = 86400 * 10^7 100ns
  const ULONGLONG kDay = 86400ULL * 10000000ULL;
  if (days >= 0) ui.QuadPart += (ULONGLONG)days * kDay;
  else ui.QuadPart -= (ULONGLONG)(-days) * kDay;
  ft.dwLowDateTime = ui.LowPart;
  ft.dwHighDateTime = ui.HighPart;
  SYSTEMTIME out{};
  FileTimeToSystemTime(&ft, &out);
  out.wHour = 0;
  out.wMinute = 0;
  out.wSecond = 0;
  out.wMilliseconds = 0;
  return out;
}

int GetWindowDpi(HWND hwnd) {
  // GetDpiForWindow exists on Win10+. If missing, fallback to 96.
  HMODULE user32 = GetModuleHandleW(L"user32.dll");
  if (!user32) return 96;
  using Fn = UINT(WINAPI*)(HWND);
  auto p = reinterpret_cast<Fn>(GetProcAddress(user32, "GetDpiForWindow"));
  if (!p) return 96;
  return (int)p(hwnd);
}

int ScalePx(HWND hwnd, int px96) {
  int dpi = GetWindowDpi(hwnd);
  return MulDiv(px96, dpi, 96);
}

void SetControlFont(HWND hwnd, HFONT font) {
  if (!hwnd || !font) return;
  SendMessageW(hwnd, WM_SETFONT, (WPARAM)font, TRUE);
}

HFONT CreateUiFont(HWND hwnd) {
  LOGFONTW lf{};
  HDC hdc = GetDC(hwnd);
  int logPixelsY = hdc ? GetDeviceCaps(hdc, LOGPIXELSY) : 96;
  if (hdc) ReleaseDC(hwnd, hdc);
  lf.lfHeight = -MulDiv(10, logPixelsY, 72);
  wcscpy_s(lf.lfFaceName, L"Segoe UI");
  return CreateFontIndirectW(&lf);
}

static std::wstring FormatWin32Message(DWORD err) {
  wchar_t* buf = nullptr;
  DWORD flags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
  DWORD len = FormatMessageW(flags, nullptr, err, 0, (LPWSTR)&buf, 0, nullptr);
  std::wstring msg;
  if (len && buf) msg.assign(buf, buf + len);
  if (buf) LocalFree(buf);
  return msg;
}

void ShowLastErrorBox(HWND parent, const wchar_t* context) {
  DWORD err = GetLastError();
  std::wstring msg = L"操作失败: ";
  if (context) msg += context;
  msg += L"\n\nWin32 错误码: ";
  wchar_t code[32]{};
  wsprintfW(code, L"%lu", err);
  msg += code;
  msg += L"\n";
  msg += FormatWin32Message(err);
  MessageBoxW(parent, msg.c_str(), L"WorkLogLite", MB_OK | MB_ICONERROR);
}

void ShowInfoBox(HWND parent, const wchar_t* text, const wchar_t* title) {
  MessageBoxW(parent, text ? text : L"", title ? title : L"WorkLogLite", MB_OK | MB_ICONINFORMATION);
}

std::string Base64Encode(const std::string& bytes) {
  if (bytes.empty()) return "";
  DWORD out_len = 0;
  if (!CryptBinaryToStringA(reinterpret_cast<const BYTE*>(bytes.data()), (DWORD)bytes.size(),
                            CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, nullptr, &out_len)) {
    return "";
  }
  std::string out;
  out.resize(out_len);
  if (!CryptBinaryToStringA(reinterpret_cast<const BYTE*>(bytes.data()), (DWORD)bytes.size(),
                            CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF, out.data(), &out_len)) {
    return "";
  }
  // out_len includes null terminator.
  if (!out.empty() && out.back() == '\0') out.pop_back();
  return out;
}

bool Base64Decode(const std::string& b64, std::string* out_bytes) {
  if (!out_bytes) return false;
  out_bytes->clear();
  if (b64.empty()) return true;

  DWORD needed = 0;
  if (!CryptStringToBinaryA(b64.c_str(), (DWORD)b64.size(), CRYPT_STRING_BASE64, nullptr, &needed, nullptr, nullptr)) {
    return false;
  }
  std::string bytes;
  bytes.resize(needed);
  if (!CryptStringToBinaryA(b64.c_str(), (DWORD)b64.size(), CRYPT_STRING_BASE64,
                            reinterpret_cast<BYTE*>(bytes.data()), &needed, nullptr, nullptr)) {
    return false;
  }
  bytes.resize(needed);
  *out_bytes = std::move(bytes);
  return true;
}

#include "data_ops.h"

#include "win32_util.h"

#include <windows.h>
#include <shlwapi.h>

#include <cwctype>

namespace {

bool DirExistsLoose(const std::wstring& path) {
  DWORD attr = GetFileAttributesW(path.c_str());
  return (attr != INVALID_FILE_ATTRIBUTES) && ((attr & FILE_ATTRIBUTE_DIRECTORY) != 0);
}

bool PathExistsLoose(const std::wstring& path) {
  return GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES;
}

std::wstring JoinPathLoose(const std::wstring& a, const std::wstring& b) {
  if (a.empty()) return b;
  if (b.empty()) return a;
  wchar_t tail = a.back();
  if (tail == L'\\' || tail == L'/') return a + b;
  return a + L"\\" + b;
}

std::wstring ParentDirLoose(const std::wstring& path) {
  wchar_t buf[MAX_PATH * 4]{};
  wcsncpy_s(buf, _countof(buf), path.c_str(), _TRUNCATE);
  if (!PathRemoveFileSpecW(buf)) return L"";
  return std::wstring(buf);
}

std::wstring NormalizeDirForCompare(const std::wstring& in) {
  if (in.empty()) return in;
  wchar_t full[MAX_PATH * 4]{};
  DWORD n = GetFullPathNameW(in.c_str(), (DWORD)_countof(full), full, nullptr);
  std::wstring s = (n > 0 && n < _countof(full)) ? std::wstring(full) : in;
  while (!s.empty() && (s.back() == L'\\' || s.back() == L'/')) s.pop_back();
  for (auto& ch : s) ch = (wchar_t)towlower(ch);
  return s;
}

bool IsSameOrChildDir(const std::wstring& parent, const std::wstring& child) {
  std::wstring p = NormalizeDirForCompare(parent);
  std::wstring c = NormalizeDirForCompare(child);
  if (p.empty() || c.empty()) return false;
  if (c == p) return true;
  if (c.size() <= p.size()) return false;
  if (c.compare(0, p.size(), p) != 0) return false;
  wchar_t next = c[p.size()];
  return next == L'\\' || next == L'/';
}

std::wstring FormatWin32ErrorMessage(DWORD err) {
  wchar_t* buf = nullptr;
  DWORD flags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
  DWORD len = FormatMessageW(flags, nullptr, err, 0, (LPWSTR)&buf, 0, nullptr);
  std::wstring msg;
  if (len && buf) msg.assign(buf, buf + len);
  if (buf) LocalFree(buf);
  while (!msg.empty() && (msg.back() == L'\r' || msg.back() == L'\n')) msg.pop_back();
  return msg;
}

bool CopyDirRecursiveImpl(const std::wstring& src_dir, const std::wstring& dst_dir, std::wstring* err) {
  if (err) err->clear();
  if (!DirExistsLoose(src_dir)) {
    if (err) *err = L"Source folder not found: " + src_dir;
    return false;
  }
  if (!EnsureDirExists(dst_dir)) {
    if (err) *err = L"Failed to create folder: " + dst_dir;
    return false;
  }

  WIN32_FIND_DATAW fd{};
  HANDLE h = FindFirstFileW(JoinPathLoose(src_dir, L"*").c_str(), &fd);
  if (h == INVALID_HANDLE_VALUE) {
    DWORD le = GetLastError();
    if (err) *err = L"Failed to list folder: " + src_dir + L" (error=" + std::to_wstring(le) + L" " + FormatWin32ErrorMessage(le) + L")";
    return false;
  }

  auto close_find = [&](HANDLE hh) { FindClose(hh); };
  for (;;) {
    const wchar_t* name = fd.cFileName;
    if (name && wcscmp(name, L".") != 0 && wcscmp(name, L"..") != 0) {
      std::wstring sp = JoinPathLoose(src_dir, name);
      std::wstring dp = JoinPathLoose(dst_dir, name);
      bool is_dir = (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0;
      bool is_reparse = (fd.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0;

      if (is_dir && !is_reparse) {
        if (!CopyDirRecursiveImpl(sp, dp, err)) {
          close_find(h);
          return false;
        }
      } else {
        if (!CopyFileW(sp.c_str(), dp.c_str(), FALSE)) {
          DWORD le = GetLastError();
          if (err) {
            *err = L"Failed to copy file:\n" + sp + L"\n->\n" + dp +
                   L"\n(error=" + std::to_wstring(le) + L" " + FormatWin32ErrorMessage(le) + L")";
          }
          close_find(h);
          return false;
        }
      }
    }

    if (!FindNextFileW(h, &fd)) {
      DWORD le = GetLastError();
      close_find(h);
      if (le == ERROR_NO_MORE_FILES) return true;
      if (err) *err = L"Failed to enumerate folder: " + src_dir + L" (error=" + std::to_wstring(le) + L")";
      return false;
    }
  }
}

bool DeleteDirRecursiveImpl(const std::wstring& dir, std::wstring* err) {
  if (err) err->clear();
  DWORD attr = GetFileAttributesW(dir.c_str());
  if (attr == INVALID_FILE_ATTRIBUTES) return true;
  if ((attr & FILE_ATTRIBUTE_DIRECTORY) == 0) {
    if (DeleteFileW(dir.c_str()) != 0) return true;
    DWORD le = GetLastError();
    if (err) *err = L"Failed to delete file: " + dir + L" (error=" + std::to_wstring(le) + L" " + FormatWin32ErrorMessage(le) + L")";
    return false;
  }

  if ((attr & FILE_ATTRIBUTE_REPARSE_POINT) != 0) {
    if (RemoveDirectoryW(dir.c_str()) != 0) return true;
    DWORD le = GetLastError();
    if (err) *err = L"Failed to remove folder: " + dir + L" (error=" + std::to_wstring(le) + L" " + FormatWin32ErrorMessage(le) + L")";
    return false;
  }

  WIN32_FIND_DATAW fd{};
  HANDLE h = FindFirstFileW(JoinPathLoose(dir, L"*").c_str(), &fd);
  if (h != INVALID_HANDLE_VALUE) {
    auto close_find = [&](HANDLE hh) { FindClose(hh); };
    for (;;) {
      const wchar_t* name = fd.cFileName;
      if (name && wcscmp(name, L".") != 0 && wcscmp(name, L"..") != 0) {
        std::wstring child = JoinPathLoose(dir, name);
        if (!DeleteDirRecursiveImpl(child, err)) {
          close_find(h);
          return false;
        }
      }
      if (!FindNextFileW(h, &fd)) {
        DWORD le = GetLastError();
        close_find(h);
        if (le != ERROR_NO_MORE_FILES) {
          if (err) *err = L"Failed to enumerate folder: " + dir + L" (error=" + std::to_wstring(le) + L" " + FormatWin32ErrorMessage(le) + L")";
          return false;
        }
        break;
      }
    }
  }

  if (RemoveDirectoryW(dir.c_str()) != 0) return true;
  DWORD le = GetLastError();
  if (err) *err = L"Failed to remove folder: " + dir + L" (error=" + std::to_wstring(le) + L" " + FormatWin32ErrorMessage(le) + L")";
  return false;
}

bool DeleteDirContentsRecursiveImpl(const std::wstring& dir, std::wstring* err) {
  if (err) err->clear();
  if (!DirExistsLoose(dir)) return true;

  WIN32_FIND_DATAW fd{};
  HANDLE h = FindFirstFileW(JoinPathLoose(dir, L"*").c_str(), &fd);
  if (h == INVALID_HANDLE_VALUE) {
    DWORD le = GetLastError();
    if (le == ERROR_FILE_NOT_FOUND || le == ERROR_PATH_NOT_FOUND) return true;
    if (err) *err = L"Failed to list folder: " + dir + L" (error=" + std::to_wstring(le) + L" " + FormatWin32ErrorMessage(le) + L")";
    return false;
  }

  auto close_find = [&](HANDLE hh) { FindClose(hh); };
  for (;;) {
    const wchar_t* name = fd.cFileName;
    if (name && wcscmp(name, L".") != 0 && wcscmp(name, L"..") != 0) {
      std::wstring child = JoinPathLoose(dir, name);
      if (!DeleteDirRecursiveImpl(child, err)) {
        close_find(h);
        return false;
      }
    }
    if (!FindNextFileW(h, &fd)) {
      DWORD le = GetLastError();
      close_find(h);
      if (le == ERROR_NO_MORE_FILES) return true;
      if (err) *err = L"Failed to enumerate folder: " + dir + L" (error=" + std::to_wstring(le) + L" " + FormatWin32ErrorMessage(le) + L")";
      return false;
    }
  }
}

bool MoveDirNoReplace(const std::wstring& from, const std::wstring& to, std::wstring* err, DWORD* out_error = nullptr) {
  if (err) err->clear();
  if (out_error) *out_error = ERROR_SUCCESS;
  if (MoveFileExW(from.c_str(), to.c_str(), MOVEFILE_WRITE_THROUGH) != 0) return true;
  DWORD le = GetLastError();
  if (out_error) *out_error = le;
  if (err) *err = L"Move failed:\n" + from + L"\n->\n" + to + L"\n(error=" + std::to_wstring(le) + L" " + FormatWin32ErrorMessage(le) + L")";
  return false;
}

std::wstring NowTimestampCompact() {
  SYSTEMTIME st{};
  GetLocalTime(&st);
  wchar_t buf[32]{};
  wsprintfW(buf, L"%04d%02d%02d_%02d%02d%02d",
            (int)st.wYear, (int)st.wMonth, (int)st.wDay,
            (int)st.wHour, (int)st.wMinute, (int)st.wSecond);
  return std::wstring(buf);
}

std::wstring MakeUniqueChildPath(const std::wstring& parent, const std::wstring& prefix) {
  std::wstring stamp = NowTimestampCompact();
  for (int i = 0; i < 1000; ++i) {
    std::wstring name = prefix + stamp;
    if (i > 0) name += L"-" + std::to_wstring(i + 1);
    std::wstring candidate = JoinPathLoose(parent, name);
    if (!PathExistsLoose(candidate)) return candidate;
  }
  return JoinPathLoose(parent, prefix + stamp + L"-fallback");
}

bool LooksLikeDataRoot(const std::wstring& dir) {
  if (!DirExistsLoose(dir)) return false;
  if (FileExists(JoinPathLoose(dir, L"categories.txt"))) return true;
  if (FileExists(JoinPathLoose(dir, L"tasks.wlt"))) return true;
  if (FileExists(JoinPathLoose(dir, L"recurring_meetings.wlrp"))) return true;
  if (FileExists(JoinPathLoose(dir, L"auth.wla"))) return true;

  WIN32_FIND_DATAW fd{};
  HANDLE h = FindFirstFileW(JoinPathLoose(dir, L"*").c_str(), &fd);
  if (h == INVALID_HANDLE_VALUE) return false;
  bool ok = false;
  do {
    if ((fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) == 0) continue;
    if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
    const wchar_t* n = fd.cFileName;
    if (wcslen(n) == 4 && iswdigit(n[0]) && iswdigit(n[1]) && iswdigit(n[2]) && iswdigit(n[3])) {
      ok = true;
      break;
    }
  } while (FindNextFileW(h, &fd));
  FindClose(h);
  return ok;
}

}  // namespace

bool ResolveWorkLogLiteImportRoot(const std::wstring& picked_folder, std::wstring* out_root) {
  if (out_root) out_root->clear();
  if (LooksLikeDataRoot(picked_folder)) {
    if (out_root) *out_root = picked_folder;
    return true;
  }
  std::wstring sub = JoinPathLoose(picked_folder, L"data");
  if (LooksLikeDataRoot(sub)) {
    if (out_root) *out_root = sub;
    return true;
  }
  return false;
}

bool ExportWorkLogLiteData(const std::wstring& src_root,
                           const std::wstring& picked_folder,
                           std::wstring* out_export_dir,
                           std::wstring* err) {
  if (out_export_dir) out_export_dir->clear();
  if (err) err->clear();
  if (!DirExistsLoose(src_root)) {
    if (err) *err = L"Source data folder not found: " + src_root;
    return false;
  }
  if (!EnsureDirExists(picked_folder)) {
    if (err) *err = L"Failed to create export folder: " + picked_folder;
    return false;
  }

  std::wstring dst = MakeUniqueChildPath(picked_folder, L"WorkLogLite-export-");
  if (IsSameOrChildDir(src_root, dst)) {
    if (err) *err = L"Export target cannot be inside the current data folder.";
    return false;
  }
  if (!CopyDirRecursiveImpl(src_root, dst, err)) return false;
  if (out_export_dir) *out_export_dir = dst;
  return true;
}

bool ImportWorkLogLiteData(const std::wstring& picked_folder,
                           const std::wstring& dst_root,
                           std::wstring* out_backup_dir,
                           std::wstring* err) {
  if (out_backup_dir) out_backup_dir->clear();
  if (err) err->clear();

  std::wstring src_root;
  if (!ResolveWorkLogLiteImportRoot(picked_folder, &src_root)) {
    if (err) *err = L"Selected folder does not look like a WorkLogLite data directory.";
    return false;
  }
  if (NormalizeDirForCompare(src_root) == NormalizeDirForCompare(dst_root)) {
    if (err) *err = L"Import source matches the current data folder.";
    return false;
  }

  std::wstring parent = ParentDirLoose(dst_root);
  if (parent.empty()) {
    if (err) *err = L"Unable to determine the parent folder for the current data directory.";
    return false;
  }
  if (!EnsureDirExists(parent)) {
    if (err) *err = L"Failed to create/import parent folder: " + parent;
    return false;
  }

  std::wstring tmp = MakeUniqueChildPath(parent, L"WorkLogLite-data-tmp-");
  std::wstring backup = MakeUniqueChildPath(parent, L"WorkLogLite-data-backup-");
  std::wstring cleanup_err;
  DWORD move_error = ERROR_SUCCESS;

  if (!CopyDirRecursiveImpl(src_root, tmp, err)) {
    DeleteDirRecursiveImpl(tmp, &cleanup_err);
    return false;
  }

  bool had_old_root = DirExistsLoose(dst_root);
  if (had_old_root) {
    if (!MoveDirNoReplace(dst_root, backup, err, &move_error)) {
      bool can_fallback = move_error == ERROR_ACCESS_DENIED || move_error == ERROR_SHARING_VIOLATION ||
                          move_error == ERROR_LOCK_VIOLATION;
      if (!can_fallback) {
        DeleteDirRecursiveImpl(tmp, &cleanup_err);
        return false;
      }

      if (!CopyDirRecursiveImpl(dst_root, backup, err)) {
        DeleteDirRecursiveImpl(tmp, &cleanup_err);
        return false;
      }
      if (!DeleteDirContentsRecursiveImpl(dst_root, err)) {
        DeleteDirRecursiveImpl(tmp, &cleanup_err);
        return false;
      }
    }
  } else if (!EnsureDirExists(dst_root)) {
    if (err) *err = L"Failed to create destination data folder: " + dst_root;
    DeleteDirRecursiveImpl(tmp, &cleanup_err);
    return false;
  }

  if (DirExistsLoose(dst_root)) {
    if (!CopyDirRecursiveImpl(tmp, dst_root, err)) {
      if (had_old_root) {
        std::wstring rollback_err;
        DeleteDirContentsRecursiveImpl(dst_root, &rollback_err);
        CopyDirRecursiveImpl(backup, dst_root, &rollback_err);
      }
      DeleteDirRecursiveImpl(tmp, &cleanup_err);
      return false;
    }
    DeleteDirRecursiveImpl(tmp, &cleanup_err);
  } else {
    if (!MoveDirNoReplace(tmp, dst_root, err)) {
      DeleteDirRecursiveImpl(tmp, &cleanup_err);
      if (had_old_root) {
        std::wstring rollback_err;
        MoveDirNoReplace(backup, dst_root, &rollback_err);
      }
      return false;
    }
  }

  if (out_backup_dir && had_old_root) *out_backup_dir = backup;
  return true;
}

bool ClearWorkLogLiteData(const std::wstring& root,
                          std::wstring* out_backup_dir,
                          std::wstring* err) {
  if (out_backup_dir) out_backup_dir->clear();
  if (err) err->clear();

  std::wstring parent = ParentDirLoose(root);
  if (parent.empty()) {
    if (err) *err = L"Unable to determine the parent folder for the current data directory.";
    return false;
  }
  if (!EnsureDirExists(parent)) {
    if (err) *err = L"Failed to create data parent folder: " + parent;
    return false;
  }
  if (!EnsureDirExists(root)) {
    if (err) *err = L"Failed to create data folder before clearing: " + root;
    return false;
  }

  std::wstring backup = MakeUniqueChildPath(parent, L"WorkLogLite-data-cleared-");
  DWORD move_error = ERROR_SUCCESS;
  if (!MoveDirNoReplace(root, backup, err, &move_error)) {
    bool can_fallback = move_error == ERROR_ACCESS_DENIED || move_error == ERROR_SHARING_VIOLATION ||
                        move_error == ERROR_LOCK_VIOLATION;
    if (!can_fallback) return false;
    if (!CopyDirRecursiveImpl(root, backup, err)) return false;
    if (!DeleteDirContentsRecursiveImpl(root, err)) return false;
    if (out_backup_dir) *out_backup_dir = backup;
    return true;
  }

  if (!EnsureDirExists(root)) {
    std::wstring rollback_err;
    MoveDirNoReplace(backup, root, &rollback_err);
    if (err && err->empty()) *err = L"Failed to recreate the empty data folder: " + root;
    return false;
  }

  if (out_backup_dir) *out_backup_dir = backup;
  return true;
}

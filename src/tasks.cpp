#include "tasks.h"

#include "storage.h"
#include "win32_util.h"

#include <sstream>

static bool DateLeq(const SYSTEMTIME& a, const SYSTEMTIME& b) {
  if (a.wYear != b.wYear) return a.wYear < b.wYear;
  if (a.wMonth != b.wMonth) return a.wMonth < b.wMonth;
  return a.wDay <= b.wDay;
}

static bool DateEq(const SYSTEMTIME& a, const SYSTEMTIME& b) {
  return a.wYear == b.wYear && a.wMonth == b.wMonth && a.wDay == b.wDay;
}

static bool DateGeq(const SYSTEMTIME& a, const SYSTEMTIME& b) {
  return DateLeq(b, a);
}

std::wstring GetTasksFilePath() {
  return JoinPath(GetDataRootDir(), L"tasks.wlt");
}

static const char* StatusToStr(EntryStatus s) {
  switch (s) {
    case EntryStatus::Todo: return "todo";
    case EntryStatus::Doing: return "doing";
    case EntryStatus::Blocked: return "blocked";
    case EntryStatus::Done: return "done";
    case EntryStatus::None:
    default: return "none";
  }
}

static EntryStatus StatusFromStr(const std::wstring& s) {
  if (s == L"todo") return EntryStatus::Todo;
  if (s == L"doing") return EntryStatus::Doing;
  if (s == L"blocked") return EntryStatus::Blocked;
  if (s == L"done") return EntryStatus::Done;
  return EntryStatus::Todo;
}

static bool ReadFileUtf8(const std::wstring& path, std::wstring* out, std::wstring* err) {
  if (!out) return false;
  FILE* f = nullptr;
  _wfopen_s(&f, path.c_str(), L"rb");
  if (!f) {
    if (err) *err = L"Failed to open file: " + path;
    return false;
  }
  fseek(f, 0, SEEK_END);
  long size = ftell(f);
  fseek(f, 0, SEEK_SET);
  if (size < 0) size = 0;
  std::string bytes;
  bytes.resize((size_t)size);
  if (size > 0) {
    size_t n = fread(bytes.data(), 1, (size_t)size, f);
    bytes.resize(n);
  }
  fclose(f);
  *out = Utf8ToWide(bytes);
  return true;
}

static bool WriteFileUtf8Atomic(const std::wstring& path, const std::string& bytes, std::wstring* err) {
  std::wstring tmp = path + L".tmp";
  {
    FILE* f = nullptr;
    _wfopen_s(&f, tmp.c_str(), L"wb");
    if (!f) {
      if (err) *err = L"Failed to write temp file: " + tmp;
      return false;
    }
    size_t n = fwrite(bytes.data(), 1, bytes.size(), f);
    fflush(f);
    fclose(f);
    if (n != bytes.size()) {
      if (err) *err = L"Failed to write temp file: " + tmp;
      return false;
    }
  }
  if (!MoveFileExW(tmp.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
    if (err) *err = L"Failed to replace file: " + path;
    DeleteFileW(tmp.c_str());
    return false;
  }
  return true;
}

static bool FileExists2(const std::wstring& path) {
  return FileExists(path);
}

static bool ParseKV(const std::wstring& line, std::wstring* out_k, std::wstring* out_v) {
  size_t eq = line.find(L'=');
  if (eq == std::wstring::npos) return false;
  if (out_k) *out_k = line.substr(0, eq);
  if (out_v) *out_v = line.substr(eq + 1);
  return true;
}

bool LoadTasks(std::vector<Task>* out, std::wstring* err) {
  if (!out) return false;
  out->clear();

  EnsureDirExists(GetDataRootDir());
  std::wstring path = GetTasksFilePath();
  if (!FileExists2(path)) return true;

  std::wstring content;
  if (!ReadFileUtf8(path, &content, err)) return false;

  std::wstringstream ss(content);
  std::wstring line;
  Task cur{};
  bool in_task = false;
  int version = 1;

  auto flush = [&]() {
    if (!in_task) return;
    if (!cur.id.empty() && !cur.title.empty()) out->push_back(cur);
    cur = Task{};
    in_task = false;
  };

  while (std::getline(ss, line)) {
    if (!line.empty() && line.back() == L'\r') line.pop_back();
    if (line.empty()) continue;
    if (line == L"WLT1") {
      version = 1;
      continue;
    }
    if (line == L"WLT2") {
      version = 2;
      continue;
    }
    if (line == L"TASK") {
      flush();
      in_task = true;
      continue;
    }
    if (line == L"END") {
      flush();
      continue;
    }
    if (!in_task) continue;

    std::wstring k, v;
    if (!ParseKV(line, &k, &v)) continue;
    if (k == L"id") cur.id = v;
    else if (k == L"start") ParseYYYYMMDD(v, &cur.start);
    else if (k == L"end") ParseYYYYMMDD(v, &cur.end);
    else if (k == L"status") cur.status = StatusFromStr(v);
    else if (k == L"basis_b64") {
      std::string bytes;
      if (Base64Decode(WideToUtf8(v), &bytes)) cur.basis = Utf8ToWide(bytes);
    }
    else if (k == L"desc_plain_b64") {
      std::string bytes;
      if (Base64Decode(WideToUtf8(v), &bytes)) cur.desc_plain = Utf8ToWide(bytes);
    }
    else if (k == L"desc_rtf_b64") {
      // Already base64; store as-is
      cur.desc_rtf_b64 = WideToUtf8(v);
    }
    else if (k == L"category_b64") {
      std::string bytes;
      if (Base64Decode(WideToUtf8(v), &bytes)) cur.category = Utf8ToWide(bytes);
    } else if (k == L"title_b64") {
      std::string bytes;
      if (Base64Decode(WideToUtf8(v), &bytes)) cur.title = Utf8ToWide(bytes);
    }
  }
  flush();
  (void)version;
  return true;
}

bool SaveTasks(const std::vector<Task>& tasks, std::wstring* err) {
  EnsureDirExists(GetDataRootDir());
  std::wstring path = GetTasksFilePath();

  std::string out;
  out += "WLT2\n";
  for (const auto& t : tasks) {
    if (t.id.empty() || t.title.empty()) continue;
    out += "TASK\n";
    out += "id=" + WideToUtf8(t.id) + "\n";
    out += "category_b64=" + Base64Encode(WideToUtf8(t.category)) + "\n";
    out += "title_b64=" + Base64Encode(WideToUtf8(t.title)) + "\n";
    out += "basis_b64=" + Base64Encode(WideToUtf8(t.basis)) + "\n";
    out += "desc_plain_b64=" + Base64Encode(WideToUtf8(t.desc_plain)) + "\n";
    out += "desc_rtf_b64=" + t.desc_rtf_b64 + "\n";
    out += "start=" + WideToUtf8(FormatDateYYYYMMDD(t.start)) + "\n";
    out += "end=" + WideToUtf8(FormatDateYYYYMMDD(t.end)) + "\n";
    out += "status=" + std::string(StatusToStr(t.status)) + "\n";
    out += "END\n";
  }
  return WriteFileUtf8Atomic(path, out, err);
}

bool TaskActiveOn(const Task& t, const SYSTEMTIME& date) {
  // inclusive range. If start/end are zero, treat as inactive.
  if (t.start.wYear == 0 || t.end.wYear == 0) return false;
  if (DateLeq(date, t.end) && DateGeq(date, t.start)) return true;
  if (DateEq(date, t.start) || DateEq(date, t.end)) return true;
  return false;
}

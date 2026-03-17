#include "storage.h"

#include "win32_util.h"

#include <windows.h>
#include <shlwapi.h>

#include <cstdio>
#include <sstream>

static std::wstring NormalizeCategoryFromTag(const std::wstring& tag) {
  // Backward compatible: previous versions used english tags.
  if (tag == L"work") return L"工作";
  if (tag == L"meeting") return L"会议";
  if (tag == L"goal") return L"目标";
  return tag;
}

static bool StartsWith(const std::wstring& s, const std::wstring& prefix) {
  return s.size() >= prefix.size() && s.compare(0, prefix.size(), prefix) == 0;
}

static std::wstring TrimRightCR(const std::wstring& s) {
  if (!s.empty() && s.back() == L'\r') return s.substr(0, s.size() - 1);
  return s;
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

std::wstring GetDataRootDir() {
  std::wstring portable = JoinPath(GetExeDir(), L"data");
  if (EnsureDirExists(portable)) return portable;

  std::wstring fallback = JoinPath(GetLocalAppDataDir(), L"WorkLogLite\\data");
  EnsureDirExists(fallback);
  return fallback;
}

static std::wstring GetDayFilePathWithExt(const SYSTEMTIME& date, const wchar_t* ext) {
  wchar_t y[8]{}, m[8]{};
  wsprintfW(y, L"%04d", (int)date.wYear);
  wsprintfW(m, L"%02d", (int)date.wMonth);
  std::wstring rel = std::wstring(y) + L"\\" + std::wstring(m);
  std::wstring dir = JoinPath(GetDataRootDir(), rel);
  std::wstring name = FormatDateYYYYMMDD(date) + ext;
  return JoinPath(dir, name);
}

std::wstring GetDayFilePath(const SYSTEMTIME& date) {
  return GetDayFilePathWithExt(date, L".wlr");
}

static std::wstring GetLegacyDayFilePath(const SYSTEMTIME& date) {
  return GetDayFilePathWithExt(date, L".wlmd");
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
  return EntryStatus::None;
}

static const char* TypeToStr(EntryType t) {
  switch (t) {
    case EntryType::TaskProgress: return "task";
    case EntryType::Meeting: return "meeting";
    case EntryType::Note:
    default: return "note";
  }
}

static EntryType TypeFromStr(const std::wstring& s) {
  if (s == L"task") return EntryType::TaskProgress;
  if (s == L"meeting") return EntryType::Meeting;
  return EntryType::Note;
}

static bool EnsureParentDir(const std::wstring& path, std::wstring* err) {
  wchar_t dirbuf[MAX_PATH]{};
  wcsncpy_s(dirbuf, _countof(dirbuf), path.c_str(), _TRUNCATE);
  PathRemoveFileSpecW(dirbuf);
  if (!EnsureDirExists(dirbuf)) {
    if (err) *err = L"Failed to create dir: " + std::wstring(dirbuf);
    return false;
  }
  return true;
}

static bool ShouldPersistEntry(const Entry& e) {
  if (e.placeholder) return false;
  bool has_any = !e.title.empty() || !e.body_plain.empty() || !e.body_rtf_b64.empty() ||
                 !e.start_time.empty() || !e.end_time.empty() || e.status != EntryStatus::None;
  if (!has_any) return false;

  // For task placeholders, avoid writing empty records unless there's real progress/status/time.
  if (e.type == EntryType::TaskProgress) {
    bool has_task_any = !e.body_plain.empty() || !e.body_rtf_b64.empty() ||
                        !e.start_time.empty() || !e.end_time.empty() || e.status != EntryStatus::None;
    return has_task_any;
  }
  // For meeting templates, avoid persisting if only title exists.
  if (e.type == EntryType::Meeting) {
    bool has_meeting_any = !e.body_plain.empty() || !e.body_rtf_b64.empty() ||
                           !e.start_time.empty() || !e.end_time.empty();
    return has_meeting_any;
  }
  return true;
}

static void WriteB64Field(std::string& out, const char* key, const std::wstring& value) {
  out += key;
  out += "=";
  out += Base64Encode(WideToUtf8(value));
  out += "\n";
}

static void WriteB64FieldBytes(std::string& out, const char* key, const std::string& bytes) {
  out += key;
  out += "=";
  out += Base64Encode(bytes);
  out += "\n";
}

static void WriteField(std::string& out, const char* key, const std::wstring& value) {
  out += key;
  out += "=";
  out += WideToUtf8(value);
  out += "\n";
}

static void WriteFieldA(std::string& out, const char* key, const char* value) {
  out += key;
  out += "=";
  out += value ? value : "";
  out += "\n";
}

static bool ParseKV(const std::wstring& line, std::wstring* out_k, std::wstring* out_v) {
  size_t eq = line.find(L'=');
  if (eq == std::wstring::npos) return false;
  if (out_k) *out_k = line.substr(0, eq);
  if (out_v) *out_v = line.substr(eq + 1);
  return true;
}

static bool LoadDayFileWlr(const std::wstring& path, const SYSTEMTIME& date, DayData* out, std::wstring* err) {
  std::wstring content;
  if (!ReadFileUtf8(path, &content, err)) return false;

  out->date = date;
  out->entries.clear();

  std::wstringstream ss(content);
  std::wstring line;

  Entry cur{};
  bool in_entry = false;

  auto flush = [&]() {
    if (!in_entry) return;
    out->entries.push_back(cur);
    cur = Entry{};
    in_entry = false;
  };

  while (std::getline(ss, line)) {
    line = TrimRightCR(line);
    if (line.empty()) continue;

    if (line == L"WLR1") continue;
    if (line == L"ENTRY") {
      flush();
      in_entry = true;
      continue;
    }
    if (line == L"END") {
      flush();
      continue;
    }
    if (!in_entry) continue;

    std::wstring k, v;
    if (!ParseKV(line, &k, &v)) continue;

    if (k == L"id") cur.id = v;
    else if (k == L"type") cur.type = TypeFromStr(v);
    else if (k == L"status") cur.status = StatusFromStr(v);
    else if (k == L"start") cur.start_time = v;
    else if (k == L"end") cur.end_time = v;
    else if (k == L"task_id_b64") {
      std::string bytes;
      if (!Base64Decode(WideToUtf8(v), &bytes)) continue;
      cur.task_id = Utf8ToWide(bytes);
    } else if (k == L"category_b64") {
      std::string bytes;
      if (!Base64Decode(WideToUtf8(v), &bytes)) continue;
      cur.category = Utf8ToWide(bytes);
    } else if (k == L"title_b64") {
      std::string bytes;
      if (!Base64Decode(WideToUtf8(v), &bytes)) continue;
      cur.title = Utf8ToWide(bytes);
    } else if (k == L"body_plain_b64") {
      std::string bytes;
      if (!Base64Decode(WideToUtf8(v), &bytes)) continue;
      cur.body_plain = Utf8ToWide(bytes);
    } else if (k == L"body_rtf_b64") {
      cur.body_rtf_b64 = WideToUtf8(v);
    }
  }
  flush();

  // Assign ids if missing.
  for (size_t i = 0; i < out->entries.size(); i++) {
    if (!out->entries[i].id.empty()) continue;
    wchar_t buf[64]{};
    wsprintfW(buf, L"e%04d", (int)i + 1);
    out->entries[i].id = buf;
  }
  return true;
}

static bool LoadDayFileLegacyWlmd(const std::wstring& path, const SYSTEMTIME& date, DayData* out, std::wstring* err) {
  std::wstring content;
  if (!ReadFileUtf8(path, &content, err)) return false;

  out->date = date;
  out->entries.clear();

  std::wstringstream ss(content);
  std::wstring line;

  Entry current{};
  bool in_entry = false;
  std::wstring body_accum;

  auto flush_current = [&]() {
    if (!in_entry) return;
    current.type = EntryType::Note;
    current.status = EntryStatus::None;
    current.body_plain = body_accum;
    while (!current.body_plain.empty() && (current.body_plain.back() == L'\n' || current.body_plain.back() == L'\r')) {
      current.body_plain.pop_back();
    }
    out->entries.push_back(current);
    current = Entry{};
    in_entry = false;
    body_accum.clear();
  };

  while (std::getline(ss, line)) {
    line = TrimRightCR(line);

    if (StartsWith(line, L"- [")) {
      flush_current();

      size_t rb = line.find(L"]");
      if (rb == std::wstring::npos) continue;
      std::wstring tag = line.substr(3, rb - 3);
      Entry e{};
      e.category = NormalizeCategoryFromTag(tag);

      std::wstring rest = line.substr(rb + 1);
      while (!rest.empty() && rest.front() == L' ') rest.erase(rest.begin());

      if (rest.size() >= 11 && rest[2] == L':' && rest[5] == L'-' && rest[8] == L':') {
        e.start_time = rest.substr(0, 5);
        e.end_time = rest.substr(6, 5);
        rest = rest.substr(11);
        while (!rest.empty() && rest.front() == L' ') rest.erase(rest.begin());
      }

      e.title = rest;
      current = e;
      in_entry = true;
      continue;
    }

    if (!in_entry) continue;
    if (StartsWith(line, L"  ")) {
      body_accum.append(line.substr(2));
      body_accum.push_back(L'\n');
    } else if (line.empty()) {
      body_accum.push_back(L'\n');
    } else {
      body_accum.append(line);
      body_accum.push_back(L'\n');
    }
  }

  flush_current();
  return true;
}

bool LoadDayFile(const SYSTEMTIME& date, DayData* out, std::wstring* err) {
  if (!out) return false;
  out->date = date;
  out->entries.clear();

  std::wstring wlr = GetDayFilePath(date);
  if (FileExists(wlr)) return LoadDayFileWlr(wlr, date, out, err);

  std::wstring wlmd = GetLegacyDayFilePath(date);
  if (FileExists(wlmd)) return LoadDayFileLegacyWlmd(wlmd, date, out, err);

  return true;
}

bool SaveDayFile(const DayData& day, std::wstring* err) {
  std::wstring root = GetDataRootDir();
  if (!EnsureDirExists(root)) {
    if (err) *err = L"Failed to ensure data dir: " + root;
    return false;
  }

  std::wstring path = GetDayFilePath(day.date);
  if (!EnsureParentDir(path, err)) return false;

  std::string out;
  out += "WLR1\n";
  out += "date=" + WideToUtf8(FormatDateYYYYMMDD(day.date)) + "\n";

  for (const auto& e : day.entries) {
    if (!ShouldPersistEntry(e)) continue;

    out += "ENTRY\n";
    WriteField(out, "id", e.id.empty() ? L"e" : e.id);
    WriteFieldA(out, "type", TypeToStr(e.type));
    WriteFieldA(out, "status", StatusToStr(e.status));
    WriteField(out, "start", e.start_time);
    WriteField(out, "end", e.end_time);
    WriteB64Field(out, "task_id_b64", e.task_id);
    WriteB64Field(out, "category_b64", e.category);
    WriteB64Field(out, "title_b64", e.title);
    WriteB64Field(out, "body_plain_b64", e.body_plain);

    // body_rtf_b64 is already base64, store as-is.
    WriteField(out, "body_rtf_b64", Utf8ToWide(e.body_rtf_b64));
    out += "END\n";
  }

  return WriteFileUtf8Atomic(path, out, err);
}

#include "recurring.h"

#include "storage.h"
#include "win32_util.h"

#include <sstream>

static bool DateLeq(const SYSTEMTIME& a, const SYSTEMTIME& b) {
  if (a.wYear != b.wYear) return a.wYear < b.wYear;
  if (a.wMonth != b.wMonth) return a.wMonth < b.wMonth;
  return a.wDay <= b.wDay;
}

static bool DateGeq(const SYSTEMTIME& a, const SYSTEMTIME& b) {
  return DateLeq(b, a);
}

static int DayOfWeekSun0(const SYSTEMTIME& d) {
  SYSTEMTIME tmp = d;
  FILETIME ft{};
  SystemTimeToFileTime(&tmp, &ft);
  FileTimeToSystemTime(&ft, &tmp);
  // SYSTEMTIME.wDayOfWeek: 0=Sunday
  return (int)tmp.wDayOfWeek;
}

static int DaysBetween(const SYSTEMTIME& a, const SYSTEMTIME& b) {
  // b - a in days
  FILETIME fta{}, ftb{};
  SYSTEMTIME sa = a, sb = b;
  SystemTimeToFileTime(&sa, &fta);
  SystemTimeToFileTime(&sb, &ftb);
  ULARGE_INTEGER ua{}, ub{};
  ua.LowPart = fta.dwLowDateTime;
  ua.HighPart = fta.dwHighDateTime;
  ub.LowPart = ftb.dwLowDateTime;
  ub.HighPart = ftb.dwHighDateTime;
  const ULONGLONG kDay = 86400ULL * 10000000ULL;
  if (ub.QuadPart < ua.QuadPart) return -(int)((ua.QuadPart - ub.QuadPart) / kDay);
  return (int)((ub.QuadPart - ua.QuadPart) / kDay);
}

std::wstring GetRecurringMeetingsFilePath() {
  return JoinPath(GetDataRootDir(), L"recurring_meetings.wlrp");
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

static bool ParseKV(const std::wstring& line, std::wstring* out_k, std::wstring* out_v) {
  size_t eq = line.find(L'=');
  if (eq == std::wstring::npos) return false;
  if (out_k) *out_k = line.substr(0, eq);
  if (out_v) *out_v = line.substr(eq + 1);
  return true;
}

bool LoadRecurringMeetings(std::vector<RecurringMeeting>* out, std::wstring* err) {
  if (!out) return false;
  out->clear();

  EnsureDirExists(GetDataRootDir());
  std::wstring path = GetRecurringMeetingsFilePath();
  if (!FileExists(path)) return true;

  std::wstring content;
  if (!ReadFileUtf8(path, &content, err)) return false;

  std::wstringstream ss(content);
  std::wstring line;
  RecurringMeeting cur{};
  bool in = false;
  bool header_ok = false;

  auto flush = [&]() {
    if (!in) return;
    if (!cur.id.empty() && !cur.title.empty()) out->push_back(cur);
    cur = RecurringMeeting{};
    in = false;
  };

  while (std::getline(ss, line)) {
    if (!line.empty() && line.back() == L'\r') line.pop_back();
    if (line.empty()) continue;
    if (!header_ok) {
      header_ok = (line == L"WLRP1");
      continue;
    }
    if (line == L"REC") {
      flush();
      in = true;
      continue;
    }
    if (line == L"END") {
      flush();
      continue;
    }
    if (!in) continue;

    std::wstring k, v;
    if (!ParseKV(line, &k, &v)) continue;
    if (k == L"id") cur.id = v;
    else if (k == L"weekday") cur.weekday = _wtoi(v.c_str());
    else if (k == L"interval_weeks") cur.interval_weeks = _wtoi(v.c_str());
    else if (k == L"start") ParseYYYYMMDD(v, &cur.start);
    else if (k == L"end") ParseYYYYMMDD(v, &cur.end);
    else if (k == L"start_time") cur.start_time = v;
    else if (k == L"end_time") cur.end_time = v;
    else if (k == L"category_b64") {
      std::string bytes;
      if (Base64Decode(WideToUtf8(v), &bytes)) cur.category = Utf8ToWide(bytes);
    } else if (k == L"title_b64") {
      std::string bytes;
      if (Base64Decode(WideToUtf8(v), &bytes)) cur.title = Utf8ToWide(bytes);
    } else if (k == L"template_b64") {
      std::string bytes;
      if (Base64Decode(WideToUtf8(v), &bytes)) cur.template_plain = Utf8ToWide(bytes);
    }
  }
  flush();

  // Normalize
  for (auto& m : *out) {
    if (m.interval_weeks <= 0) m.interval_weeks = 1;
    if (m.weekday < 0) m.weekday = 0;
    if (m.weekday > 6) m.weekday = 6;
  }
  return true;
}

bool SaveRecurringMeetings(const std::vector<RecurringMeeting>& items, std::wstring* err) {
  EnsureDirExists(GetDataRootDir());
  std::wstring path = GetRecurringMeetingsFilePath();

  std::string out;
  out += "WLRP1\n";
  for (const auto& m : items) {
    if (m.id.empty() || m.title.empty()) continue;
    out += "REC\n";
    out += "id=" + WideToUtf8(m.id) + "\n";
    out += "weekday=" + std::to_string(m.weekday) + "\n";
    out += "interval_weeks=" + std::to_string(m.interval_weeks) + "\n";
    out += "start=" + WideToUtf8(FormatDateYYYYMMDD(m.start)) + "\n";
    out += "end=" + WideToUtf8(FormatDateYYYYMMDD(m.end)) + "\n";
    out += "start_time=" + WideToUtf8(m.start_time) + "\n";
    out += "end_time=" + WideToUtf8(m.end_time) + "\n";
    out += "category_b64=" + Base64Encode(WideToUtf8(m.category)) + "\n";
    out += "title_b64=" + Base64Encode(WideToUtf8(m.title)) + "\n";
    out += "template_b64=" + Base64Encode(WideToUtf8(m.template_plain)) + "\n";
    out += "END\n";
  }
  return WriteFileUtf8Atomic(path, out, err);
}

bool RecurringMeetingOccursOn(const RecurringMeeting& m, const SYSTEMTIME& date) {
  if (m.start.wYear == 0 || m.end.wYear == 0) return false;
  if (!DateGeq(date, m.start) || !DateLeq(date, m.end)) return false;
  if (DayOfWeekSun0(date) != m.weekday) return false;
  if (m.interval_weeks <= 1) return true;
  int days = DaysBetween(m.start, date);
  if (days < 0) return false;
  int weeks = days / 7;
  return (weeks % m.interval_weeks) == 0;
}


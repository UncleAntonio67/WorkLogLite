#include "report.h"

#include "storage.h"
#include "tasks.h"
#include "win32_util.h"

#include <algorithm>
#include <map>
#include <sstream>
#include <string>
#include <vector>

static SYSTEMTIME MakeDate(int y, int m, int d) {
  SYSTEMTIME st{};
  st.wYear = (WORD)y;
  st.wMonth = (WORD)m;
  st.wDay = (WORD)d;
  return st;
}

ReportRange QuarterRangeForDate(const SYSTEMTIME& selected) {
  int y = (int)selected.wYear;
  int m = (int)selected.wMonth;
  int q = (m - 1) / 3 + 1;
  int sm = (q - 1) * 3 + 1;
  int em = sm + 2;
  ReportRange r{};
  r.start = MakeDate(y, sm, 1);
  r.end = MakeDate(y, em, DaysInMonth(y, em));
  return r;
}

ReportRange YearRangeForDate(const SYSTEMTIME& selected) {
  int y = (int)selected.wYear;
  ReportRange r{};
  r.start = MakeDate(y, 1, 1);
  r.end = MakeDate(y, 12, 31);
  return r;
}

static bool DateLeq(const SYSTEMTIME& a, const SYSTEMTIME& b) {
  if (a.wYear != b.wYear) return a.wYear < b.wYear;
  if (a.wMonth != b.wMonth) return a.wMonth < b.wMonth;
  return a.wDay <= b.wDay;
}

static const wchar_t* StatusToCN(EntryStatus s) {
  switch (s) {
    case EntryStatus::Todo: return L"未开始";
    case EntryStatus::Doing: return L"进行中";
    case EntryStatus::Blocked: return L"阻塞";
    case EntryStatus::Done: return L"已完成";
    case EntryStatus::None:
    default: return L"";
  }
}

static std::wstring TimeRange(const Entry& e) {
  if (e.start_time.empty() && e.end_time.empty()) return L"";
  std::wstring st = e.start_time.empty() ? L"??:??" : e.start_time;
  std::wstring et = e.end_time.empty() ? L"??:??" : e.end_time;
  return st + L"-" + et;
}

static int ParseHHMMMinutes(const std::wstring& t) {
  // Returns minutes since 00:00, or -1 if invalid/empty.
  if (t.size() != 5) return -1;
  if (t[2] != L':') return -1;
  if (t[0] < L'0' || t[0] > L'9') return -1;
  if (t[1] < L'0' || t[1] > L'9') return -1;
  if (t[3] < L'0' || t[3] > L'9') return -1;
  if (t[4] < L'0' || t[4] > L'9') return -1;
  int hh = (t[0] - L'0') * 10 + (t[1] - L'0');
  int mm = (t[3] - L'0') * 10 + (t[4] - L'0');
  if (hh < 0 || hh > 23 || mm < 0 || mm > 59) return -1;
  return hh * 60 + mm;
}

static void EmitBodyIndented(std::wstringstream* md, const std::wstring& body) {
  if (!md) return;
  if (body.empty()) return;
  std::wstringstream bss(body);
  std::wstring line;
  while (std::getline(bss, line)) {
    if (!line.empty() && line.back() == L'\r') line.pop_back();
    (*md) << L"  " << line << L"\n";
  }
}

bool GenerateReportMarkdown(const ReportRange& range, std::wstring* out_md, std::wstring* err) {
  if (!out_md) return false;
  out_md->clear();

  SYSTEMTIME cur = range.start;
  std::wstringstream md;
  md << L"# 工作汇总 " << FormatDateYYYYMMDD(range.start) << L" ~ " << FormatDateYYYYMMDD(range.end) << L"\n\n";

  while (DateLeq(cur, range.end)) {
    DayData day{};
    std::wstring load_err;
    if (!LoadDayFile(cur, &day, &load_err)) {
      if (err) *err = L"读取失败: " + FormatDateYYYYMMDD(cur) + L"\n" + load_err;
      return false;
    }

    if (!day.entries.empty()) {
      md << L"## " << FormatDateYYYYMMDD(cur) << L"\n";
      for (const auto& e : day.entries) {
        md << L"- [" << (e.category.empty() ? L"工作" : e.category) << L"] ";
        std::wstring tr = TimeRange(e);
        if (!tr.empty()) md << tr << L" ";
        md << (e.title.empty() ? L"(无标题)" : e.title) << L"\n";
        EmitBodyIndented(&md, e.body_plain);
      }
      md << L"\n";
    }

    cur = AddDays(cur, 1);
  }

  *out_md = md.str();
  return true;
}

struct CatRecord {
  SYSTEMTIME date{};
  EntryType type{EntryType::Note};
  std::wstring time;
  std::wstring title;
  std::wstring body;
};

struct TaskBucket {
  Task task{};
  bool has_task{false};
  std::vector<CatRecord> records;
};

bool GenerateReportMarkdownByCategory(const ReportRange& range, std::wstring* out_md, std::wstring* err) {
  if (!out_md) return false;
  out_md->clear();

  // Load tasks to enrich task progress grouping.
  std::vector<Task> tasks;
  {
    std::wstring terr;
    LoadTasks(&tasks, &terr);
  }
  std::map<std::wstring, Task> task_by_id;
  for (const auto& t : tasks) task_by_id[t.id] = t;

  // category -> task_id -> bucket
  std::map<std::wstring, std::map<std::wstring, TaskBucket>> task_progress;
  std::map<std::wstring, std::vector<CatRecord>> meetings;
  std::map<std::wstring, std::vector<CatRecord>> notes;

  SYSTEMTIME cur = range.start;
  while (DateLeq(cur, range.end)) {
    DayData day{};
    std::wstring load_err;
    if (!LoadDayFile(cur, &day, &load_err)) {
      if (err) *err = L"读取失败: " + FormatDateYYYYMMDD(cur) + L"\n" + load_err;
      return false;
    }

    for (const auto& e : day.entries) {
      std::wstring cat = e.category.empty() ? L"(未分类)" : e.category;
      CatRecord r{};
      r.date = cur;
      r.type = e.type;
      r.time = TimeRange(e);
      r.title = e.title.empty() ? L"(无标题)" : e.title;
      r.body = e.body_plain;

      if (e.type == EntryType::TaskProgress && !e.task_id.empty()) {
        std::wstring tid = e.task_id;
        TaskBucket& b = task_progress[cat][tid];

        auto it = task_by_id.find(tid);
        if (it != task_by_id.end()) {
          b.task = it->second;
          b.has_task = true;
          if (!it->second.category.empty()) cat = it->second.category;
        } else {
          // Missing task metadata: still group by task_id under current category.
          b.has_task = false;
        }

        // If task category overrides, ensure bucket is under the effective category map.
        if (b.has_task && !b.task.category.empty() && b.task.category != cat) {
          TaskBucket& b2 = task_progress[b.task.category][tid];
          if (!b2.has_task) {
            b2.task = b.task;
            b2.has_task = true;
          }
          b2.records.push_back(std::move(r));
        } else {
          b.records.push_back(std::move(r));
        }
      } else if (e.type == EntryType::Meeting) {
        meetings[cat].push_back(std::move(r));
      } else {
        notes[cat].push_back(std::move(r));
      }
    }

    cur = AddDays(cur, 1);
  }

  // Merge keys.
  std::map<std::wstring, bool> cats;
  for (const auto& kv : task_progress) cats[kv.first] = true;
  for (const auto& kv : meetings) cats[kv.first] = true;
  for (const auto& kv : notes) cats[kv.first] = true;

  std::wstringstream md;
  md << L"# 工作汇总(按分类) " << FormatDateYYYYMMDD(range.start) << L" ~ " << FormatDateYYYYMMDD(range.end) << L"\n\n";

  auto emit_records = [&](const std::vector<CatRecord>& rs) {
    for (const auto& r : rs) {
      md << L"- " << FormatDateYYYYMMDD(r.date) << L" ";
      if (!r.time.empty()) md << r.time << L" ";
      md << r.title << L"\n";
      EmitBodyIndented(&md, r.body);
    }
  };

  for (const auto& ck : cats) {
    const std::wstring& cat = ck.first;
    md << L"## 分类: " << cat << L"\n\n";

    auto itp = task_progress.find(cat);
    if (itp != task_progress.end() && !itp->second.empty()) {
      md << L"### 任务进展\n\n";
      for (const auto& tk : itp->second) {
        const TaskBucket& b = tk.second;
        std::wstring title = b.has_task ? b.task.title : (L"(任务: " + tk.first + L")");
        std::wstring st = b.has_task ? std::wstring(StatusToCN(b.task.status)) : L"";
        md << L"#### " << title;
        if (!st.empty()) md << L"  [" << st << L"]";
        md << L"\n\n";

        if (b.has_task) {
          if (!b.task.basis.empty()) md << L"- 依据: " << b.task.basis << L"\n";
          if (!b.task.desc_plain.empty()) {
            md << L"- 说明:\n";
            EmitBodyIndented(&md, b.task.desc_plain);
          }
          if (!b.task.basis.empty() || !b.task.desc_plain.empty()) md << L"\n";
        }

        emit_records(b.records);
        md << L"\n";
      }
    }

    auto itm = meetings.find(cat);
    if (itm != meetings.end() && !itm->second.empty()) {
      md << L"### 会议记录\n";
      emit_records(itm->second);
      md << L"\n";
    }

    auto itn = notes.find(cat);
    if (itn != notes.end() && !itn->second.empty()) {
      md << L"### 其他记录\n";
      emit_records(itn->second);
      md << L"\n";
    }
  }

  *out_md = md.str();
  return true;
}

static std::wstring CsvEscape(const std::wstring& s) {
  bool need_quotes = false;
  for (wchar_t ch : s) {
    if (ch == L',' || ch == L'"' || ch == L'\n' || ch == L'\r') {
      need_quotes = true;
      break;
    }
  }
  std::wstring out;
  if (!need_quotes) return s;
  out.reserve(s.size() + 2);
  out.push_back(L'"');
  for (wchar_t ch : s) {
    if (ch == L'"') out.push_back(L'"');
    out.push_back(ch);
  }
  out.push_back(L'"');
  return out;
}

static std::wstring TypeToStrW(EntryType t) {
  switch (t) {
    case EntryType::TaskProgress: return L"task";
    case EntryType::Meeting: return L"meeting";
    case EntryType::Note:
    default: return L"note";
  }
}

bool GenerateReportCsvFlat(const ReportRange& range, std::wstring* out_csv, std::wstring* err) {
  if (!out_csv) return false;
  out_csv->clear();

  // Load tasks to enrich task entries.
  std::vector<Task> tasks;
  {
    std::wstring terr;
    LoadTasks(&tasks, &terr);
  }
  std::map<std::wstring, Task> task_by_id;
  for (const auto& t : tasks) task_by_id[t.id] = t;

  std::wstringstream csv;
  csv << L"date,entry_type,category,task_id,task_title,task_status,start_time,end_time,title,body_plain,task_basis,task_desc\n";

  SYSTEMTIME cur = range.start;
  while (DateLeq(cur, range.end)) {
    DayData day{};
    std::wstring load_err;
    if (!LoadDayFile(cur, &day, &load_err)) {
      if (err) *err = L"读取失败: " + FormatDateYYYYMMDD(cur) + L"\n" + load_err;
      return false;
    }

    for (const auto& e : day.entries) {
      std::wstring date = FormatDateYYYYMMDD(cur);
      std::wstring etype = TypeToStrW(e.type);
      std::wstring cat = e.category;
      std::wstring task_id;
      std::wstring task_title;
      std::wstring task_status;
      std::wstring task_basis;
      std::wstring task_desc;

      if (e.type == EntryType::TaskProgress && !e.task_id.empty()) {
        task_id = e.task_id;
        auto it = task_by_id.find(e.task_id);
        if (it != task_by_id.end()) {
          const Task& t = it->second;
          if (!t.category.empty()) cat = t.category;
          task_title = t.title;
          task_status = StatusToCN(t.status);
          task_basis = t.basis;
          task_desc = t.desc_plain;
        }
      }

      std::wstring title = e.title;
      std::wstring body = e.body_plain;

      csv << CsvEscape(date) << L","
          << CsvEscape(etype) << L","
          << CsvEscape(cat) << L","
          << CsvEscape(task_id) << L","
          << CsvEscape(task_title) << L","
          << CsvEscape(task_status) << L","
          << CsvEscape(e.start_time) << L","
          << CsvEscape(e.end_time) << L","
          << CsvEscape(title) << L","
          << CsvEscape(body) << L","
          << CsvEscape(task_basis) << L","
          << CsvEscape(task_desc) << L"\n";
    }

    cur = AddDays(cur, 1);
  }

  *out_csv = csv.str();
  return true;
}

static const Task* FindTaskById(const std::vector<Task>& tasks, const std::wstring& id) {
  for (const auto& t : tasks) {
    if (t.id == id) return &t;
  }
  return nullptr;
}

static std::wstring StatusToCNTask(EntryStatus s) {
  switch (s) {
    case EntryStatus::Todo: return L"未开始";
    case EntryStatus::Doing: return L"进行中";
    case EntryStatus::Blocked: return L"阻塞";
    case EntryStatus::Done: return L"已完成";
    case EntryStatus::None:
    default: return L"";
  }
}

bool GenerateTaskProgressMarkdown(const std::wstring& task_id,
                                 const ReportRange& range,
                                 bool include_empty_days,
                                 std::wstring* out_md,
                                 std::wstring* err) {
  if (!out_md) return false;
  out_md->clear();
  if (task_id.empty()) {
    if (err) *err = L"task_id 为空。";
    return false;
  }

  // Load task metadata (best-effort).
  std::vector<Task> tasks;
  {
    std::wstring terr;
    LoadTasks(&tasks, &terr);
  }
  const Task* task = FindTaskById(tasks, task_id);

  std::wstringstream md;
  std::wstring title = task ? task->title : (L"(任务 " + task_id + L")");
  md << L"# 任务进展汇总: " << title << L"\n\n";

  md << L"- task_id: `" << task_id << L"`\n";
  if (task) {
    if (!task->category.empty()) md << L"- 分类: " << task->category << L"\n";
    std::wstring st = StatusToCNTask(task->status);
    if (!st.empty()) md << L"- 状态: " << st << L"\n";
    if (task->start.wYear != 0 && task->end.wYear != 0) {
      md << L"- 任务周期: " << FormatDateYYYYMMDD(task->start) << L" ~ " << FormatDateYYYYMMDD(task->end) << L"\n";
    }
    if (!task->basis.empty()) md << L"- 依据: " << task->basis << L"\n";
    if (!task->desc_plain.empty()) {
      md << L"- 说明:\n";
      EmitBodyIndented(&md, task->desc_plain);
    }
  }
  md << L"\n";

  md << L"范围: " << FormatDateYYYYMMDD(range.start) << L" ~ " << FormatDateYYYYMMDD(range.end) << L"\n\n";

  int days_with_progress = 0;
  int total_entries = 0;

  SYSTEMTIME cur = range.start;
  while (DateLeq(cur, range.end)) {
    DayData day{};
    std::wstring load_err;
    if (!LoadDayFile(cur, &day, &load_err)) {
      if (err) *err = L"读取失败: " + FormatDateYYYYMMDD(cur) + L"\n" + load_err;
      return false;
    }

    struct Row {
      int start_min{-1};
      std::wstring time;
      std::wstring title;
      std::wstring body;
    };
    std::vector<Row> rows;
    rows.reserve(8);
    for (const auto& e : day.entries) {
      if (e.type != EntryType::TaskProgress) continue;
      if (e.task_id != task_id) continue;
      Row r{};
      r.start_min = ParseHHMMMinutes(e.start_time);
      r.time = TimeRange(e);
      r.title = e.title.empty() ? L"(无标题)" : e.title;
      r.body = e.body_plain;
      rows.push_back(std::move(r));
    }

    if (!rows.empty() || include_empty_days) {
      md << L"## " << FormatDateYYYYMMDD(cur) << L"\n";
      if (rows.empty()) {
        md << L"(无记录)\n\n";
      } else {
        days_with_progress++;

        // Bucket by time-of-day for readability.
        enum : int { kLateNight = 0, kMorning = 1, kAfternoon = 2, kEvening = 3, kNoTime = 4, kBuckets = 5 };
        auto bucket_of = [&](const Row& r) -> int {
          if (r.start_min < 0) return kNoTime;
          if (r.start_min < 6 * 60) return kLateNight;      // 00:00-05:59
          if (r.start_min < 12 * 60) return kMorning;       // 06:00-11:59
          if (r.start_min < 18 * 60) return kAfternoon;     // 12:00-17:59
          return kEvening;                                  // 18:00-23:59
        };
        const wchar_t* bucket_title[kBuckets] = {
            L"深夜 (00:00-06:00)",
            L"上午 (06:00-12:00)",
            L"下午 (12:00-18:00)",
            L"晚上 (18:00-24:00)",
            L"未填写时间",
        };

        std::vector<Row> b[kBuckets];
        for (auto& r : rows) {
          b[bucket_of(r)].push_back(std::move(r));
        }
        auto sort_bucket = [&](std::vector<Row>& v) {
          std::sort(v.begin(), v.end(), [](const Row& a, const Row& b) {
            int am = a.start_min < 0 ? 99999 : a.start_min;
            int bm = b.start_min < 0 ? 99999 : b.start_min;
            if (am != bm) return am < bm;
            return a.title < b.title;
          });
        };
        for (int i = 0; i < kBuckets; i++) sort_bucket(b[i]);

        for (int i = 0; i < kBuckets; i++) {
          if (b[i].empty()) continue;
          md << L"### " << bucket_title[i] << L"\n";
          for (const auto& r : b[i]) {
            total_entries++;
            md << L"- ";
            if (!r.time.empty()) md << r.time << L" ";
            md << r.title << L"\n";
            EmitBodyIndented(&md, r.body);
            md << L"\n";
          }
          md << L"\n";
        }
      }
    }

    cur = AddDays(cur, 1);
  }

  md << L"---\n";
  md << L"统计: 有进展日期 " << days_with_progress << L" 天, 进展条目 " << total_entries << L" 条。\n";

  *out_md = md.str();
  return true;
}

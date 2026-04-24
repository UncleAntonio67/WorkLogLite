#include "demo.h"
#include "categories.h"
#include "data_ops.h"
#include "report.h"
#include "storage.h"
#include "tasks.h"
#include "types.h"
#include "win32_util.h"

#include <windows.h>

#include <iostream>
#include <locale>
#include <string>
#include <vector>

namespace {

int g_failures = 0;

void LogLine(const std::string& line) {
  std::cout << line << std::endl;
}

void Check(bool cond, const wchar_t* msg) {
  std::string text = WideToUtf8(msg ? std::wstring(msg) : L"");
  if (cond) {
    LogLine("[PASS] " + text);
  } else {
    LogLine("[FAIL] " + text);
    g_failures++;
  }
}

SYSTEMTIME MakeDate(int y, int m, int d) {
  SYSTEMTIME st{};
  st.wYear = (WORD)y;
  st.wMonth = (WORD)m;
  st.wDay = (WORD)d;
  return st;
}

std::wstring TempTestRoot() {
  wchar_t temp[MAX_PATH]{};
  DWORD n = GetTempPathW(MAX_PATH, temp);
  std::wstring root = (n > 0 && n < MAX_PATH) ? std::wstring(temp) : L".\\";
  root = JoinPath(root, L"WorkLogLiteSmoke");
  wchar_t suffix[64]{};
  wsprintfW(suffix, L"%lu", GetCurrentProcessId());
  root = JoinPath(root, suffix);
  return root;
}

std::wstring GetEnvVar(const wchar_t* name) {
  wchar_t buf[4096]{};
  DWORD n = GetEnvironmentVariableW(name, buf, (DWORD)_countof(buf));
  if (n == 0 || n >= _countof(buf)) return L"";
  return std::wstring(buf);
}

struct ScopedDataRoot {
  std::wstring prev;

  explicit ScopedDataRoot(const std::wstring& root) : prev(GetEnvVar(L"WORKLOGLITE_DATA_DIR")) {
    EnsureDirExists(root);
    SetEnvironmentVariableW(L"WORKLOGLITE_DATA_DIR", root.c_str());
  }

  ~ScopedDataRoot() {
    if (prev.empty()) SetEnvironmentVariableW(L"WORKLOGLITE_DATA_DIR", nullptr);
    else SetEnvironmentVariableW(L"WORKLOGLITE_DATA_DIR", prev.c_str());
  }
};

std::wstring ParentDirOf(const std::wstring& path) {
  size_t p = path.find_last_of(L"\\/");
  if (p == std::wstring::npos) return L".";
  return path.substr(0, p);
}

bool WriteUtf8FileRaw(const std::wstring& path, const std::string& bytes) {
  EnsureDirExists(ParentDirOf(path));
  FILE* f = nullptr;
  _wfopen_s(&f, path.c_str(), L"wb");
  if (!f) return false;
  size_t n = fwrite(bytes.data(), 1, bytes.size(), f);
  fclose(f);
  return n == bytes.size();
}

std::wstring ReadUtf8FileRaw(const std::wstring& path) {
  FILE* f = nullptr;
  _wfopen_s(&f, path.c_str(), L"rb");
  if (!f) return L"";
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
  return Utf8ToWide(bytes);
}

bool DirExistsLoose(const std::wstring& path) {
  DWORD attr = GetFileAttributesW(path.c_str());
  return (attr != INVALID_FILE_ATTRIBUTES) && ((attr & FILE_ATTRIBUTE_DIRECTORY) != 0);
}

void SeedDataSet(const std::wstring& root, const std::wstring& tag) {
  ScopedDataRoot scoped(root);

  std::wstring err;
  std::vector<std::wstring> cats = {L"分类-" + tag, L"共享"};
  Check(SaveCategories(cats, &err), L"SeedDataSet saves categories");

  Task task{};
  task.id = L"task-" + tag;
  task.category = cats[0];
  task.title = L"任务-" + tag;
  task.materials_dir = JoinPath(root, L"materials-" + tag);
  task.basis = L"依据-" + tag;
  task.desc_plain = L"描述-" + tag;
  task.start = MakeDate(2026, 4, 1);
  task.end = MakeDate(2026, 4, 5);
  task.status = EntryStatus::Doing;
  Check(SaveTasks({task}, &err), L"SeedDataSet saves tasks");

  DayData day{};
  day.date = MakeDate(2026, 4, 21);
  Entry entry{};
  entry.id = L"entry-" + tag;
  entry.type = EntryType::Note;
  entry.category = cats[0];
  entry.start_time = L"2026-04-21";
  entry.end_time = L"2026-04-22";
  entry.title = L"标题-" + tag;
  entry.body_plain = L"正文-" + tag;
  day.entries.push_back(entry);
  Check(SaveDayFile(day, &err), L"SeedDataSet saves day data");

  Check(WriteUtf8FileRaw(JoinPath(root, L"quick_replies.txt"), WideToUtf8(L"快捷-" + tag + L"\n")),
        L"SeedDataSet writes auxiliary file");
  Check(WriteUtf8FileRaw(JoinPath(root, L"auth.wla"), "WLA2\nenabled=0\niters=0\nsalt_b64=\nhash_b64=\n"),
        L"SeedDataSet writes auth file");
}

void TestCategories() {
  std::vector<std::wstring> cats = {L" 工作 ", L"会议", L"", L"工作", L"  复盘  "};
  NormalizeCategories(&cats);
  Check(cats.size() == 3, L"NormalizeCategories removes blanks and duplicates");

  std::wstring err;
  Check(SaveCategories(cats, &err), L"SaveCategories succeeds");

  std::vector<std::wstring> loaded;
  Check(LoadCategories(&loaded, &err), L"LoadCategories succeeds");
  Check(loaded == cats, L"Categories roundtrip matches");
}

void TestDateUtilities() {
  SYSTEMTIME st{};
  Check(ParseYYYYMMDD(L"2024-02-29", &st), L"ParseYYYYMMDD accepts leap day");
  Check(!ParseYYYYMMDD(L"2023-02-29", &st), L"ParseYYYYMMDD rejects invalid leap day");
  Check(FormatDateYYYYMMDD(AddDays(MakeDate(2026, 1, 31), 1)) == L"2026-02-01",
        L"AddDays crosses month boundary");
}

void TestDayRoundtripAndValidation() {
  DayData day{};
  day.date = MakeDate(2026, 4, 21);

  Entry note{};
  note.id = L"note-1";
  note.type = EntryType::Note;
  note.category = L"工作";
  note.start_time = L"2026-04-21";
  note.end_time = L"2026-04-23";
  note.title = L"跨日事项";
  note.body_plain = L"正文 A";
  day.entries.push_back(note);

  Entry task_prog{};
  task_prog.id = L"task-1";
  task_prog.type = EntryType::TaskProgress;
  task_prog.task_id = L"task-alpha";
  task_prog.category = L"项目";
  task_prog.start_time = L"2026-04-21";
  task_prog.end_time = L"2026-04-22";
  task_prog.title = L"任务进展";
  task_prog.body_plain = L"完成接口联调";
  day.entries.push_back(task_prog);

  std::wstring err;
  Check(SaveDayFile(day, &err), L"SaveDayFile succeeds for valid entries");

  DayData loaded{};
  Check(LoadDayFile(day.date, &loaded, &err), L"LoadDayFile succeeds after roundtrip");
  Check(loaded.entries.size() == 2, L"Day file roundtrip preserves entry count");
  Check(loaded.entries[0].start_time == note.start_time &&
        loaded.entries[0].end_time == note.end_time &&
        loaded.entries[1].task_id == task_prog.task_id,
        L"Day file roundtrip preserves key fields");

  DayData bad{};
  bad.date = MakeDate(2026, 4, 22);
  Entry invalid{};
  invalid.id = L"bad-1";
  invalid.type = EntryType::Note;
  invalid.category = L"工作";
  invalid.start_time = L"2026-04-22";
  invalid.end_time = L"2026-04-22";
  invalid.title = L"非法日期范围";
  invalid.body_plain = L"x";
  bad.entries.push_back(invalid);
  Check(!SaveDayFile(bad, &err), L"SaveDayFile rejects equal start/end date");

  invalid.end_time = L"2026-04-21";
  bad.entries[0] = invalid;
  Check(!SaveDayFile(bad, &err), L"SaveDayFile rejects reversed date range");
}

void TestLegacyLoadCompatibility() {
  SYSTEMTIME date = MakeDate(2026, 4, 19);
  std::wstring root = GetDataRootDir();
  std::wstring rel = JoinPath(JoinPath(root, L"2026"), L"04");
  std::wstring legacy = JoinPath(rel, L"2026-04-19.wlmd");
  std::string body =
      "- [work] 09:30-10:15 老格式条目\n"
      "  第一行\n"
      "  第二行\n";
  Check(WriteUtf8FileRaw(legacy, body), L"Legacy wlmd test fixture written");

  DayData loaded{};
  std::wstring err;
  Check(LoadDayFile(date, &loaded, &err), L"LoadDayFile reads legacy wlmd");
  Check(loaded.entries.size() == 1 &&
        loaded.entries[0].start_time == L"09:30" &&
        loaded.entries[0].end_time == L"10:15",
        L"Legacy wlmd preserves HH:MM time fields");
}

void TestTasksAndReports() {
  Task t{};
  t.id = L"task-alpha";
  t.category = L"项目";
  t.title = L"支付重构";
  t.materials_dir = L"C:\\materials\\task-alpha";
  t.basis = L"会议纪要";
  t.desc_plain = L"按阶段推进";
  t.start = MakeDate(2026, 4, 1);
  t.end = MakeDate(2026, 4, 30);
  t.status = EntryStatus::Doing;

  std::wstring err;
  Check(SaveTasks({t}, &err), L"SaveTasks succeeds for valid task");

  std::vector<Task> tasks;
  Check(LoadTasks(&tasks, &err), L"LoadTasks succeeds after roundtrip");
  Check(tasks.size() == 1 &&
        tasks[0].title == t.title &&
        tasks[0].materials_dir == t.materials_dir,
        L"Task roundtrip preserves key fields");
  Check(TaskActiveOn(tasks[0], MakeDate(2026, 4, 21)), L"TaskActiveOn true inside range");
  Check(!TaskActiveOn(tasks[0], MakeDate(2026, 5, 1)), L"TaskActiveOn false outside range");

  Task bad = t;
  bad.id = L"task-bad";
  bad.end = bad.start;
  Check(!SaveTasks({bad}, &err), L"SaveTasks rejects equal start/end date");

  bad = t;
  bad.id.clear();
  Check(!SaveTasks({bad}, &err), L"SaveTasks rejects missing id");

  bad = t;
  bad.id = L"task-no-title";
  bad.title.clear();
  Check(!SaveTasks({bad}, &err), L"SaveTasks rejects missing title");

  ReportRange r{};
  r.start = MakeDate(2026, 4, 21);
  r.end = MakeDate(2026, 4, 21);

  std::wstring md;
  Check(GenerateReportMarkdown(r, &md, &err), L"GenerateReportMarkdown succeeds");
  Check(md.find(L"跨日事项") != std::wstring::npos, L"Report markdown contains note title");
  Check(md.find(L"2026-04-21 ~ 2026-04-23") != std::wstring::npos, L"Report markdown shows date range");

  std::wstring csv;
  Check(GenerateReportCsvFlat(r, &csv, &err), L"GenerateReportCsvFlat succeeds");
  Check(csv.find(L"2026-04-21") != std::wstring::npos, L"CSV contains date strings");

  std::wstring task_md;
  Check(GenerateTaskProgressMarkdown(L"task-alpha", r, true, &task_md, &err),
        L"GenerateTaskProgressMarkdown succeeds");
  Check(task_md.find(L"任务进展") != std::wstring::npos, L"Task progress markdown contains entry title");
}

void TestDemoGeneration() {
  std::wstring err;
  Check(GenerateDemoData(MakeDate(2026, 4, 21), &err), L"GenerateDemoData succeeds");

  std::vector<std::wstring> cats;
  Check(LoadCategories(&cats, &err), L"LoadCategories succeeds after demo generation");
  Check(!cats.empty(), L"Demo generation leaves categories available");

  std::vector<Task> tasks;
  Check(LoadTasks(&tasks, &err), L"LoadTasks succeeds after demo generation");
  Check(!tasks.empty(), L"Demo generation creates long-term tasks");

  ReportRange r{};
  r.start = MakeDate(2026, 4, 1);
  r.end = MakeDate(2026, 4, 30);
  std::wstring md;
  Check(GenerateReportMarkdown(r, &md, &err), L"GenerateReportMarkdown succeeds after demo generation");
  Check(md.find(L"支付模块重构") != std::wstring::npos || md.find(L"压测") != std::wstring::npos,
        L"Demo generation produces reportable content");
}

void TestDataOps() {
  std::wstring base = JoinPath(TempTestRoot(), L"data_ops");
  EnsureDirExists(base);

  std::wstring export_src = JoinPath(base, L"export_src");
  SeedDataSet(export_src, L"export");

  std::wstring export_parent = JoinPath(base, L"exports");
  std::wstring export_dir;
  std::wstring err;
  Check(ExportWorkLogLiteData(export_src, export_parent, &export_dir, &err),
        L"ExportWorkLogLiteData succeeds");
  Check(!export_dir.empty() && FileExists(JoinPath(export_dir, L"categories.txt")),
        L"Export copies categories");
  Check(FileExists(JoinPath(export_dir, L"tasks.wlt")) &&
        FileExists(JoinPath(export_dir, L"auth.wla")) &&
        FileExists(JoinPath(export_dir, L"quick_replies.txt")),
        L"Export copies auxiliary files");
  Check(ReadUtf8FileRaw(JoinPath(export_dir, L"quick_replies.txt")).find(L"快捷-export") != std::wstring::npos,
        L"Export preserves file contents");

  std::wstring nested_export = JoinPath(export_src, L"nested_export");
  Check(!ExportWorkLogLiteData(export_src, nested_export, nullptr, &err),
        L"Export rejects target inside source data folder");

  std::wstring import_wrapper = JoinPath(base, L"import_bundle");
  std::wstring import_src = JoinPath(import_wrapper, L"data");
  SeedDataSet(import_src, L"import");

  std::wstring resolved;
  Check(ResolveWorkLogLiteImportRoot(import_wrapper, &resolved) && resolved == import_src,
        L"ResolveWorkLogLiteImportRoot accepts wrapper folder");
  Check(ResolveWorkLogLiteImportRoot(import_src, &resolved) && resolved == import_src,
        L"ResolveWorkLogLiteImportRoot accepts direct data root");

  std::wstring import_dst = JoinPath(base, L"import_live");
  SeedDataSet(import_dst, L"old");

  std::wstring backup_dir;
  Check(ImportWorkLogLiteData(import_wrapper, import_dst, &backup_dir, &err),
        L"ImportWorkLogLiteData replaces existing root");
  Check(!backup_dir.empty() && FileExists(JoinPath(backup_dir, L"categories.txt")),
        L"Import keeps a backup of replaced data");
  {
    ScopedDataRoot scoped(import_dst);
    std::vector<std::wstring> cats;
    Check(LoadCategories(&cats, &err), L"Imported categories load successfully");
    Check(!cats.empty() && cats[0].find(L"分类-import") != std::wstring::npos,
          L"Import replaces categories with source data");

    std::vector<Task> tasks;
    Check(LoadTasks(&tasks, &err), L"Imported tasks load successfully");
    Check(!tasks.empty() && tasks[0].title.find(L"任务-import") != std::wstring::npos,
          L"Import replaces tasks with source data");

    DayData day{};
    Check(LoadDayFile(MakeDate(2026, 4, 21), &day, &err), L"Imported day file loads successfully");
    Check(!day.entries.empty() && day.entries[0].title.find(L"标题-import") != std::wstring::npos,
          L"Import replaces day entries with source data");
  }

  std::wstring fresh_dst = JoinPath(base, L"fresh_import_live");
  std::wstring fresh_backup;
  Check(ImportWorkLogLiteData(import_wrapper, fresh_dst, &fresh_backup, &err),
        L"ImportWorkLogLiteData succeeds when destination root does not exist");
  Check(fresh_backup.empty(), L"Import to fresh root does not create a backup folder");
  {
    ScopedDataRoot scoped(fresh_dst);
    std::vector<std::wstring> cats;
    Check(LoadCategories(&cats, &err), L"Fresh imported categories load successfully");
    Check(!cats.empty() && cats[0].find(L"分类-import") != std::wstring::npos,
          L"Fresh import populates destination data");
  }

  Check(!ImportWorkLogLiteData(import_dst, import_dst, nullptr, &err),
        L"Import rejects using the current data folder as source");

  std::wstring clear_root = JoinPath(base, L"clear_live");
  SeedDataSet(clear_root, L"clear");
  std::wstring clear_backup;
  Check(ClearWorkLogLiteData(clear_root, &clear_backup, &err), L"ClearWorkLogLiteData succeeds");
  Check(!clear_backup.empty() && FileExists(JoinPath(clear_backup, L"categories.txt")),
        L"Clear keeps a backup directory");
  Check(DirExistsLoose(clear_root) && !FileExists(JoinPath(clear_root, L"categories.txt")) &&
            !FileExists(JoinPath(clear_root, L"tasks.wlt")),
        L"Clear recreates an empty data root");
  {
    ScopedDataRoot scoped(clear_root);
    Check(SaveCategories({L"重建"}, &err), L"Data root remains writable after clear");
  }
}

}  // namespace

int wmain() {
  LogLine("WorkLogLite smoke tests");

  std::wstring root = TempTestRoot();
  EnsureDirExists(root);
  SetEnvironmentVariableW(L"WORKLOGLITE_DATA_DIR", root.c_str());

  TestDateUtilities();
  TestCategories();
  TestDayRoundtripAndValidation();
  TestLegacyLoadCompatibility();
  TestTasksAndReports();
  TestDemoGeneration();
  TestDataOps();

  if (g_failures == 0) {
    LogLine("All smoke tests passed.");
    return 0;
  }
  std::cout << "Smoke tests failed: " << g_failures << std::endl;
  return 1;
}

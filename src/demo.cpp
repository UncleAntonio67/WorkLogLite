#include "demo.h"

#include "categories.h"
#include "storage.h"
#include "tasks.h"
#include "win32_util.h"

#include <algorithm>
#include <map>
#include <string>
#include <vector>

static bool StartsWith(const std::wstring& s, const std::wstring& prefix) {
  return s.size() >= prefix.size() && s.compare(0, prefix.size(), prefix) == 0;
}

static bool DateLeq(const SYSTEMTIME& a, const SYSTEMTIME& b) {
  if (a.wYear != b.wYear) return a.wYear < b.wYear;
  if (a.wMonth != b.wMonth) return a.wMonth < b.wMonth;
  return a.wDay <= b.wDay;
}

static int DayOfWeekSun0(const SYSTEMTIME& d) {
  SYSTEMTIME tmp = d;
  FILETIME ft{};
  SystemTimeToFileTime(&tmp, &ft);
  FileTimeToSystemTime(&ft, &tmp);
  return (int)tmp.wDayOfWeek;  // 0=Sunday
}

static bool IsWeekend(const SYSTEMTIME& d) {
  int w = DayOfWeekSun0(d);
  return (w == 0 || w == 6);
}

static bool HasEntryIdPrefix(const DayData& day, const std::wstring& prefix) {
  for (const auto& e : day.entries) {
    if (StartsWith(e.id, prefix)) return true;
  }
  return false;
}

static bool DayHasTitleType(const DayData& day, EntryType type, const std::wstring& title) {
  for (const auto& e : day.entries) {
    if (e.type == type && e.title == title) return true;
  }
  return false;
}

static void AddCategoryIfMissing(std::vector<std::wstring>* cats, const std::wstring& c) {
  if (!cats) return;
  for (const auto& x : *cats) {
    if (x == c) return;
  }
  cats->push_back(c);
}

static std::wstring MDJoin(const std::vector<std::wstring>& lines) {
  std::wstring out;
  for (size_t i = 0; i < lines.size(); i++) {
    out += lines[i];
    out += L"\n";
  }
  return out;
}

static Task MakeDemoTask(const std::wstring& id,
                         const std::wstring& cat,
                         const std::wstring& title,
                         const std::wstring& basis,
                         const std::wstring& desc_md,
                         const SYSTEMTIME& start,
                         const SYSTEMTIME& end,
                         EntryStatus st) {
  Task t{};
  t.id = id;
  t.category = cat;
  t.title = title;
  t.basis = basis;
  t.desc_plain = desc_md;
  t.desc_rtf_b64.clear();
  t.start = start;
  t.end = end;
  t.status = st;
  return t;
}

#if 0  // Recurring meetings removed from app; keep demo generator code disabled for reference.
static RecurringMeeting MakeDemoRecurring(const std::wstring& id,
                                         const std::wstring& cat,
                                         const std::wstring& title,
                                         int weekday,
                                         int interval_weeks,
                                         const SYSTEMTIME& start,
                                         const SYSTEMTIME& end,
                                         const std::wstring& start_time,
                                         const std::wstring& end_time,
                                         const std::wstring& template_md) {
  RecurringMeeting m{};
  m.id = id;
  m.category = cat;
  m.title = title;
  m.weekday = weekday;
  m.interval_weeks = interval_weeks;
  m.start = start;
  m.end = end;
  m.start_time = start_time;
  m.end_time = end_time;
  m.template_plain = template_md;
  return m;
}

static bool HasTaskId(const std::vector<Task>& tasks, const std::wstring& id) {
  for (const auto& t : tasks) {
    if (t.id == id) return true;
  }
  return false;
}

static bool HasRecurringId(const std::vector<RecurringMeeting>& ms, const std::wstring& id) {
  for (const auto& m : ms) {
    if (m.id == id) return true;
  }
  return false;
}

static void AddOrUpdateDemoRecurring(std::vector<RecurringMeeting>* ms, const RecurringMeeting& m) {
  if (!ms) return;
  for (auto& x : *ms) {
    if (x.id == m.id) {
      x = m;
      return;
    }
  }
  ms->push_back(m);
}
#endif  // 0

static void AddOrUpdateDemoTask(std::vector<Task>* tasks, const Task& t) {
  if (!tasks) return;
  for (auto& x : *tasks) {
    if (x.id == t.id) {
      x = t;
      return;
    }
  }
  tasks->push_back(t);
}

bool GenerateDemoData(const SYSTEMTIME& anchor, std::wstring* err) {
  // Load current data (best-effort). We only append if missing.
  std::vector<std::wstring> cats;
  {
    std::wstring cerr;
    LoadCategories(&cats, &cerr);
  }

  std::vector<Task> tasks;
  {
    std::wstring terr;
    LoadTasks(&tasks, &terr);
  }

#if 0
  std::vector<RecurringMeeting> rec;
  {
    std::wstring rerr;
    LoadRecurringMeetings(&rec, &rerr);
  }
#endif

  // Categories aligned with realistic daily work.
  for (const auto& c : std::vector<std::wstring>{
           L"例会", L"技术审查", L"项目", L"研发", L"沟通", L"运维", L"学习", L"杂项",
       }) {
    AddCategoryIfMissing(&cats, c);
  }
  NormalizeCategories(&cats);
  {
    std::wstring werr;
    if (!SaveCategories(cats, &werr)) {
      if (err) *err = werr;
      return false;
    }
  }

  // Define a demo time window.
  SYSTEMTIME start = AddDays(anchor, -16);
  SYSTEMTIME end = AddDays(anchor, +14);

  // Demo tasks: long-running, created based on meetings/reviews.
  {
    Task t1 = MakeDemoTask(
        L"demo-task-pay-refactor",
        L"研发",
        L"支付模块重构(一个月)",
        L"依据: 2026-03-06 技术审查会结论 + 需求变更邮件(支付合并结算流程)",
        MDJoin({
            L"目标: 拆分支付适配层, 降低耦合, 支持多渠道扩展。",
            L"",
            L"范围:",
            L"- 抽象支付网关接口",
            L"- 重构回调验签与幂等",
            L"- 补齐单元测试与回归清单",
            L"",
            L"验收标准:",
            L"- 线上核心链路无回归, 关键接口 99p < 200ms",
            L"- 关键路径覆盖率 >= 70%",
            L"",
            L"风险/依赖:",
            L"- 依赖风控侧签名字段统一",
            L"- 需要联调沙箱环境可用",
        }),
        AddDays(anchor, -10),
        AddDays(anchor, +20),
        EntryStatus::Doing);
    AddOrUpdateDemoTask(&tasks, t1);

    Task t2 = MakeDemoTask(
        L"demo-task-loadtest",
        L"项目",
        L"Q2 性能压测与优化",
        L"依据: 2026-03-04 周例会 Action Item + 历史告警(高峰时段超时)",
        MDJoin({
            L"目标: 找出瓶颈并输出优化计划, 完成一次压测闭环。",
            L"",
            L"输出物:",
            L"- 压测报告(场景/指标/结果)",
            L"- 优化清单(优先级/负责人/预计收益)",
            L"- 回归结果与容量建议",
        }),
        AddDays(anchor, -7),
        AddDays(anchor, +21),
        EntryStatus::Todo);
    AddOrUpdateDemoTask(&tasks, t2);

    Task t3 = MakeDemoTask(
        L"demo-task-monitoring",
        L"运维",
        L"上线回归与监控完善",
        L"依据: 2026-03-01 线上故障复盘",
        MDJoin({
            L"目标: 建立可观测性最小集, 避免同类故障再次发生。",
            L"",
            L"内容:",
            L"- 增加关键埋点与错误码统计",
            L"- 增加慢查询与依赖报警",
            L"- 上线回归 checklist",
        }),
        AddDays(anchor, -14),
        AddDays(anchor, +10),
        EntryStatus::Doing);
    AddOrUpdateDemoTask(&tasks, t3);
  }

  {
    std::wstring terr;
    if (!SaveTasks(tasks, &terr)) {
      if (err) *err = terr;
      return false;
    }
  }

  // Recurring meetings (weekly cadence) with templates. (Disabled: feature removed)
#if 0
  {
    RecurringMeeting standup = MakeDemoRecurring(
        L"demo-rec-standup",
        L"例会",
        L"周例会(同步进展)",
        1,  // Monday
        1,
        AddDays(anchor, -60),
        AddDays(anchor, +120),
        L"10:00",
        L"10:30",
        MDJoin({
            L"## 议程",
            L"- 上周遗留问题跟进",
            L"- 本周重点与风险",
            L"",
            L"## 参会人",
            L"- 我 / 后端 / 测试 / 产品",
            L"",
            L"## 结论",
            L"- ",
            L"",
            L"## Action Items",
            L"- [ ] (负责人: 我, 截止: )",
        }));

    RecurringMeeting tech = MakeDemoRecurring(
        L"demo-rec-techreview",
        L"技术审查",
        L"技术审查会",
        3,  // Wednesday
        1,
        AddDays(anchor, -60),
        AddDays(anchor, +120),
        L"16:00",
        L"17:00",
        MDJoin({
            L"## 评审主题",
            L"- ",
            L"",
            L"## 方案要点",
            L"- ",
            L"",
            L"## 风险与边界",
            L"- ",
            L"",
            L"## 结论",
            L"- [ ] 通过 / [ ] 修改后通过 / [ ] 重提",
            L"",
            L"## Action Items",
            L"- [ ] (负责人: , 截止: )",
        }));

    RecurringMeeting retro = MakeDemoRecurring(
        L"demo-rec-retro",
        L"例会",
        L"周复盘(改进点)",
        5,  // Friday
        1,
        AddDays(anchor, -60),
        AddDays(anchor, +120),
        L"17:00",
        L"17:45",
        MDJoin({
            L"## 本周做得好的",
            L"- ",
            L"",
            L"## 需要改进的",
            L"- ",
            L"",
            L"## 下周行动",
            L"- [ ] ",
        }));

    AddOrUpdateDemoRecurring(&rec, standup);
    AddOrUpdateDemoRecurring(&rec, tech);
    AddOrUpdateDemoRecurring(&rec, retro);
  }

  {
    std::wstring rerr;
    if (!SaveRecurringMeetings(rec, &rerr)) {
      if (err) *err = rerr;
      return false;
    }
  }
#endif

  // Add a selection of daily entries (meetings + progress + misc).
  for (SYSTEMTIME d = start; DateLeq(d, end); d = AddDays(d, 1)) {
    if (IsWeekend(d)) continue;

    DayData day{};
    std::wstring lerr;
    if (!LoadDayFile(d, &day, &lerr)) {
      if (err) *err = L"读取失败: " + FormatDateYYYYMMDD(d) + L"\n" + lerr;
      return false;
    }

    // Avoid re-adding demo entries repeatedly.
    std::wstring prefix = L"demo-" + FormatDateYYYYMMDD(d) + L"-";
    if (HasEntryIdPrefix(day, prefix)) continue;

    int w = DayOfWeekSun0(d);

    // Standup on Monday
    if (w == 1 && !DayHasTitleType(day, EntryType::Meeting, L"周例会(同步进展)")) {
      Entry e{};
      e.id = prefix + L"m1";
      e.type = EntryType::Meeting;
      e.category = L"例会";
      e.start_time = L"10:00";
      e.end_time = L"10:30";
      e.title = L"周例会(同步进展)";
      e.body_plain = MDJoin({
          L"## 结论",
          L"- 本周优先级: 支付重构, 先打通验签幂等",
          L"- 性能压测准备: 先补齐场景与数据",
          L"",
          L"## Action Items",
          L"- [ ] 支付重构: 完成回调验签重构 (负责人: 我, 截止: " + FormatDateYYYYMMDD(AddDays(d, 4)) + L")",
          L"- [ ] 压测方案: 输出压测计划 (负责人: 我, 截止: " + FormatDateYYYYMMDD(AddDays(d, 3)) + L")",
      });
      day.entries.push_back(std::move(e));
    }

    // Tech review on Wednesday
    if (w == 3 && !DayHasTitleType(day, EntryType::Meeting, L"技术审查会")) {
      Entry e{};
      e.id = prefix + L"m2";
      e.type = EntryType::Meeting;
      e.category = L"技术审查";
      e.start_time = L"16:00";
      e.end_time = L"17:00";
      e.title = L"技术审查会";
      e.body_plain = MDJoin({
          L"## 评审主题",
          L"- 支付回调验签与幂等改造",
          L"",
          L"## 方案要点",
          L"- 幂等 key: (order_id, channel, callback_id)",
          L"- 统一验签入口, 每个渠道适配器只负责参数映射",
          L"",
          L"## 结论",
          L"- 修改后通过: 补充回滚方案与监控指标",
          L"",
          L"## Action Items",
          L"- [ ] 增加监控: 验签失败率/幂等冲突率 (负责人: 我, 截止: " + FormatDateYYYYMMDD(AddDays(d, 5)) + L")",
      });
      day.entries.push_back(std::move(e));
    }

    // Weekly retro on Friday
    if (w == 5 && !DayHasTitleType(day, EntryType::Meeting, L"周复盘(改进点)")) {
      Entry e{};
      e.id = prefix + L"m3";
      e.type = EntryType::Meeting;
      e.category = L"例会";
      e.start_time = L"17:00";
      e.end_time = L"17:45";
      e.title = L"周复盘(改进点)";
      e.body_plain = MDJoin({
          L"## 本周做得好的",
          L"- 需求澄清提前, 返工减少",
          L"",
          L"## 需要改进的",
          L"- 性能问题暴露太晚, 需要更早压测",
          L"",
          L"## 下周行动",
          L"- [ ] 提前准备压测数据与场景",
      });
      day.entries.push_back(std::move(e));
    }

    // Task progress: write actual progress on some days (so列表里不只是占位).
    bool write_progress = DateLeq(AddDays(anchor, -8), d) && DateLeq(d, anchor);  // last ~8 workdays
    if (write_progress) {
      auto add_prog = [&](const std::wstring& tid, const std::wstring& cat, const std::wstring& title,
                          const std::wstring& body) {
        Entry e{};
        e.id = prefix + L"t-" + tid;
        e.type = EntryType::TaskProgress;
        e.task_id = tid;
        e.category = cat;
        e.title = title;
        e.body_plain = body;
        day.entries.push_back(std::move(e));
      };

      add_prog(L"demo-task-pay-refactor", L"研发", L"支付模块重构(一个月)",
               MDJoin({
                   L"今日进展:",
                   L"- 完成回调验签入口拆分",
                   L"- 补齐了幂等冲突日志",
                   L"",
                   L"明日计划:",
                   L"- 联调渠道A回调参数",
                   L"",
                   L"问题/风险:",
                   L"- 沙箱环境偶发超时, 需要运维协助排查",
               }));

      if (DayOfWeekSun0(d) == 2 || DayOfWeekSun0(d) == 4) {
        add_prog(L"demo-task-loadtest", L"项目", L"Q2 性能压测与优化",
                 MDJoin({
                     L"今日进展:",
                     L"- 补齐压测场景: 登录/下单/支付回调",
                     L"- 输出目标指标: QPS/99p/错误率",
                     L"",
                     L"下一步:",
                     L"- 准备压测数据与压测脚本",
                 }));
      }

      add_prog(L"demo-task-monitoring", L"运维", L"上线回归与监控完善",
               MDJoin({
                   L"今日进展:",
                   L"- 增加验签失败率监控",
                   L"- 增加慢查询告警阈值",
               }));
    }

    // Misc small work records (碎片化).
    {
      Entry e{};
      e.id = prefix + L"n1";
      e.type = EntryType::Note;
      e.category = L"杂项";
      e.title = L"零碎工作记录";
      e.body_plain = MDJoin({
          L"- 回复产品问题: 确认字段含义与边界",
          L"- 处理一个小bug: 参数为空导致崩溃",
          L"- 跟进测试反馈: 用例补充",
      });
      day.entries.push_back(std::move(e));
    }

    // A future goal (so你能看到“未来某日的工作目标”).
    if (d.wYear == AddDays(anchor, 7).wYear &&
        d.wMonth == AddDays(anchor, 7).wMonth &&
        d.wDay == AddDays(anchor, 7).wDay) {
      Entry e{};
      e.id = prefix + L"goal1";
      e.type = EntryType::Note;
      e.category = L"项目";
      e.title = L"下周目标(示例)";
      e.body_plain = MDJoin({
          L"- 完成支付重构第一阶段, 可灰度",
          L"- 输出性能压测报告初稿",
          L"- 监控完善覆盖核心链路",
      });
      day.entries.push_back(std::move(e));
    }

    std::wstring serr;
    if (!SaveDayFile(day, &serr)) {
      if (err) *err = L"保存失败: " + FormatDateYYYYMMDD(d) + L"\n" + serr;
      return false;
    }
  }

  return true;
}

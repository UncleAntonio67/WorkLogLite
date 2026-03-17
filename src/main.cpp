#include "categories.h"
#include "demo.h"
#include "report.h"
#include "storage.h"
#include "tasks.h"
#include "recurring.h"
#include "types.h"
#include "win32_util.h"

#include <commctrl.h>
#include <commdlg.h>
#include <richedit.h>
#include <shellapi.h>
#include <windowsx.h>

#include <string>
#include <vector>
#include <cwctype>
#include <cstring>
#include <sstream>
#include <string_view>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")

static const wchar_t* kAppTitle = L"WorkLogLite";

enum : UINT {
  IDM_REPORT_QUARTER = 1001,
  IDM_REPORT_YEAR = 1002,
  IDM_REPORT_QUARTER_BY_CAT = 1003,
  IDM_REPORT_YEAR_BY_CAT = 1004,
  IDM_EXPORT_QUARTER_CSV = 1005,
  IDM_EXPORT_YEAR_CSV = 1006,
  IDM_OPEN_DATA_DIR = 1010,
  IDM_MANAGE_CATEGORIES = 1011,
  IDM_PREVIEW_ENTRY = 1020,
  IDM_PREVIEW_DAY = 1021,
  IDM_MANAGE_TASKS = 1030,
  IDM_MANAGE_RECURRING_MEETINGS = 1040,
  IDM_GENERATE_DEMO_DATA = 1050,
  IDM_HELP = 1100,

  // Context menu (ListView)
  IDM_CTX_CREATE_TASK_FROM_MEETING = 20001,
  IDM_CTX_INSERT_MEETING_TEMPLATE = 20002,

  // Formatting (RichEdit body)
  IDM_FMT_BOLD = 30001,
  IDM_FMT_ITALIC = 30002,
  IDM_FMT_UNDERLINE = 30003,
  IDM_FMT_BULLET = 30004,
  IDM_FMT_NUMBERING = 30005,
  IDM_FMT_ALIGN_LEFT = 30006,
  IDM_FMT_ALIGN_CENTER = 30007,
  IDM_FMT_ALIGN_RIGHT = 30008,
  IDM_FMT_INDENT_INC = 30009,
  IDM_FMT_INDENT_DEC = 30010,
  IDM_FMT_CLEAR = 30011,
};

enum : int {
  IDC_CAL = 2001,
  IDC_LIST = 2002,
  IDC_CATEGORY = 2003,
  IDC_ST_DAY = 2008,
  IDC_START = 2004,
  IDC_END = 2005,
  IDC_TITLE = 2006,
  IDC_BODY = 2007,
  IDC_BTN_NEW = 2010,
  IDC_BTN_SAVE = 2011,
  IDC_BTN_DEL = 2012,
  IDC_BTN_PREVIEW = 2013,
  IDC_STATUS = 2020,
  IDC_CB_STATUS = 2021,
  IDC_LBL_CATEGORY = 2030,
  IDC_LBL_START = 2031,
  IDC_LBL_END = 2032,
  IDC_LBL_STATUS = 2033,
  IDC_LBL_TITLE = 2034,
  IDC_LBL_BODY = 2035,
};

struct AppState {
  HWND hwnd{};
  HWND cal{};
  HWND st_day{};
  HWND list{};
  HWND lbl_category{};
  HWND lbl_start{};
  HWND lbl_end{};
  HWND lbl_status{};
  HWND lbl_title{};
  HWND lbl_body{};
  HWND cb_category{};
  HWND ed_start{};
  HWND ed_end{};
  HWND cb_status{};
  HWND ed_title{};
  HWND ed_body{};
  HWND btn_new{};
  HWND btn_save{};
  HWND btn_del{};
  HWND btn_preview{};
  HWND status{};
  HFONT font{};

  SYSTEMTIME selected{};
  DayData day{};
  int editing_index{-1};  // -1 => new entry
  bool editor_dirty{false};

  std::vector<std::wstring> categories;
  std::vector<Task> tasks;
  std::vector<RecurringMeeting> recurring_meetings;
};

static bool IsEditorTrulyDirty(AppState* s);

static AppState* GetState(HWND hwnd) {
  return reinterpret_cast<AppState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
}

static void SetState(HWND hwnd, AppState* s) {
  SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(s));
}

static std::wstring GetWindowTextWString(HWND hwnd) {
  int len = GetWindowTextLengthW(hwnd);
  std::wstring s;
  s.resize((size_t)len);
  if (len > 0) GetWindowTextW(hwnd, s.data(), len + 1);
  return s;
}

static void SetEditText(HWND hwnd, const std::wstring& s) {
  SetWindowTextW(hwnd, s.c_str());
}

static void RefreshCategoryCombo(AppState* s) {
  // Preserve current text
  std::wstring cur = GetWindowTextWString(s->cb_category);
  SendMessageW(s->cb_category, CB_RESETCONTENT, 0, 0);
  for (const auto& c : s->categories) {
    SendMessageW(s->cb_category, CB_ADDSTRING, 0, (LPARAM)c.c_str());
  }
  if (!cur.empty()) SetWindowTextW(s->cb_category, cur.c_str());
}

static void PopulateTimeCombo(HWND cb) {
  if (!cb) return;
  SendMessageW(cb, CB_RESETCONTENT, 0, 0);
  SendMessageW(cb, CB_ADDSTRING, 0, (LPARAM)L"");  // empty option
  for (int h = 0; h <= 23; h++) {
    for (int m = 0; m < 60; m += 15) {
      wchar_t buf[16]{};
      wsprintfW(buf, L"%02d:%02d", h, m);
      SendMessageW(cb, CB_ADDSTRING, 0, (LPARAM)buf);
    }
  }
  SendMessageW(cb, CB_SETCURSEL, 0, 0);
}

static void SetEditCue(HWND edit, const wchar_t* cue) {
  if (!edit) return;
  // EM_SETCUEBANNER supported on Vista+. No-op on older systems.
  SendMessageW(edit, EM_SETCUEBANNER, (WPARAM)TRUE, (LPARAM)cue);
}

struct RichStreamOutCtx {
  std::string bytes;
};

static DWORD CALLBACK RichStreamOutCb(DWORD_PTR cookie, LPBYTE buf, LONG cb, LONG* pcb) {
  auto* ctx = reinterpret_cast<RichStreamOutCtx*>(cookie);
  if (!ctx || !buf || cb <= 0) {
    if (pcb) *pcb = 0;
    return 1;
  }
  ctx->bytes.append(reinterpret_cast<const char*>(buf), (size_t)cb);
  if (pcb) *pcb = cb;
  return 0;
}

struct RichStreamInCtx {
  const char* data{};
  size_t size{};
  size_t pos{};
};

static DWORD CALLBACK RichStreamInCb(DWORD_PTR cookie, LPBYTE buf, LONG cb, LONG* pcb) {
  auto* ctx = reinterpret_cast<RichStreamInCtx*>(cookie);
  if (!ctx || !buf || cb <= 0) {
    if (pcb) *pcb = 0;
    return 1;
  }
  size_t remain = (ctx->pos < ctx->size) ? (ctx->size - ctx->pos) : 0;
  size_t n = remain < (size_t)cb ? remain : (size_t)cb;
  if (n > 0) memcpy(buf, ctx->data + ctx->pos, n);
  ctx->pos += n;
  if (pcb) *pcb = (LONG)n;
  return 0;
}

static std::string RichEditGetRtfBytes(HWND rich) {
  RichStreamOutCtx ctx{};
  EDITSTREAM es{};
  es.dwCookie = (DWORD_PTR)&ctx;
  es.pfnCallback = RichStreamOutCb;
  SendMessageW(rich, EM_STREAMOUT, (WPARAM)SF_RTF, (LPARAM)&es);
  return ctx.bytes;
}

static void RichEditSetRtfBytes(HWND rich, const std::string& bytes) {
  RichStreamInCtx ctx{};
  ctx.data = bytes.data();
  ctx.size = bytes.size();
  ctx.pos = 0;
  EDITSTREAM es{};
  es.dwCookie = (DWORD_PTR)&ctx;
  es.pfnCallback = RichStreamInCb;
  SendMessageW(rich, EM_STREAMIN, (WPARAM)SF_RTF, (LPARAM)&es);
}

static void RichToggleCharEffect(HWND rich, DWORD effect_flag) {
  if (!rich) return;
  CHARFORMAT2W cf{};
  cf.cbSize = sizeof(cf);
  SendMessageW(rich, EM_GETCHARFORMAT, (WPARAM)SCF_SELECTION, (LPARAM)&cf);

  bool enabled = (cf.dwEffects & effect_flag) != 0;
  CHARFORMAT2W out{};
  out.cbSize = sizeof(out);
  out.dwMask = effect_flag;
  out.dwEffects = enabled ? 0 : effect_flag;
  SendMessageW(rich, EM_SETCHARFORMAT, (WPARAM)SCF_SELECTION, (LPARAM)&out);
}

static void RichToggleBullet(HWND rich) {
  if (!rich) return;
  PARAFORMAT2 pf{};
  pf.cbSize = sizeof(pf);
  SendMessageW(rich, EM_GETPARAFORMAT, 0, (LPARAM)&pf);
  bool is_bullet = (pf.wNumbering == PFN_BULLET);

  PARAFORMAT2 out{};
  out.cbSize = sizeof(out);
  out.dwMask = PFM_NUMBERING;
  out.wNumbering = is_bullet ? 0 : PFN_BULLET;
  SendMessageW(rich, EM_SETPARAFORMAT, 0, (LPARAM)&out);
}

static void RichToggleNumbering(HWND rich) {
  if (!rich) return;
  PARAFORMAT2 pf{};
  pf.cbSize = sizeof(pf);
  SendMessageW(rich, EM_GETPARAFORMAT, 0, (LPARAM)&pf);
  bool is_numbered = (pf.wNumbering != 0 && pf.wNumbering != PFN_BULLET);

  PARAFORMAT2 out{};
  out.cbSize = sizeof(out);
  out.dwMask = PFM_NUMBERING | PFM_NUMBERINGSTYLE | PFM_NUMBERINGSTART | PFM_NUMBERINGTAB | PFM_STARTINDENT | PFM_OFFSET;
  if (is_numbered) {
    out.wNumbering = 0;
    out.wNumberingStyle = 0;
    out.wNumberingStart = 0;
    out.wNumberingTab = 0;
    out.dxStartIndent = 0;
    out.dxOffset = 0;
  } else {
    // A reasonable default numbering format.
    out.wNumbering = PFN_ARABIC;
    out.wNumberingStyle = PFNS_PERIOD;
    out.wNumberingStart = 1;
    out.wNumberingTab = 360;   // 0.25"
    out.dxStartIndent = 720;   // 0.5"
    out.dxOffset = -360;       // hanging indent (number is left of text)
  }
  SendMessageW(rich, EM_SETPARAFORMAT, 0, (LPARAM)&out);
}

static void RichSetParaAlignment(HWND rich, WORD align) {
  if (!rich) return;
  PARAFORMAT2 out{};
  out.cbSize = sizeof(out);
  out.dwMask = PFM_ALIGNMENT;
  out.wAlignment = align;
  SendMessageW(rich, EM_SETPARAFORMAT, 0, (LPARAM)&out);
}

static void RichAdjustIndent(HWND rich, int delta_twips) {
  if (!rich || delta_twips == 0) return;
  PARAFORMAT2 pf{};
  pf.cbSize = sizeof(pf);
  SendMessageW(rich, EM_GETPARAFORMAT, 0, (LPARAM)&pf);

  bool is_list = (pf.wNumbering != 0);
  int start = (int)pf.dxStartIndent;
  int offset = (int)pf.dxOffset;

  start += delta_twips;
  if (start < 0) start = 0;
  if (is_list) {
    offset += delta_twips;
    // Keep offset reasonable; a hanging indent is typically negative.
    if (offset > 0) offset = 0;
  }

  PARAFORMAT2 out{};
  out.cbSize = sizeof(out);
  out.dwMask = PFM_STARTINDENT | PFM_OFFSET;
  out.dxStartIndent = start;
  out.dxOffset = offset;
  SendMessageW(rich, EM_SETPARAFORMAT, 0, (LPARAM)&out);
}

static void RichClearFormatting(HWND rich) {
  if (!rich) return;

  // Clear common character-level formatting on the selection.
  CHARFORMAT2W cf{};
  cf.cbSize = sizeof(cf);
  cf.dwMask = CFM_BOLD | CFM_ITALIC | CFM_UNDERLINE | CFM_STRIKEOUT | CFM_COLOR;
  cf.dwEffects = CFE_AUTOCOLOR;  // also implies "not bold/italic/underline/strikeout"
  SendMessageW(rich, EM_SETCHARFORMAT, (WPARAM)SCF_SELECTION, (LPARAM)&cf);

  // Clear paragraph-level formatting on the selection.
  PARAFORMAT2 pf{};
  pf.cbSize = sizeof(pf);
  pf.dwMask = PFM_NUMBERING | PFM_ALIGNMENT | PFM_STARTINDENT | PFM_OFFSET | PFM_SPACEBEFORE | PFM_SPACEAFTER;
  pf.wNumbering = 0;
  pf.wAlignment = PFA_LEFT;
  pf.dxStartIndent = 0;
  pf.dxOffset = 0;
  pf.dySpaceBefore = 0;
  pf.dySpaceAfter = 0;
  SendMessageW(rich, EM_SETPARAFORMAT, 0, (LPARAM)&pf);
}

static HWND GetComboEditHandle(HWND combo) {
  COMBOBOXINFO cbi{};
  cbi.cbSize = sizeof(cbi);
  if (!GetComboBoxInfo(combo, &cbi)) return nullptr;
  return cbi.hwndItem;  // the edit control for CBS_DROPDOWN
}

static void UpdateDayHeader(AppState* s) {
  if (!s->st_day) return;
  std::wstring date = FormatDateYYYYMMDD(s->selected);
  std::wstring text = date + L"    条目: ";
  wchar_t buf[32]{};
  wsprintfW(buf, L"%d", (int)s->day.entries.size());
  text += buf;
  SetWindowTextW(s->st_day, text.c_str());
}

static void UpdateStatusBar(AppState* s) {
  if (!s->status) return;
  // Part 1 is a static hint; part 0 is dynamic info.
  std::wstring hint = L"快捷键: Ctrl+S 保存  Ctrl+N 新增  Ctrl+P 预览  Ctrl+K 分类  Ctrl+T 长期任务  Ctrl+R 周期会议  F1 帮助";
  SendMessageW(s->status, SB_SETTEXTW, 1, (LPARAM)hint.c_str());

  std::wstring info = FormatDateYYYYMMDD(s->selected);
  info += L"  ";
  if (s->editing_index >= 0) {
    wchar_t buf[32]{};
    wsprintfW(buf, L"#%d", s->editing_index + 1);
    info += buf;
  } else {
    info += L"(新建)";
  }
  if (IsEditorTrulyDirty(s)) info += L"  未保存";
  SendMessageW(s->status, SB_SETTEXTW, 0, (LPARAM)info.c_str());
}

static void UpdateStatusParts(AppState* s) {
  if (!s || !s->status) return;
  RECT rc{};
  GetClientRect(s->hwnd, &rc);
  int w = rc.right - rc.left;
  int right_hint = w - ScalePx(s->hwnd, 520);
  if (right_hint < ScalePx(s->hwnd, 220)) right_hint = w / 2;
  int parts[2]{right_hint, -1};
  SendMessageW(s->status, SB_SETPARTS, 2, (LPARAM)parts);
}

static const wchar_t* StatusToCN(EntryStatus st) {
  switch (st) {
    case EntryStatus::Todo: return L"未开始";
    case EntryStatus::Doing: return L"进行中";
    case EntryStatus::Blocked: return L"阻塞";
    case EntryStatus::Done: return L"已完成";
    case EntryStatus::None:
    default: return L"无";
  }
}

static EntryStatus StatusFromComboSel(int idx) {
  switch (idx) {
    case 1: return EntryStatus::Todo;
    case 2: return EntryStatus::Doing;
    case 3: return EntryStatus::Blocked;
    case 4: return EntryStatus::Done;
    default: return EntryStatus::None;
  }
}

static int ComboSelFromStatus(EntryStatus st) {
  switch (st) {
    case EntryStatus::Todo: return 1;
    case EntryStatus::Doing: return 2;
    case EntryStatus::Blocked: return 3;
    case EntryStatus::Done: return 4;
    case EntryStatus::None:
    default: return 0;
  }
}

static Task* FindTaskById(AppState* s, const std::wstring& id) {
  if (!s) return nullptr;
  for (auto& t : s->tasks) {
    if (t.id == id) return &t;
  }
  return nullptr;
}

static EntryStatus EffectiveStatus(AppState* s, const Entry& e) {
  if (e.status != EntryStatus::None) return e.status;
  if (e.type == EntryType::TaskProgress && !e.task_id.empty()) {
    Task* t = FindTaskById(s, e.task_id);
    if (t) return t->status;
  }
  return EntryStatus::None;
}

static bool EntryIsTaskProgress(const Entry& e) {
  return e.type == EntryType::TaskProgress && !e.task_id.empty();
}

static COLORREF StatusBkColor(EntryStatus st) {
  switch (st) {
    case EntryStatus::Todo: return RGB(245, 245, 245);
    case EntryStatus::Doing: return RGB(225, 240, 255);
    case EntryStatus::Blocked: return RGB(255, 230, 230);
    case EntryStatus::Done: return RGB(225, 250, 235);
    case EntryStatus::None:
    default: return CLR_NONE;
  }
}

static void SetEditorDirty(AppState* s, bool dirty) {
  s->editor_dirty = dirty;
  // Simple affordance: append "*" to title bar when dirty.
  std::wstring title = kAppTitle;
  if (dirty) title += L" *";
  SetWindowTextW(s->hwnd, title.c_str());
}

static bool EditorHasMeaningfulContent(AppState* s) {
  std::wstring category = GetWindowTextWString(s->cb_category);
  std::wstring start_time = GetWindowTextWString(s->ed_start);
  std::wstring end_time = GetWindowTextWString(s->ed_end);
  std::wstring title = GetWindowTextWString(s->ed_title);
  std::wstring body = GetWindowTextWString(s->ed_body);

  // If user hasn't typed anything besides category, don't nag on navigation.
  bool has_any = !title.empty() || !body.empty() || !start_time.empty() || !end_time.empty();
  (void)category;
  return has_any;
}

static bool EditorMatchesEntry(AppState* s, const Entry& e) {
  Entry cur{};
  cur.category = GetWindowTextWString(s->cb_category);
  cur.start_time = GetWindowTextWString(s->ed_start);
  cur.end_time = GetWindowTextWString(s->ed_end);
  cur.title = GetWindowTextWString(s->ed_title);
  cur.body_plain = GetWindowTextWString(s->ed_body);
  return cur.category == e.category && cur.start_time == e.start_time && cur.end_time == e.end_time &&
         cur.title == e.title && cur.body_plain == e.body_plain;
}

static bool IsEditorTrulyDirty(AppState* s) {
  // RichEdit can be modified by formatting changes without changing plain text.
  bool body_modified = false;
  if (s->ed_body) body_modified = (SendMessageW(s->ed_body, EM_GETMODIFY, 0, 0) != 0);

  if (!s->editor_dirty && !body_modified) return false;
  if (s->editing_index >= 0 && s->editing_index < (int)s->day.entries.size()) {
    if (body_modified) return true;
    return !EditorMatchesEntry(s, s->day.entries[(size_t)s->editing_index]);
  }
  // For new entries, avoid nagging on empty content even if RichEdit modify flag was toggled.
  return EditorHasMeaningfulContent(s);
}

static void ClearEditor(AppState* s) {
  if (!s->categories.empty()) SetWindowTextW(s->cb_category, s->categories[0].c_str());
  else SetWindowTextW(s->cb_category, L"");
  SetEditText(s->ed_start, L"");
  SetEditText(s->ed_end, L"");
  if (s->cb_status) SendMessageW(s->cb_status, CB_SETCURSEL, 0, 0);
  SetEditText(s->ed_title, L"");
  SetEditText(s->ed_body, L"");
  if (s->ed_title) SendMessageW(s->ed_title, EM_SETMODIFY, FALSE, 0);
  if (s->ed_body) SendMessageW(s->ed_body, EM_SETMODIFY, FALSE, 0);
  s->editing_index = -1;
  SetEditorDirty(s, false);
  EnableWindow(s->cb_category, TRUE);
  EnableWindow(s->ed_title, TRUE);
  if (s->cb_status) EnableWindow(s->cb_status, FALSE);
  UpdateStatusBar(s);
}

static void FillEditorFromEntry(AppState* s, const Entry& e, int index) {
  bool is_task = EntryIsTaskProgress(e);
  SetWindowTextW(s->cb_category, e.category.c_str());
  SetEditText(s->ed_start, e.start_time);
  SetEditText(s->ed_end, e.end_time);
  if (s->cb_status) {
    SendMessageW(s->cb_status, CB_SETCURSEL, ComboSelFromStatus(EffectiveStatus(s, e)), 0);
    EnableWindow(s->cb_status, is_task ? TRUE : FALSE);
  }
  SetEditText(s->ed_title, e.title);
  if (!e.body_rtf_b64.empty()) {
    std::string rtf;
    if (Base64Decode(e.body_rtf_b64, &rtf)) {
      RichEditSetRtfBytes(s->ed_body, rtf);
    } else {
      SetEditText(s->ed_body, e.body_plain);
    }
  } else {
    SetEditText(s->ed_body, e.body_plain);
  }
  if (s->ed_title) SendMessageW(s->ed_title, EM_SETMODIFY, FALSE, 0);
  if (s->ed_body) SendMessageW(s->ed_body, EM_SETMODIFY, FALSE, 0);
  s->editing_index = index;
  SetEditorDirty(s, false);
  EnableWindow(s->cb_category, is_task ? FALSE : TRUE);
  EnableWindow(s->ed_title, is_task ? FALSE : TRUE);
  UpdateStatusBar(s);
}

static std::wstring TimeRangeText(const Entry& e) {
  if (e.start_time.empty() && e.end_time.empty()) return L"";
  std::wstring st = e.start_time.empty() ? L"??:??" : e.start_time;
  std::wstring et = e.end_time.empty() ? L"??:??" : e.end_time;
  return st + L"-" + et;
}

static void RefreshList(AppState* s) {
  ListView_DeleteAllItems(s->list);

  for (int i = 0; i < (int)s->day.entries.size(); i++) {
    const auto& e = s->day.entries[(size_t)i];

    LVITEMW it{};
    it.mask = LVIF_TEXT;
    it.iItem = i;
    std::wstring t0 = TimeRangeText(e);
    it.pszText = const_cast<wchar_t*>(t0.c_str());
    ListView_InsertItem(s->list, (LVITEM*)&it);

    ListView_SetItemText(s->list, i, 1, const_cast<wchar_t*>(e.category.c_str()));
    ListView_SetItemText(s->list, i, 2, const_cast<wchar_t*>(e.title.c_str()));
    EntryStatus st = EffectiveStatus(s, e);
    ListView_SetItemText(s->list, i, 3, const_cast<wchar_t*>(StatusToCN(st)));
  }

  UpdateDayHeader(s);
}

static bool LoadSelectedDay(AppState* s, const SYSTEMTIME& date) {
  DayData dd{};
  std::wstring err;
  if (!LoadDayFile(date, &dd, &err)) {
    ShowInfoBox(s->hwnd, err.c_str(), L"读取失败");
    return false;
  }
  s->selected = date;
  s->day = dd;

  // Inject recurring meetings as placeholders (not persisted unless saved).
  if (!s->recurring_meetings.empty()) {
    std::vector<std::wstring> existing_titles;
    existing_titles.reserve(s->day.entries.size());
    for (const auto& e : s->day.entries) {
      if (e.type == EntryType::Meeting) existing_titles.push_back(e.title);
    }
    for (const auto& rm : s->recurring_meetings) {
      if (!RecurringMeetingOccursOn(rm, date)) continue;
      bool found = false;
      for (const auto& t : existing_titles) {
        if (t == rm.title) { found = true; break; }
      }
      if (found) continue;

      Entry e{};
      e.id = L"rec-" + rm.id;
      e.placeholder = true;
      e.type = EntryType::Meeting;
      e.category = rm.category.empty() ? L"会议" : rm.category;
      e.title = rm.title;
      e.start_time = rm.start_time;
      e.end_time = rm.end_time;
      // Pre-fill template for easier note taking. It is still a placeholder and will not persist unless user saves.
      if (!rm.template_plain.empty()) e.body_plain = rm.template_plain;
      s->day.entries.push_back(std::move(e));
    }
  }

  // Inject active long-term tasks as placeholder entries (not persisted unless edited/saved).
  if (!s->tasks.empty()) {
    std::vector<std::wstring> existing;
    existing.reserve(s->day.entries.size());
    for (const auto& e : s->day.entries) {
      if (EntryIsTaskProgress(e)) existing.push_back(e.task_id);
    }
    for (const auto& t : s->tasks) {
      if (!TaskActiveOn(t, date)) continue;
      bool found = false;
      for (const auto& tid : existing) {
        if (tid == t.id) { found = true; break; }
      }
      if (found) continue;

      Entry e{};
      e.id = L"task-" + t.id;
      e.type = EntryType::TaskProgress;
      e.placeholder = true;
      e.task_id = t.id;
      e.category = t.category;
      e.title = t.title;
      e.status = EntryStatus::None;  // use task's status as effective status
      s->day.entries.push_back(std::move(e));
    }
  }

  // Merge categories from loaded entries into in-memory list (not persisted until save/manage).
  for (const auto& e : s->day.entries) {
    if (!e.category.empty() && !CategoriesContains(s->categories, e.category)) {
      s->categories.push_back(e.category);
    }
  }
  NormalizeCategories(&s->categories);
  RefreshCategoryCombo(s);

  RefreshList(s);
  ClearEditor(s);
  UpdateDayHeader(s);
  UpdateStatusBar(s);
  return true;
}

static bool ValidateTimeField(const std::wstring& t) {
  if (t.empty()) return true;
  if (t.size() != 5) return false;
  if (t[2] != L':') return false;
  if (t[0] < L'0' || t[0] > L'9') return false;
  if (t[1] < L'0' || t[1] > L'9') return false;
  if (t[3] < L'0' || t[3] > L'9') return false;
  if (t[4] < L'0' || t[4] > L'9') return false;
  int hh = (t[0] - L'0') * 10 + (t[1] - L'0');
  int mm = (t[3] - L'0') * 10 + (t[4] - L'0');
  return (hh >= 0 && hh <= 23 && mm >= 0 && mm <= 59);
}

static bool SaveEditorToModel(AppState* s, bool* out_changed) {
  if (out_changed) *out_changed = false;
  Entry e{};
  if (s->editing_index >= 0 && s->editing_index < (int)s->day.entries.size()) {
    e = s->day.entries[(size_t)s->editing_index];
  }
  e.category = GetWindowTextWString(s->cb_category);
  e.start_time = GetWindowTextWString(s->ed_start);
  e.end_time = GetWindowTextWString(s->ed_end);
  e.title = GetWindowTextWString(s->ed_title);
  e.body_plain = GetWindowTextWString(s->ed_body);
  // Persist rich text (RTF) if available (RichEdit control).
  if (s->ed_body) {
    // NOTE: RichEdit streams out a non-empty RTF header even for empty content.
    // Only persist RTF when there is real text content.
    if (!e.body_plain.empty()) {
      std::string rtf = RichEditGetRtfBytes(s->ed_body);
      if (!rtf.empty()) e.body_rtf_b64 = Base64Encode(rtf);
    } else {
      e.body_rtf_b64.clear();
    }
  }

  // Status (only meaningful for long-term tasks). We treat it as a property of the Task itself.
  bool task_status_changed = false;
  if (EntryIsTaskProgress(e) && s->cb_status) {
    int idx = (int)SendMessageW(s->cb_status, CB_GETCURSEL, 0, 0);
    EntryStatus chosen = StatusFromComboSel(idx);
    Task* t = FindTaskById(s, e.task_id);
    if (t && t->status != chosen) {
      t->status = chosen;
      task_status_changed = true;
    }
    // Keep per-day entry status empty; list uses EffectiveStatus(task.status) when entry.status == None.
    e.status = EntryStatus::None;
  } else {
    e.status = EntryStatus::None;
  }

  // Trim category
  while (!e.category.empty() && iswspace(e.category.front())) e.category.erase(e.category.begin());
  while (!e.category.empty() && iswspace(e.category.back())) e.category.pop_back();
  if (e.category.empty()) {
    ShowInfoBox(s->hwnd, L"请填写分类，例如：工作 / 会议 / 目标 / 沟通 / 研发。", L"分类必填");
    HWND h = GetComboEditHandle(s->cb_category);
    if (h) {
      SetFocus(h);
      SendMessageW(h, EM_SETSEL, 0, -1);
    } else {
      SetFocus(s->cb_category);
    }
    return false;
  }

  if (!ValidateTimeField(e.start_time) || !ValidateTimeField(e.end_time)) {
    ShowInfoBox(s->hwnd, L"时间格式建议为 HH:MM，例如 09:30。也可以留空。", L"时间格式");
    SetFocus(!ValidateTimeField(e.start_time) ? s->ed_start : s->ed_end);
    return false;
  }

  bool has_any = false;
  if (EntryIsTaskProgress(e)) {
    has_any = !e.body_plain.empty() || !e.start_time.empty() || !e.end_time.empty();
  } else {
    has_any = !e.title.empty() || !e.body_plain.empty() || !e.start_time.empty() || !e.end_time.empty();
  }
  if (!has_any) {
    // Treat as "nothing to save": allow navigation without blocking.
    if (out_changed) *out_changed = task_status_changed;
    if (task_status_changed) {
      std::wstring terr;
      SaveTasks(s->tasks, &terr);
      RefreshList(s);
    }
    return true;
  }

  if (s->editing_index >= 0 && s->editing_index < (int)s->day.entries.size()) {
    s->day.entries[(size_t)s->editing_index] = e;
  } else {
    // Give it a stable id for persistence.
    if (e.id.empty()) {
      SYSTEMTIME st{};
      GetLocalTime(&st);
      wchar_t buf[64]{};
      wsprintfW(buf, L"n%04d%02d%02d-%02d%02d%02d",
                (int)st.wYear, (int)st.wMonth, (int)st.wDay, (int)st.wHour, (int)st.wMinute, (int)st.wSecond);
      e.id = buf;
    }
    s->day.entries.push_back(e);
    s->editing_index = (int)s->day.entries.size() - 1;
  }
  if (out_changed) *out_changed = true;

  // Once user saves, this entry should persist.
  s->day.entries[(size_t)s->editing_index].placeholder = false;

  // Add category to list if new, and persist categories best-effort.
  if (!CategoriesContains(s->categories, e.category)) {
    s->categories.push_back(e.category);
    NormalizeCategories(&s->categories);
    RefreshCategoryCombo(s);
    std::wstring cat_err;
    SaveCategories(s->categories, &cat_err);
  }

  if (task_status_changed) {
    std::wstring terr;
    SaveTasks(s->tasks, &terr);
    RefreshList(s);
  }
  return true;
}

static bool PersistDay(AppState* s) {
  std::wstring err;
  if (!SaveDayFile(s->day, &err)) {
    ShowInfoBox(s->hwnd, err.c_str(), L"保存失败");
    return false;
  }
  return true;
}

static bool SaveCurrent(AppState* s) {
  bool changed = false;
  if (!SaveEditorToModel(s, &changed)) return false;
  if (!changed) {
    SetEditorDirty(s, false);
    if (s->ed_title) SendMessageW(s->ed_title, EM_SETMODIFY, FALSE, 0);
    if (s->ed_body) SendMessageW(s->ed_body, EM_SETMODIFY, FALSE, 0);
    UpdateStatusBar(s);
    return true;
  }
  if (!PersistDay(s)) return false;
  RefreshList(s);
  if (s->editing_index >= 0) {
    ListView_SetItemState(s->list, s->editing_index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_EnsureVisible(s->list, s->editing_index, FALSE);
  }
  SetEditorDirty(s, false);
  if (s->ed_title) SendMessageW(s->ed_title, EM_SETMODIFY, FALSE, 0);
  if (s->ed_body) SendMessageW(s->ed_body, EM_SETMODIFY, FALSE, 0);
  UpdateStatusBar(s);
  return true;
}

static void DeleteCurrent(AppState* s) {
  if (s->editing_index < 0 || s->editing_index >= (int)s->day.entries.size()) return;
  int idx = s->editing_index;
  s->day.entries.erase(s->day.entries.begin() + idx);

  // Save after delete.
  std::wstring err;
  if (!s->day.entries.empty()) {
    if (!SaveDayFile(s->day, &err)) {
      ShowInfoBox(s->hwnd, err.c_str(), L"删除失败");
      return;
    }
  } else {
    std::wstring path = GetDayFilePath(s->selected);
    DeleteFileW(path.c_str());
  }

  RefreshList(s);
  ClearEditor(s);
}

static bool PromptSaveIfDirty(AppState* s) {
  if (!IsEditorTrulyDirty(s)) return true;
  int r = MessageBoxW(s->hwnd, L"当前编辑尚未保存。是否保存？", kAppTitle, MB_YESNOCANCEL | MB_ICONWARNING);
  if (r == IDCANCEL) return false;
  if (r == IDNO) {
    SetEditorDirty(s, false);
    return true;
  }
  return SaveCurrent(s);
}

static bool WriteUtf8File(const std::wstring& path, const std::wstring& content, std::wstring* err) {
  std::string bytes = WideToUtf8(content);
  std::wstring tmp = path + L".tmp";
  {
    FILE* f = nullptr;
    _wfopen_s(&f, tmp.c_str(), L"wb");
    if (!f) {
      if (err) *err = L"无法写入文件: " + tmp;
      return false;
    }
    fwrite(bytes.data(), 1, bytes.size(), f);
    fclose(f);
  }
  if (!MoveFileExW(tmp.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
    if (err) *err = L"写入失败: " + path;
    DeleteFileW(tmp.c_str());
    return false;
  }
  return true;
}

static bool WriteUtf8FileWithBom(const std::wstring& path, const std::wstring& content, std::wstring* err) {
  std::string bytes = WideToUtf8(content);
  const unsigned char bom[3] = {0xEF, 0xBB, 0xBF};
  std::wstring tmp = path + L".tmp";
  {
    FILE* f = nullptr;
    _wfopen_s(&f, tmp.c_str(), L"wb");
    if (!f) {
      if (err) *err = L"无法写入文件: " + tmp;
      return false;
    }
    fwrite(bom, 1, 3, f);
    fwrite(bytes.data(), 1, bytes.size(), f);
    fclose(f);
  }
  if (!MoveFileExW(tmp.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
    if (err) *err = L"写入失败: " + path;
    DeleteFileW(tmp.c_str());
    return false;
  }
  return true;
}

static bool CopyToClipboard(HWND hwnd, const std::wstring& text) {
  if (!OpenClipboard(hwnd)) return false;
  EmptyClipboard();
  size_t bytes = (text.size() + 1) * sizeof(wchar_t);
  HGLOBAL h = GlobalAlloc(GMEM_MOVEABLE, bytes);
  if (!h) {
    CloseClipboard();
    return false;
  }
  void* p = GlobalLock(h);
  memcpy(p, text.c_str(), bytes);
  GlobalUnlock(h);
  SetClipboardData(CF_UNICODETEXT, h);
  CloseClipboard();
  return true;
}

static Entry GetEditorEntrySnapshot(AppState* s) {
  Entry e{};
  e.category = GetWindowTextWString(s->cb_category);
  e.start_time = GetWindowTextWString(s->ed_start);
  e.end_time = GetWindowTextWString(s->ed_end);
  e.title = GetWindowTextWString(s->ed_title);
  e.body_plain = GetWindowTextWString(s->ed_body);

  // Trim category
  while (!e.category.empty() && iswspace(e.category.front())) e.category.erase(e.category.begin());
  while (!e.category.empty() && iswspace(e.category.back())) e.category.pop_back();
  return e;
}

static std::wstring EntryToMarkdownBlock(const Entry& e) {
  std::wstring md;
  md += L"- [";
  md += e.category.empty() ? L"(未分类)" : e.category;
  md += L"] ";
  if (!e.start_time.empty() || !e.end_time.empty()) {
    std::wstring st = e.start_time.empty() ? L"??:??" : e.start_time;
    std::wstring et = e.end_time.empty() ? L"??:??" : e.end_time;
    md += st + L"-" + et + L" ";
  }
  md += e.title.empty() ? L"(无标题)" : e.title;
  md += L"\n";
  if (!e.body_plain.empty()) {
    std::wstringstream bss(e.body_plain);
    std::wstring line;
    while (std::getline(bss, line)) {
      if (!line.empty() && line.back() == L'\r') line.pop_back();
      md += L"  ";
      md += line;
      md += L"\n";
    }
  }
  return md;
}

static std::wstring BuildMarkdownForPreviewEntry(AppState* s) {
  Entry e = GetEditorEntrySnapshot(s);
  std::wstring md;
  md += L"# ";
  md += FormatDateYYYYMMDD(s->selected);
  md += L"\n\n";
  md += EntryToMarkdownBlock(e);
  return md;
}

static std::wstring BuildMarkdownForPreviewDay(AppState* s) {
  std::wstring md;
  md += L"# ";
  md += FormatDateYYYYMMDD(s->selected);
  md += L"\n\n";
  if (s->day.entries.empty()) {
    md += L"(当天没有记录)\n";
    return md;
  }
  for (const auto& e : s->day.entries) {
    md += EntryToMarkdownBlock(e);
    md += L"\n";
  }
  return md;
}

static bool SaveFileDialog(HWND owner, const std::wstring& default_name, std::wstring* out_path) {
  wchar_t buf[MAX_PATH]{};
  wcsncpy_s(buf, _countof(buf), default_name.c_str(), _TRUNCATE);

  OPENFILENAMEW ofn{};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = L"Markdown (*.md)\0*.md\0All Files (*.*)\0*.*\0";
  ofn.lpstrFile = buf;
  ofn.nMaxFile = MAX_PATH;
  ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;
  ofn.lpstrDefExt = L"md";
  if (!GetSaveFileNameW(&ofn)) return false;
  if (out_path) *out_path = std::wstring(buf);
  return true;
}

static bool SaveCsvFileDialog(HWND owner, const std::wstring& default_name, std::wstring* out_path) {
  wchar_t buf[MAX_PATH]{};
  wcsncpy_s(buf, _countof(buf), default_name.c_str(), _TRUNCATE);

  OPENFILENAMEW ofn{};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = L"CSV (*.csv)\0*.csv\0All Files (*.*)\0*.*\0";
  ofn.lpstrFile = buf;
  ofn.nMaxFile = MAX_PATH;
  ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;
  ofn.lpstrDefExt = L"csv";
  if (!GetSaveFileNameW(&ofn)) return false;
  if (out_path) *out_path = std::wstring(buf);
  return true;
}

struct ReportWindowState {
  HWND hwnd{};
  HWND edit{};
  HWND btn_copy{};
  HWND btn_export{};
  HWND btn_close{};
  HFONT font{};
  std::wstring markdown;
  std::wstring default_name;
};

static LRESULT CALLBACK ReportWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
  auto* s = reinterpret_cast<ReportWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
    case WM_CREATE: {
      auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
      s = reinterpret_cast<ReportWindowState*>(cs->lpCreateParams);
      s->hwnd = hwnd;
      SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(s));

      s->font = CreateUiFont(hwnd);

      s->edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY |
                                    WS_VSCROLL | WS_TABSTOP,
                                0, 0, 10, 10, hwnd, (HMENU)1, nullptr, nullptr);
      SetEditText(s->edit, s->markdown);

      s->btn_copy = CreateWindowExW(0, L"BUTTON", L"复制",
                                    WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                    0, 0, 10, 10, hwnd, (HMENU)2, nullptr, nullptr);
      s->btn_export = CreateWindowExW(0, L"BUTTON", L"导出",
                                      WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                      0, 0, 10, 10, hwnd, (HMENU)3, nullptr, nullptr);
      s->btn_close = CreateWindowExW(0, L"BUTTON", L"关闭",
                                     WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                     0, 0, 10, 10, hwnd, (HMENU)4, nullptr, nullptr);

      SetControlFont(s->edit, s->font);
      SetControlFont(s->btn_copy, s->font);
      SetControlFont(s->btn_export, s->font);
      SetControlFont(s->btn_close, s->font);
      return 0;
    }
    case WM_SIZE: {
      int w = LOWORD(lParam);
      int h = HIWORD(lParam);
      int pad = ScalePx(hwnd, 10);
      int btn_h = ScalePx(hwnd, 30);
      int btn_w = ScalePx(hwnd, 90);
      int y_btn = h - pad - btn_h;
      MoveWindow(s->edit, pad, pad, w - 2 * pad, y_btn - 2 * pad, TRUE);

      int x = w - pad - btn_w;
      MoveWindow(s->btn_close, x, y_btn, btn_w, btn_h, TRUE);
      x -= pad + btn_w;
      MoveWindow(s->btn_export, x, y_btn, btn_w, btn_h, TRUE);
      x -= pad + btn_w;
      MoveWindow(s->btn_copy, x, y_btn, btn_w, btn_h, TRUE);
      return 0;
    }
    case WM_COMMAND: {
      int id = LOWORD(wParam);
      if (id == 2) {
        if (!CopyToClipboard(hwnd, s->markdown)) ShowInfoBox(hwnd, L"复制失败。", L"提示");
        return 0;
      }
      if (id == 3) {
        std::wstring path;
        if (!SaveFileDialog(hwnd, s->default_name, &path)) return 0;
        std::wstring err;
        if (!WriteUtf8File(path, s->markdown, &err)) {
          ShowInfoBox(hwnd, err.c_str(), L"导出失败");
          return 0;
        }
        ShowInfoBox(hwnd, L"已导出。", L"提示");
        return 0;
      }
      if (id == 4) {
        DestroyWindow(hwnd);
        return 0;
      }
      return 0;
    }
    case WM_DESTROY: {
      if (s && s->font) DeleteObject(s->font);
      return 0;
    }
  }
  return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void ShowReportWindow(HWND owner, const std::wstring& title, const std::wstring& md, const std::wstring& default_name) {
  WNDCLASSW wc{};
  wc.lpfnWndProc = ReportWndProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.lpszClassName = L"WorkLogLiteReportWnd";
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
  RegisterClassW(&wc);

  ReportWindowState state{};
  state.markdown = md;
  state.default_name = default_name;

  HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME, wc.lpszClassName, title.c_str(),
                              WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                              CW_USEDEFAULT, CW_USEDEFAULT, 900, 650,
                              owner, nullptr, wc.hInstance, &state);
  if (!hwnd) {
    ShowLastErrorBox(owner, L"CreateWindowEx(Report)");
    return;
  }

  EnableWindow(owner, FALSE);
  // Modal loop without posting WM_QUIT (WM_QUIT would terminate the whole app).
  while (IsWindow(hwnd)) {
    MSG msg{};
    while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
      if (msg.message == WM_QUIT) {
        // Re-post and exit; caller (main loop) will handle it.
        PostQuitMessage((int)msg.wParam);
        EnableWindow(owner, TRUE);
        return;
      }
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
    WaitMessage();
  }
  EnableWindow(owner, TRUE);
  SetForegroundWindow(owner);
}

struct PreviewWindowState {
  HWND hwnd{};
  HWND rich{};
  HWND btn_copy{};
  HWND btn_close{};
  HFONT font{};
  std::wstring markdown;
};

static void RichEditApplyFormat(HWND rich, LONG start, LONG end, const CHARFORMAT2W& cf) {
  if (!rich) return;
  CHARRANGE cr{};
  cr.cpMin = start;
  cr.cpMax = end;
  SendMessageW(rich, EM_EXSETSEL, 0, (LPARAM)&cr);
  SendMessageW(rich, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
}

static void RichEditSetAllDefault(HWND rich, HFONT font) {
  if (!rich) return;
  SendMessageW(rich, WM_SETFONT, (WPARAM)font, TRUE);
  CHARRANGE cr{0, -1};
  SendMessageW(rich, EM_EXSETSEL, 0, (LPARAM)&cr);
  CHARFORMAT2W cf{};
  cf.cbSize = sizeof(cf);
  cf.dwMask = CFM_FACE | CFM_SIZE | CFM_COLOR;
  cf.crTextColor = RGB(20, 20, 20);
  cf.yHeight = 10 * 20;  // 10pt
  wcscpy_s(cf.szFaceName, L"Segoe UI");
  SendMessageW(rich, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
}

static std::vector<std::pair<LONG, LONG>> FindPairedDelims(const std::wstring& s, wchar_t delim, size_t line_start, size_t line_end) {
  // Return ranges [a,b) between paired `delim` in [line_start,line_end)
  std::vector<std::pair<LONG, LONG>> out;
  size_t i = line_start;
  while (i < line_end) {
    size_t a = s.find(delim, i);
    if (a == std::wstring::npos || a >= line_end) break;
    size_t b = s.find(delim, a + 1);
    if (b == std::wstring::npos || b >= line_end) break;
    out.push_back({(LONG)a, (LONG)(b + 1)});  // include backticks
    i = b + 1;
  }
  return out;
}

static std::vector<std::pair<LONG, LONG>> FindPairedToken(const std::wstring& s, const std::wstring& token, size_t line_start, size_t line_end) {
  // Return ranges [a,b) for paired token occurrences in [line_start,line_end)
  std::vector<std::pair<LONG, LONG>> out;
  if (token.empty()) return out;
  size_t i = line_start;
  while (i < line_end) {
    size_t a = s.find(token, i);
    if (a == std::wstring::npos || a >= line_end) break;
    size_t b = s.find(token, a + token.size());
    if (b == std::wstring::npos || b >= line_end) break;
    out.push_back({(LONG)a, (LONG)(b + token.size())});
    i = b + token.size();
  }
  return out;
}

static void RenderMarkdownToRichEdit(HWND rich, HFONT font, const std::wstring& md) {
  SetWindowTextW(rich, md.c_str());
  RichEditSetAllDefault(rich, font);

  // Basic styling: headings, code fences, inline code (backticks).
  bool in_code = false;
  size_t pos = 0;
  while (pos < md.size()) {
    size_t line_end = md.find(L'\n', pos);
    if (line_end == std::wstring::npos) line_end = md.size();

    // Determine line text range [pos,line_end)
    std::wstring_view line(md.data() + pos, line_end - pos);
    if (!line.empty() && line.back() == L'\r') line = line.substr(0, line.size() - 1);

    if (line.rfind(L"```", 0) == 0) {
      // Style fence line and toggle state.
      CHARFORMAT2W cf{};
      cf.cbSize = sizeof(cf);
      cf.dwMask = CFM_FACE | CFM_SIZE | CFM_COLOR;
      cf.crTextColor = RGB(120, 120, 120);
      cf.yHeight = 9 * 20;
      wcscpy_s(cf.szFaceName, L"Consolas");
      RichEditApplyFormat(rich, (LONG)pos, (LONG)line_end, cf);
      in_code = !in_code;
    } else if (in_code) {
      CHARFORMAT2W cf{};
      cf.cbSize = sizeof(cf);
      cf.dwMask = CFM_FACE | CFM_SIZE | CFM_COLOR;
      cf.crTextColor = RGB(40, 40, 40);
      cf.yHeight = 9 * 20;
      wcscpy_s(cf.szFaceName, L"Consolas");
      RichEditApplyFormat(rich, (LONG)pos, (LONG)line_end, cf);
    } else if (line.rfind(L"# ", 0) == 0) {
      CHARFORMAT2W cf{};
      cf.cbSize = sizeof(cf);
      cf.dwMask = CFM_BOLD | CFM_SIZE | CFM_COLOR;
      cf.dwEffects = CFE_BOLD;
      cf.crTextColor = RGB(10, 60, 120);
      cf.yHeight = 14 * 20;
      RichEditApplyFormat(rich, (LONG)pos, (LONG)line_end, cf);
    } else if (line.rfind(L"## ", 0) == 0) {
      CHARFORMAT2W cf{};
      cf.cbSize = sizeof(cf);
      cf.dwMask = CFM_BOLD | CFM_SIZE | CFM_COLOR;
      cf.dwEffects = CFE_BOLD;
      cf.crTextColor = RGB(10, 60, 120);
      cf.yHeight = 12 * 20;
      RichEditApplyFormat(rich, (LONG)pos, (LONG)line_end, cf);
    } else {
      // bold: **text**
      {
        auto ranges = FindPairedToken(md, L"**", pos, line_end);
        for (const auto& r : ranges) {
          CHARFORMAT2W cf{};
          cf.cbSize = sizeof(cf);
          cf.dwMask = CFM_BOLD;
          cf.dwEffects = CFE_BOLD;
          RichEditApplyFormat(rich, r.first, r.second, cf);
        }
      }
      // italic: _text_
      {
        auto ranges = FindPairedToken(md, L"_", pos, line_end);
        for (const auto& r : ranges) {
          CHARFORMAT2W cf{};
          cf.cbSize = sizeof(cf);
          cf.dwMask = CFM_ITALIC;
          cf.dwEffects = CFE_ITALIC;
          RichEditApplyFormat(rich, r.first, r.second, cf);
        }
      }
      // inline code
      auto ranges = FindPairedDelims(md, L'`', pos, line_end);
      for (const auto& r : ranges) {
        CHARFORMAT2W cf{};
        cf.cbSize = sizeof(cf);
        cf.dwMask = CFM_FACE | CFM_SIZE | CFM_COLOR;
        cf.crTextColor = RGB(120, 60, 0);
        cf.yHeight = 9 * 20;
        wcscpy_s(cf.szFaceName, L"Consolas");
        RichEditApplyFormat(rich, r.first, r.second, cf);
      }
    }

    pos = (line_end < md.size()) ? (line_end + 1) : md.size();
  }

  // Restore selection to start.
  CHARRANGE cr{0, 0};
  SendMessageW(rich, EM_EXSETSEL, 0, (LPARAM)&cr);
}

static LRESULT CALLBACK PreviewWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
  auto* s = reinterpret_cast<PreviewWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
    case WM_CREATE: {
      auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
      s = reinterpret_cast<PreviewWindowState*>(cs->lpCreateParams);
      s->hwnd = hwnd;
      SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(s));

      s->font = CreateUiFont(hwnd);

      s->rich = CreateWindowExW(WS_EX_CLIENTEDGE, L"RICHEDIT50W", L"",
                                WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY |
                                    WS_VSCROLL | WS_TABSTOP,
                                0, 0, 10, 10, hwnd, (HMENU)1, nullptr, nullptr);
      s->btn_copy = CreateWindowExW(0, L"BUTTON", L"复制 Markdown",
                                    WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                    0, 0, 10, 10, hwnd, (HMENU)2, nullptr, nullptr);
      s->btn_close = CreateWindowExW(0, L"BUTTON", L"关闭",
                                     WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                     0, 0, 10, 10, hwnd, (HMENU)3, nullptr, nullptr);

      SetControlFont(s->btn_copy, s->font);
      SetControlFont(s->btn_close, s->font);

      RenderMarkdownToRichEdit(s->rich, s->font, s->markdown);
      return 0;
    }
    case WM_SIZE: {
      int w = LOWORD(lParam);
      int h = HIWORD(lParam);
      int pad = ScalePx(hwnd, 10);
      int btn_h = ScalePx(hwnd, 30);
      int btn_w1 = ScalePx(hwnd, 120);
      int btn_w2 = ScalePx(hwnd, 90);
      int y_btn = h - pad - btn_h;
      MoveWindow(s->rich, pad, pad, w - 2 * pad, y_btn - 2 * pad, TRUE);

      int x = w - pad - btn_w2;
      MoveWindow(s->btn_close, x, y_btn, btn_w2, btn_h, TRUE);
      x -= pad + btn_w1;
      MoveWindow(s->btn_copy, x, y_btn, btn_w1, btn_h, TRUE);
      return 0;
    }
    case WM_COMMAND: {
      int id = LOWORD(wParam);
      if (id == 2) {
        if (!CopyToClipboard(hwnd, s->markdown)) ShowInfoBox(hwnd, L"复制失败。", L"提示");
        return 0;
      }
      if (id == 3) {
        DestroyWindow(hwnd);
        return 0;
      }
      return 0;
    }
    case WM_DESTROY: {
      if (s && s->font) DeleteObject(s->font);
      return 0;
    }
  }
  return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void ShowPreviewWindow(HWND owner, const std::wstring& title, const std::wstring& markdown) {
  // Ensure rich edit class is available.
  LoadLibraryW(L"Msftedit.dll");

  WNDCLASSW wc{};
  wc.lpfnWndProc = PreviewWndProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.lpszClassName = L"WorkLogLitePreviewWnd";
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
  RegisterClassW(&wc);

  PreviewWindowState state{};
  state.markdown = markdown;

  HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME, wc.lpszClassName, title.c_str(),
                              WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                              CW_USEDEFAULT, CW_USEDEFAULT, 900, 650,
                              owner, nullptr, wc.hInstance, &state);
  if (!hwnd) {
    ShowLastErrorBox(owner, L"CreateWindowEx(Preview)");
    return;
  }

  EnableWindow(owner, FALSE);
  while (IsWindow(hwnd)) {
    MSG msg{};
    while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
      if (msg.message == WM_QUIT) {
        PostQuitMessage((int)msg.wParam);
        EnableWindow(owner, TRUE);
        return;
      }
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
    WaitMessage();
  }
  EnableWindow(owner, TRUE);
  SetForegroundWindow(owner);
}

enum class ReportMode {
  ByDate = 0,
  ByCategory = 1,
};

static void DoReport(AppState* s, bool quarter, ReportMode mode) {
  ReportRange r = quarter ? QuarterRangeForDate(s->selected) : YearRangeForDate(s->selected);
  std::wstring md;
  std::wstring err;
  bool ok = false;
  if (mode == ReportMode::ByCategory) ok = GenerateReportMarkdownByCategory(r, &md, &err);
  else ok = GenerateReportMarkdown(r, &md, &err);
  if (!ok) {
    ShowInfoBox(s->hwnd, err.c_str(), L"生成失败");
    return;
  }

  std::wstring title;
  if (quarter) title = (mode == ReportMode::ByCategory) ? L"本季度报告(按分类)" : L"本季度报告";
  else title = (mode == ReportMode::ByCategory) ? L"本年度报告(按分类)" : L"本年度报告";
  std::wstring file = L"report-" + FormatDateYYYYMMDD(r.start) + L"-" + FormatDateYYYYMMDD(r.end) + L".md";
  ShowReportWindow(s->hwnd, title, md, file);
}

static void DoExportCsv(AppState* s, bool quarter) {
  ReportRange r = quarter ? QuarterRangeForDate(s->selected) : YearRangeForDate(s->selected);
  std::wstring csv;
  std::wstring err;
  if (!GenerateReportCsvFlat(r, &csv, &err)) {
    ShowInfoBox(s->hwnd, err.c_str(), L"导出失败");
    return;
  }

  std::wstring default_name = L"report-" + FormatDateYYYYMMDD(r.start) + L"-" + FormatDateYYYYMMDD(r.end) + L".csv";
  std::wstring path;
  if (!SaveCsvFileDialog(s->hwnd, default_name, &path)) return;

  std::wstring werr;
  if (!WriteUtf8FileWithBom(path, csv, &werr)) {
    ShowInfoBox(s->hwnd, werr.c_str(), L"导出失败");
    return;
  }

  // Open the exported file with system default app (Excel/Numbers/etc).
  ShellExecuteW(nullptr, L"open", path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
}

static void OpenDataDir(AppState* s) {
  std::wstring dir = GetDataRootDir();
  EnsureDirExists(dir);
  ShellExecuteW(nullptr, L"open", dir.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
}

static void ShowHelp(AppState* s) {
  const wchar_t* text =
      L"需要填写什么？\n"
      L"- 分类: 必填，可自定义(例如 工作/会议/目标/沟通/研发)。\n"
      L"- 时间: 可选，格式 HH:MM。\n"
      L"- 标题: 建议填写，用于列表和汇总更好看。\n"
      L"- 内容: 富文本(支持段落/对齐/缩进/编号/项目符号)。\n"
      L"  在正文框内: Ctrl+B 加粗, Ctrl+I 斜体, Ctrl+U 下划线。\n"
      L"  列表: Ctrl+Shift+7 编号, Ctrl+Shift+8 项目符号。\n"
      L"  段落: Ctrl+Alt+L 左对齐, Ctrl+Alt+E 居中, Ctrl+Alt+R 右对齐。\n"
      L"\n"
      L"如何操作？\n"
      L"- 日历点一天: 切换到该日期记录。\n"
      L"- 列表选中一条: 右侧进入编辑。\n"
      L"- 查看 -> 预览: 查看 Markdown 渲染效果。\n"
      L"- 报告 -> 导出 CSV: 直接导出本季度/本年度 CSV，方便 Excel/脚本处理。\n"
      L"- 长期任务: 工具 -> 管理长期任务。任务会在其日期范围内每天自动出现在列表中，方便写进展。\n"
      L"- 周期会议: 工具 -> 管理周期会议。可设置会议模板，打开当天会自动填入。\n"
      L"- 从会议创建任务: 在会议条目上右键 -> 从此会议创建长期任务。\n"
      L"- 演示数据: 工具 -> 生成示例数据(演示)。\n"
      L"\n"
      L"快捷键:\n"
      L"- Ctrl+S 保存\n"
      L"- Ctrl+N 新增\n"
      L"- Delete 删除(会二次确认)\n"
      L"- Ctrl+P 预览当前条目\n"
      L"- Ctrl+Shift+P 预览当天全部\n"
      L"- Ctrl+K 管理分类\n"
      L"- Ctrl+T 管理长期任务\n"
      L"- Ctrl+R 管理周期会议\n"
      L"- F1 使用说明\n";
  ShowInfoBox(s->hwnd, text, L"WorkLogLite 使用说明");
}

struct CatWindowState {
  HWND hwnd{};
  HWND list{};
  HWND edit{};
  HWND btn_add{};
  HWND btn_rename{};
  HWND btn_del{};
  HWND btn_save{};
  HWND btn_cancel{};
  HFONT font{};

  AppState* app{};
  std::vector<std::wstring> working;
};

static std::wstring TrimWs(std::wstring s) {
  while (!s.empty() && iswspace(s.front())) s.erase(s.begin());
  while (!s.empty() && iswspace(s.back())) s.pop_back();
  return s;
}

static void CatRefreshList(CatWindowState* s) {
  SendMessageW(s->list, LB_RESETCONTENT, 0, 0);
  for (const auto& c : s->working) {
    SendMessageW(s->list, LB_ADDSTRING, 0, (LPARAM)c.c_str());
  }
}

static int CatGetSelectedIndex(CatWindowState* s) {
  return (int)SendMessageW(s->list, LB_GETCURSEL, 0, 0);
}

static void CatSetEditFromSelection(CatWindowState* s) {
  int idx = CatGetSelectedIndex(s);
  if (idx < 0 || idx >= (int)s->working.size()) return;
  SetWindowTextW(s->edit, s->working[(size_t)idx].c_str());
  SendMessageW(s->edit, EM_SETSEL, 0, -1);
  SetFocus(s->edit);
}

static LRESULT CALLBACK CatWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
  auto* s = reinterpret_cast<CatWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
    case WM_CREATE: {
      auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
      s = reinterpret_cast<CatWindowState*>(cs->lpCreateParams);
      s->hwnd = hwnd;
      SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(s));

      s->font = CreateUiFont(hwnd);

      s->list = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
                                WS_CHILD | WS_VISIBLE | LBS_NOTIFY | WS_VSCROLL | WS_TABSTOP,
                                0, 0, 10, 10, hwnd, (HMENU)1, nullptr, nullptr);
      s->edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | WS_TABSTOP,
                                0, 0, 10, 10, hwnd, (HMENU)2, nullptr, nullptr);
      s->btn_add = CreateWindowExW(0, L"BUTTON", L"新增",
                                   WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                   0, 0, 10, 10, hwnd, (HMENU)3, nullptr, nullptr);
      s->btn_rename = CreateWindowExW(0, L"BUTTON", L"修改",
                                      WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                      0, 0, 10, 10, hwnd, (HMENU)4, nullptr, nullptr);
      s->btn_del = CreateWindowExW(0, L"BUTTON", L"删除",
                                   WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                   0, 0, 10, 10, hwnd, (HMENU)5, nullptr, nullptr);
      s->btn_save = CreateWindowExW(0, L"BUTTON", L"保存并关闭",
                                    WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                    0, 0, 10, 10, hwnd, (HMENU)6, nullptr, nullptr);
      s->btn_cancel = CreateWindowExW(0, L"BUTTON", L"取消",
                                      WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                      0, 0, 10, 10, hwnd, (HMENU)7, nullptr, nullptr);

      for (HWND h : {s->list, s->edit, s->btn_add, s->btn_rename, s->btn_del, s->btn_save, s->btn_cancel}) {
        SetControlFont(h, s->font);
      }

      NormalizeCategories(&s->working);
      CatRefreshList(s);
      if (!s->working.empty()) {
        SendMessageW(s->list, LB_SETCURSEL, 0, 0);
        CatSetEditFromSelection(s);
      }
      return 0;
    }
    case WM_SIZE: {
      int w = LOWORD(lParam);
      int h = HIWORD(lParam);
      int pad = ScalePx(hwnd, 10);
      int row = ScalePx(hwnd, 28);
      int btn_w = ScalePx(hwnd, 110);
      int btn_h = ScalePx(hwnd, 30);

      int y = pad;
      MoveWindow(s->list, pad, y, w - 2 * pad, h - (pad * 3 + row + btn_h), TRUE);
      y += h - (pad * 3 + row + btn_h);
      y += pad;

      int edit_w = w - 2 * pad - (pad + btn_w) * 3;
      if (edit_w < ScalePx(hwnd, 120)) edit_w = ScalePx(hwnd, 120);
      MoveWindow(s->edit, pad, y, edit_w, row, TRUE);

      int x = pad + edit_w + pad;
      MoveWindow(s->btn_add, x, y, btn_w, btn_h, TRUE);
      x += btn_w + pad;
      MoveWindow(s->btn_rename, x, y, btn_w, btn_h, TRUE);
      x += btn_w + pad;
      MoveWindow(s->btn_del, x, y, btn_w, btn_h, TRUE);

      int y2 = h - pad - btn_h;
      MoveWindow(s->btn_cancel, w - pad - btn_w, y2, btn_w, btn_h, TRUE);
      MoveWindow(s->btn_save, w - pad * 2 - btn_w * 2, y2, btn_w, btn_h, TRUE);
      return 0;
    }
    case WM_COMMAND: {
      int id = LOWORD(wParam);
      int code = HIWORD(wParam);
      if (id == 1 && code == LBN_SELCHANGE) {
        CatSetEditFromSelection(s);
        return 0;
      }
      if (id == 1 && code == LBN_DBLCLK) {
        CatSetEditFromSelection(s);
        return 0;
      }

      if (id == 3) {  // add
        std::wstring t = TrimWs(GetWindowTextWString(s->edit));
        if (t.empty()) return 0;
        if (!CategoriesContains(s->working, t)) {
          s->working.push_back(t);
          NormalizeCategories(&s->working);
          CatRefreshList(s);
        }
        return 0;
      }
      if (id == 4) {  // rename
        int idx = CatGetSelectedIndex(s);
        if (idx < 0 || idx >= (int)s->working.size()) return 0;
        std::wstring t = TrimWs(GetWindowTextWString(s->edit));
        if (t.empty()) return 0;
        s->working[(size_t)idx] = t;
        NormalizeCategories(&s->working);
        CatRefreshList(s);
        return 0;
      }
      if (id == 5) {  // delete
        int idx = CatGetSelectedIndex(s);
        if (idx < 0 || idx >= (int)s->working.size()) return 0;
        s->working.erase(s->working.begin() + idx);
        NormalizeCategories(&s->working);
        CatRefreshList(s);
        if (!s->working.empty()) {
          int new_idx = idx;
          if (new_idx >= (int)s->working.size()) new_idx = (int)s->working.size() - 1;
          SendMessageW(s->list, LB_SETCURSEL, new_idx, 0);
          CatSetEditFromSelection(s);
        } else {
          SetWindowTextW(s->edit, L"");
        }
        return 0;
      }
      if (id == 6) {  // save & close
        if (s->working.empty()) {
          ShowInfoBox(hwnd, L"至少需要保留一个分类。", L"提示");
          return 0;
        }
        NormalizeCategories(&s->working);
        s->app->categories = s->working;
        RefreshCategoryCombo(s->app);

        std::wstring err;
        if (!SaveCategories(s->app->categories, &err)) {
          ShowInfoBox(hwnd, err.c_str(), L"保存失败");
          return 0;
        }
        DestroyWindow(hwnd);
        return 0;
      }
      if (id == 7) {  // cancel
        DestroyWindow(hwnd);
        return 0;
      }
      return 0;
    }
    case WM_DESTROY: {
      if (s && s->font) DeleteObject(s->font);
      return 0;
    }
  }
  return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void ShowManageCategoriesWindow(AppState* app) {
  WNDCLASSW wc{};
  wc.lpfnWndProc = CatWndProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.lpszClassName = L"WorkLogLiteCategoriesWnd";
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
  RegisterClassW(&wc);

  CatWindowState state{};
  state.app = app;
  state.working = app->categories;

  HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME, wc.lpszClassName, L"管理分类",
                              WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                              CW_USEDEFAULT, CW_USEDEFAULT, 720, 520,
                              app->hwnd, nullptr, wc.hInstance, &state);
  if (!hwnd) {
    ShowLastErrorBox(app->hwnd, L"CreateWindowEx(Categories)");
    return;
  }

  EnableWindow(app->hwnd, FALSE);
  while (IsWindow(hwnd)) {
    MSG msg{};
    while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
      if (msg.message == WM_QUIT) {
        PostQuitMessage((int)msg.wParam);
        EnableWindow(app->hwnd, TRUE);
        return;
      }
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
    WaitMessage();
  }
  EnableWindow(app->hwnd, TRUE);
  SetForegroundWindow(app->hwnd);
}

struct TaskWindowState {
  HWND hwnd{};
  HWND list{};
  HWND lbl_category{};
  HWND cb_category{};
  HWND lbl_title{};
  HWND ed_title{};
  HWND lbl_basis{};
  HWND ed_basis{};
  HWND lbl_desc{};
  HWND re_desc{};
  HWND lbl_start{};
  HWND dt_start{};
  HWND lbl_end{};
  HWND dt_end{};
  HWND lbl_status{};
  HWND cb_status{};
  HWND btn_del{};
  HWND btn_new{};
  HWND btn_save{};
  HWND btn_close{};
  HFONT font{};

  AppState* app{};
  std::vector<Task> working;
  int initial_select{-1};
};

static std::wstring FormatTaskListItem(const Task& t) {
  std::wstring s = L"[";
  s += StatusToCN(t.status);
  s += L"] ";
  s += t.title;
  s += L"  (";
  s += FormatDateYYYYMMDD(t.start);
  s += L" ~ ";
  s += FormatDateYYYYMMDD(t.end);
  s += L")";
  return s;
}

static void TaskRefreshList(TaskWindowState* s) {
  SendMessageW(s->list, LB_RESETCONTENT, 0, 0);
  for (const auto& t : s->working) {
    std::wstring item = FormatTaskListItem(t);
    SendMessageW(s->list, LB_ADDSTRING, 0, (LPARAM)item.c_str());
  }
}

static int TaskGetSelectedIndex(TaskWindowState* s) {
  return (int)SendMessageW(s->list, LB_GETCURSEL, 0, 0);
}

static void TaskFillFieldsFromSelection(TaskWindowState* s) {
  int idx = TaskGetSelectedIndex(s);
  if (idx < 0 || idx >= (int)s->working.size()) return;
  const Task& t = s->working[(size_t)idx];
  SetWindowTextW(s->cb_category, t.category.c_str());
  SetWindowTextW(s->ed_title, t.title.c_str());
  SetWindowTextW(s->ed_basis, t.basis.c_str());
  DateTime_SetSystemtime(s->dt_start, GDT_VALID, &t.start);
  DateTime_SetSystemtime(s->dt_end, GDT_VALID, &t.end);
  SendMessageW(s->cb_status, CB_SETCURSEL, ComboSelFromStatus(t.status), 0);

  if (!t.desc_rtf_b64.empty()) {
    std::string rtf;
    if (Base64Decode(t.desc_rtf_b64, &rtf)) {
      RichEditSetRtfBytes(s->re_desc, rtf);
    } else {
      SetEditText(s->re_desc, t.desc_plain);
    }
  } else {
    SetEditText(s->re_desc, t.desc_plain);
  }
}

static void TaskClearFields(TaskWindowState* s) {
  if (!s) return;
  // Keep existing category list; just reset the active values.
  if (!s->app->categories.empty()) SetWindowTextW(s->cb_category, s->app->categories[0].c_str());
  SetWindowTextW(s->ed_title, L"");
  SetWindowTextW(s->ed_basis, L"");
  SetEditText(s->re_desc, L"");
  SendMessageW(s->cb_status, CB_SETCURSEL, 1, 0);
  SYSTEMTIME now{};
  GetLocalTime(&now);
  DateTime_SetSystemtime(s->dt_start, GDT_VALID, &now);
  DateTime_SetSystemtime(s->dt_end, GDT_VALID, &now);
  SendMessageW(s->list, LB_SETCURSEL, (WPARAM)-1, 0);
}

static bool TaskReadFields(TaskWindowState* s, Task* out, std::wstring* err) {
  if (!out) return false;
  out->category = TrimWs(GetWindowTextWString(s->cb_category));
  out->title = TrimWs(GetWindowTextWString(s->ed_title));
  out->basis = TrimWs(GetWindowTextWString(s->ed_basis));
  out->desc_plain = GetWindowTextWString(s->re_desc);
  if (!out->desc_plain.empty()) {
    std::string rtf = RichEditGetRtfBytes(s->re_desc);
    if (!rtf.empty()) out->desc_rtf_b64 = Base64Encode(rtf);
  } else {
    out->desc_rtf_b64.clear();
  }

  SYSTEMTIME st{};
  SYSTEMTIME ed{};
  if (DateTime_GetSystemtime(s->dt_start, &st) != GDT_VALID) {
    if (err) *err = L"请选择开始日期。";
    return false;
  }
  if (DateTime_GetSystemtime(s->dt_end, &ed) != GDT_VALID) {
    if (err) *err = L"请选择结束日期。";
    return false;
  }
  if (st.wYear == 0 || ed.wYear == 0) {
    if (err) *err = L"日期无效。";
    return false;
  }
  if (st.wYear > ed.wYear ||
      (st.wYear == ed.wYear && st.wMonth > ed.wMonth) ||
      (st.wYear == ed.wYear && st.wMonth == ed.wMonth && st.wDay > ed.wDay)) {
    if (err) *err = L"开始日期不能晚于结束日期。";
    return false;
  }
  out->start = st;
  out->end = ed;
  int sel = (int)SendMessageW(s->cb_status, CB_GETCURSEL, 0, 0);
  out->status = StatusFromComboSel(sel);

  if (out->category.empty()) {
    if (err) *err = L"分类必填。";
    return false;
  }
  if (out->title.empty()) {
    if (err) *err = L"标题必填。";
    return false;
  }

  // Encourage traceability: tasks ideally have basis + description.
  std::wstring desc_trim = TrimWs(out->desc_plain);
  if (out->basis.empty() || desc_trim.empty()) {
    int r = MessageBoxW(s->hwnd,
                        L"建议为长期任务填写“依据”和“任务说明”，方便后续汇总与追溯。\n"
                        L"当前有字段为空，仍要继续保存吗？",
                        kAppTitle, MB_YESNO | MB_ICONQUESTION);
    if (r != IDYES) return false;
  }
  return true;
}

static std::wstring NewTaskId() {
  SYSTEMTIME st{};
  GetLocalTime(&st);
  wchar_t buf[64]{};
  wsprintfW(buf, L"t%04d%02d%02d-%02d%02d%02d",
            (int)st.wYear, (int)st.wMonth, (int)st.wDay, (int)st.wHour, (int)st.wMinute, (int)st.wSecond);
  return buf;
}

static const RecurringMeeting* FindRecurringTemplateByTitle(const AppState* s, const std::wstring& title) {
  if (!s) return nullptr;
  if (title.empty()) return nullptr;
  for (const auto& rm : s->recurring_meetings) {
    if (rm.title == title && !rm.template_plain.empty()) return &rm;
  }
  return nullptr;
}

static Task BuildTaskFromMeeting(const SYSTEMTIME& meeting_date, const Entry& meeting) {
  Task t{};
  t.id = NewTaskId();
  t.title = meeting.title.empty() ? L"会议行动项" : (L"会议行动项: " + meeting.title);

  // Default category: meetings are often project work; user can adjust.
  t.category = L"项目";
  if (!meeting.category.empty()) {
    // If meeting was a tech review, it usually maps to R&D work.
    if (meeting.category.find(L"技术") != std::wstring::npos || meeting.title.find(L"技术") != std::wstring::npos) {
      t.category = L"研发";
    }
  }

  t.basis = L"会议: " + FormatDateYYYYMMDD(meeting_date) + L" " + (meeting.title.empty() ? L"(无标题)" : meeting.title);

  std::wstring desc;
  desc += L"来源会议纪要(可编辑):\n";
  if (!meeting.body_plain.empty()) {
    desc += meeting.body_plain;
  } else {
    desc += L"- (会议正文为空，请补充)\n";
  }
  desc += L"\n\n建议拆解:\n- 目标:\n- 范围:\n- 验收标准:\n- 截止日期:\n- 风险/依赖:\n";
  t.desc_plain = desc;
  t.desc_rtf_b64.clear();

  t.start = meeting_date;
  t.end = AddDays(meeting_date, 30);
  t.status = EntryStatus::Todo;
  return t;
}

static void TaskPopulateCategoryCombo(TaskWindowState* s) {
  SendMessageW(s->cb_category, CB_RESETCONTENT, 0, 0);
  for (const auto& c : s->app->categories) {
    SendMessageW(s->cb_category, CB_ADDSTRING, 0, (LPARAM)c.c_str());
  }
  if (!s->app->categories.empty()) SetWindowTextW(s->cb_category, s->app->categories[0].c_str());
}

struct RecurringWindowState {
  HWND hwnd{};
  HWND list{};
  HWND lbl_category{};
  HWND cb_category{};
  HWND lbl_title{};
  HWND ed_title{};
  HWND lbl_weekday{};
  HWND cb_weekday{};
  HWND lbl_interval{};
  HWND cb_interval{};
  HWND lbl_start{};
  HWND dt_start{};
  HWND lbl_end{};
  HWND dt_end{};
  HWND lbl_start_time{};
  HWND cb_start_time{};
  HWND lbl_end_time{};
  HWND cb_end_time{};
  HWND lbl_template{};
  HWND ed_template{};
  HWND btn_del{};
  HWND btn_new{};
  HWND btn_save{};
  HWND btn_close{};
  HFONT font{};

  AppState* app{};
  std::vector<RecurringMeeting> working;
};

static const wchar_t* WeekdayNameCN(int d) {
  static const wchar_t* names[7] = {L"周日", L"周一", L"周二", L"周三", L"周四", L"周五", L"周六"};
  if (d < 0) d = 0;
  if (d > 6) d = 6;
  return names[d];
}

static std::wstring FormatRecurringListItem(const RecurringMeeting& m) {
  std::wstring s = m.title;
  s += L"  (";
  s += WeekdayNameCN(m.weekday);
  s += L", 每";
  wchar_t buf[16]{};
  wsprintfW(buf, L"%d", m.interval_weeks);
  s += buf;
  s += L"周)";
  return s;
}

static void RecRefreshList(RecurringWindowState* s) {
  SendMessageW(s->list, LB_RESETCONTENT, 0, 0);
  for (const auto& m : s->working) {
    std::wstring item = FormatRecurringListItem(m);
    SendMessageW(s->list, LB_ADDSTRING, 0, (LPARAM)item.c_str());
  }
}

static int RecGetSelectedIndex(RecurringWindowState* s) {
  return (int)SendMessageW(s->list, LB_GETCURSEL, 0, 0);
}

static void RecFillFieldsFromSelection(RecurringWindowState* s) {
  int idx = RecGetSelectedIndex(s);
  if (idx < 0 || idx >= (int)s->working.size()) return;
  const auto& m = s->working[(size_t)idx];
  SetWindowTextW(s->cb_category, m.category.c_str());
  SetWindowTextW(s->ed_title, m.title.c_str());
  SendMessageW(s->cb_weekday, CB_SETCURSEL, m.weekday, 0);
  int interval = m.interval_weeks;
  if (interval < 1) interval = 1;
  if (interval > 4) interval = 4;
  SendMessageW(s->cb_interval, CB_SETCURSEL, interval - 1, 0);
  DateTime_SetSystemtime(s->dt_start, GDT_VALID, &m.start);
  DateTime_SetSystemtime(s->dt_end, GDT_VALID, &m.end);
  SetWindowTextW(s->cb_start_time, m.start_time.c_str());
  SetWindowTextW(s->cb_end_time, m.end_time.c_str());
  SetWindowTextW(s->ed_template, m.template_plain.c_str());
}

static std::wstring NewRecurringId() {
  SYSTEMTIME st{};
  GetLocalTime(&st);
  wchar_t buf[64]{};
  wsprintfW(buf, L"r%04d%02d%02d-%02d%02d%02d",
            (int)st.wYear, (int)st.wMonth, (int)st.wDay, (int)st.wHour, (int)st.wMinute, (int)st.wSecond);
  return buf;
}

static bool RecReadFields(RecurringWindowState* s, RecurringMeeting* out, std::wstring* err) {
  if (!out) return false;
  out->category = TrimWs(GetWindowTextWString(s->cb_category));
  out->title = TrimWs(GetWindowTextWString(s->ed_title));
  out->weekday = (int)SendMessageW(s->cb_weekday, CB_GETCURSEL, 0, 0);
  out->interval_weeks = (int)SendMessageW(s->cb_interval, CB_GETCURSEL, 0, 0) + 1;
  out->start_time = TrimWs(GetWindowTextWString(s->cb_start_time));
  out->end_time = TrimWs(GetWindowTextWString(s->cb_end_time));
  out->template_plain = GetWindowTextWString(s->ed_template);

  SYSTEMTIME st{};
  SYSTEMTIME ed{};
  if (DateTime_GetSystemtime(s->dt_start, &st) != GDT_VALID) {
    if (err) *err = L"请选择开始日期。";
    return false;
  }
  if (DateTime_GetSystemtime(s->dt_end, &ed) != GDT_VALID) {
    if (err) *err = L"请选择结束日期。";
    return false;
  }
  if (st.wYear > ed.wYear ||
      (st.wYear == ed.wYear && st.wMonth > ed.wMonth) ||
      (st.wYear == ed.wYear && st.wMonth == ed.wMonth && st.wDay > ed.wDay)) {
    if (err) *err = L"开始日期不能晚于结束日期。";
    return false;
  }
  out->start = st;
  out->end = ed;

  if (out->title.empty()) {
    if (err) *err = L"会议标题必填。";
    return false;
  }
  if (out->category.empty()) out->category = L"会议";
  if (out->weekday < 0 || out->weekday > 6) out->weekday = 1;
  if (out->interval_weeks < 1) out->interval_weeks = 1;
  return true;
}

static void PopulateWeekdayCombo(HWND cb) {
  SendMessageW(cb, CB_RESETCONTENT, 0, 0);
  for (int i = 0; i < 7; i++) SendMessageW(cb, CB_ADDSTRING, 0, (LPARAM)WeekdayNameCN(i));
  SendMessageW(cb, CB_SETCURSEL, 1, 0);
}

static void PopulateIntervalCombo(HWND cb) {
  SendMessageW(cb, CB_RESETCONTENT, 0, 0);
  SendMessageW(cb, CB_ADDSTRING, 0, (LPARAM)L"每 1 周");
  SendMessageW(cb, CB_ADDSTRING, 0, (LPARAM)L"每 2 周");
  SendMessageW(cb, CB_ADDSTRING, 0, (LPARAM)L"每 3 周");
  SendMessageW(cb, CB_ADDSTRING, 0, (LPARAM)L"每 4 周");
  SendMessageW(cb, CB_SETCURSEL, 0, 0);
}

static LRESULT CALLBACK RecWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
  auto* s = reinterpret_cast<RecurringWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
    case WM_CREATE: {
      auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
      s = reinterpret_cast<RecurringWindowState*>(cs->lpCreateParams);
      s->hwnd = hwnd;
      SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(s));

      s->font = CreateUiFont(hwnd);

      s->list = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
                                WS_CHILD | WS_VISIBLE | LBS_NOTIFY | WS_VSCROLL | WS_TABSTOP,
                                0, 0, 10, 10, hwnd, (HMENU)1, nullptr, nullptr);

      s->lbl_category = CreateWindowExW(0, L"STATIC", L"类型/分类",
                                        WS_CHILD | WS_VISIBLE,
                                        0, 0, 10, 10, hwnd, (HMENU)20, nullptr, nullptr);
      s->cb_category = CreateWindowExW(0, L"COMBOBOX", L"",
                                       WS_CHILD | WS_VISIBLE | CBS_DROPDOWN | WS_TABSTOP | WS_VSCROLL,
                                       0, 0, 10, 10, hwnd, (HMENU)2, nullptr, nullptr);
      s->lbl_title = CreateWindowExW(0, L"STATIC", L"会议标题",
                                     WS_CHILD | WS_VISIBLE,
                                     0, 0, 10, 10, hwnd, (HMENU)21, nullptr, nullptr);
      s->ed_title = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                    WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | WS_TABSTOP,
                                    0, 0, 10, 10, hwnd, (HMENU)3, nullptr, nullptr);
      s->lbl_weekday = CreateWindowExW(0, L"STATIC", L"周几",
                                       WS_CHILD | WS_VISIBLE,
                                       0, 0, 10, 10, hwnd, (HMENU)22, nullptr, nullptr);
      s->cb_weekday = CreateWindowExW(0, L"COMBOBOX", L"",
                                      WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP,
                                      0, 0, 10, 10, hwnd, (HMENU)4, nullptr, nullptr);
      s->lbl_interval = CreateWindowExW(0, L"STATIC", L"间隔(周)",
                                        WS_CHILD | WS_VISIBLE,
                                        0, 0, 10, 10, hwnd, (HMENU)23, nullptr, nullptr);
      s->cb_interval = CreateWindowExW(0, L"COMBOBOX", L"",
                                       WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP,
                                       0, 0, 10, 10, hwnd, (HMENU)5, nullptr, nullptr);
      s->lbl_start = CreateWindowExW(0, L"STATIC", L"开始日期",
                                     WS_CHILD | WS_VISIBLE,
                                     0, 0, 10, 10, hwnd, (HMENU)24, nullptr, nullptr);
      s->dt_start = CreateWindowExW(0, DATETIMEPICK_CLASSW, L"",
                                    WS_CHILD | WS_VISIBLE | DTS_SHORTDATEFORMAT | WS_TABSTOP,
                                    0, 0, 10, 10, hwnd, (HMENU)6, nullptr, nullptr);
      s->lbl_end = CreateWindowExW(0, L"STATIC", L"结束日期",
                                   WS_CHILD | WS_VISIBLE,
                                   0, 0, 10, 10, hwnd, (HMENU)25, nullptr, nullptr);
      s->dt_end = CreateWindowExW(0, DATETIMEPICK_CLASSW, L"",
                                  WS_CHILD | WS_VISIBLE | DTS_SHORTDATEFORMAT | WS_TABSTOP,
                                  0, 0, 10, 10, hwnd, (HMENU)7, nullptr, nullptr);
      s->lbl_start_time = CreateWindowExW(0, L"STATIC", L"开始时间(可选)",
                                          WS_CHILD | WS_VISIBLE,
                                          0, 0, 10, 10, hwnd, (HMENU)26, nullptr, nullptr);
      s->cb_start_time = CreateWindowExW(0, L"COMBOBOX", L"",
                                         WS_CHILD | WS_VISIBLE | CBS_DROPDOWN | WS_TABSTOP | WS_VSCROLL,
                                         0, 0, 10, 10, hwnd, (HMENU)8, nullptr, nullptr);
      s->lbl_end_time = CreateWindowExW(0, L"STATIC", L"结束时间(可选)",
                                        WS_CHILD | WS_VISIBLE,
                                        0, 0, 10, 10, hwnd, (HMENU)27, nullptr, nullptr);
      s->cb_end_time = CreateWindowExW(0, L"COMBOBOX", L"",
                                        WS_CHILD | WS_VISIBLE | CBS_DROPDOWN | WS_TABSTOP | WS_VSCROLL,
                                        0, 0, 10, 10, hwnd, (HMENU)9, nullptr, nullptr);
      PopulateTimeCombo(s->cb_start_time);
      PopulateTimeCombo(s->cb_end_time);
      PopulateWeekdayCombo(s->cb_weekday);
      PopulateIntervalCombo(s->cb_interval);

      s->lbl_template = CreateWindowExW(0, L"STATIC", L"会议模板(支持 Markdown)",
                                        WS_CHILD | WS_VISIBLE,
                                        0, 0, 10, 10, hwnd, (HMENU)28, nullptr, nullptr);
      s->ed_template = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                       WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_DISABLENOSCROLL | WS_VSCROLL | WS_TABSTOP,
                                       0, 0, 10, 10, hwnd, (HMENU)15, nullptr, nullptr);

      s->btn_new = CreateWindowExW(0, L"BUTTON", L"新建",
                                   WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                   0, 0, 10, 10, hwnd, (HMENU)10, nullptr, nullptr);
      s->btn_save = CreateWindowExW(0, L"BUTTON", L"保存",
                                      WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                      0, 0, 10, 10, hwnd, (HMENU)11, nullptr, nullptr);
      s->btn_del = CreateWindowExW(0, L"BUTTON", L"删除",
                                   WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                   0, 0, 10, 10, hwnd, (HMENU)12, nullptr, nullptr);
      s->btn_close = CreateWindowExW(0, L"BUTTON", L"关闭",
                                     WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                     0, 0, 10, 10, hwnd, (HMENU)13, nullptr, nullptr);

      for (HWND h : {s->list,
                     s->lbl_category, s->cb_category,
                     s->lbl_title, s->ed_title,
                     s->lbl_weekday, s->cb_weekday,
                     s->lbl_interval, s->cb_interval,
                     s->lbl_start, s->dt_start,
                     s->lbl_end, s->dt_end,
                     s->lbl_start_time, s->cb_start_time,
                     s->lbl_end_time, s->cb_end_time,
                     s->lbl_template, s->ed_template,
                     s->btn_new, s->btn_save, s->btn_del, s->btn_close}) {
        SetControlFont(h, s->font);
      }

      // Categories
      SendMessageW(s->cb_category, CB_RESETCONTENT, 0, 0);
      for (const auto& c : s->app->categories) SendMessageW(s->cb_category, CB_ADDSTRING, 0, (LPARAM)c.c_str());
      if (!s->app->categories.empty()) SetWindowTextW(s->cb_category, s->app->categories[0].c_str());

      SYSTEMTIME now{};
      GetLocalTime(&now);
      DateTime_SetSystemtime(s->dt_start, GDT_VALID, &now);
      DateTime_SetSystemtime(s->dt_end, GDT_VALID, &now);

      SetEditCue(GetComboEditHandle(s->cb_category), L"例如：例会/技术审查/项目会议");
      SetEditCue(s->ed_title, L"会议标题，例如：技术审查会");
      SetEditCue(s->ed_template, L"会议模板(可选，支持 Markdown)。打开当天会议时会自动填入，便于记录。");

      RecRefreshList(s);
      if (!s->working.empty()) {
        SendMessageW(s->list, LB_SETCURSEL, 0, 0);
        RecFillFieldsFromSelection(s);
      }
      return 0;
    }
    case WM_SIZE: {
      int w = LOWORD(lParam);
      int h = HIWORD(lParam);
      int pad = ScalePx(hwnd, 10);
      int row = ScalePx(hwnd, 28);
      int lbl = ScalePx(hwnd, 18);
      int gap = ScalePx(hwnd, 2);
      int btn_h = ScalePx(hwnd, 30);
      int btn_w = ScalePx(hwnd, 92);

      int left_w = ScalePx(hwnd, 320);
      MoveWindow(s->list, pad, pad, left_w, h - pad * 2, TRUE);

      int x = pad + left_w + pad;
      int right_w = w - x - pad;
      int y = pad;

      MoveWindow(s->lbl_category, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      MoveWindow(s->cb_category, x, y, right_w, row * 6, TRUE);
      y += row + pad;

      MoveWindow(s->lbl_title, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      MoveWindow(s->ed_title, x, y, right_w, row, TRUE);
      y += row + pad;

      MoveWindow(s->lbl_weekday, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      MoveWindow(s->cb_weekday, x, y, right_w, row, TRUE);
      y += row + pad;

      MoveWindow(s->lbl_interval, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      MoveWindow(s->cb_interval, x, y, right_w, row, TRUE);
      y += row + pad;

      MoveWindow(s->lbl_start, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      MoveWindow(s->dt_start, x, y, right_w, row, TRUE);
      y += row + pad;

      MoveWindow(s->lbl_end, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      MoveWindow(s->dt_end, x, y, right_w, row, TRUE);
      y += row + pad;

      MoveWindow(s->lbl_start_time, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      MoveWindow(s->cb_start_time, x, y, right_w, row * 6, TRUE);
      y += row + pad;

      MoveWindow(s->lbl_end_time, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      MoveWindow(s->cb_end_time, x, y, right_w, row * 6, TRUE);
      y += row + pad;

      MoveWindow(s->lbl_template, x, y, right_w, lbl, TRUE);
      y += lbl + gap;

      int tmpl_h = h - y - (pad + btn_h);
      if (tmpl_h < ScalePx(hwnd, 120)) tmpl_h = ScalePx(hwnd, 120);
      MoveWindow(s->ed_template, x, y, right_w, tmpl_h, TRUE);
      y += tmpl_h + pad;

      int needed = btn_w * 4 + pad * 3;
      int bw = btn_w;
      if (right_w < needed) {
        bw = (right_w - pad * 3) / 4;
        if (bw < ScalePx(hwnd, 70)) bw = ScalePx(hwnd, 70);
      }
      int bx = x;
      MoveWindow(s->btn_new, bx, y, bw, btn_h, TRUE);
      bx += bw + pad;
      MoveWindow(s->btn_save, bx, y, bw, btn_h, TRUE);
      bx += bw + pad;
      MoveWindow(s->btn_del, bx, y, bw, btn_h, TRUE);
      bx += bw + pad;
      MoveWindow(s->btn_close, bx, y, bw, btn_h, TRUE);
      return 0;
    }
    case WM_COMMAND: {
      int id = LOWORD(wParam);
      int code = HIWORD(wParam);
      if (id == 1 && (code == LBN_SELCHANGE || code == LBN_DBLCLK)) {
        RecFillFieldsFromSelection(s);
        return 0;
      }
      if (id == 10) {  // new
        if (!s->app->categories.empty()) SetWindowTextW(s->cb_category, s->app->categories[0].c_str());
        SetWindowTextW(s->ed_title, L"");
        SendMessageW(s->cb_weekday, CB_SETCURSEL, 1, 0);
        SendMessageW(s->cb_interval, CB_SETCURSEL, 0, 0);
        SYSTEMTIME now{};
        GetLocalTime(&now);
        DateTime_SetSystemtime(s->dt_start, GDT_VALID, &now);
        DateTime_SetSystemtime(s->dt_end, GDT_VALID, &now);
        SetWindowTextW(s->cb_start_time, L"");
        SetWindowTextW(s->cb_end_time, L"");
        SetWindowTextW(s->ed_template, L"");
        SendMessageW(s->list, LB_SETCURSEL, (WPARAM)-1, 0);
        HWND h = GetComboEditHandle(s->cb_category);
        SetFocus(h ? h : s->cb_category);
        return 0;
      }
      if (id == 11) {  // save (add or update)
        int idx = RecGetSelectedIndex(s);
        RecurringMeeting m{};
        std::wstring e;
        if (!RecReadFields(s, &m, &e)) {
          ShowInfoBox(hwnd, e.c_str(), L"提示");
          return 0;
        }
        if (idx >= 0 && idx < (int)s->working.size()) {
          m.id = s->working[(size_t)idx].id;
          s->working[(size_t)idx] = m;
        } else {
          m.id = NewRecurringId();
          s->working.push_back(m);
          idx = (int)s->working.size() - 1;
        }

        std::wstring err;
        if (!SaveRecurringMeetings(s->working, &err)) {
          ShowInfoBox(hwnd, err.c_str(), L"保存失败");
          return 0;
        }
        s->app->recurring_meetings = s->working;

        RecRefreshList(s);
        SendMessageW(s->list, LB_SETCURSEL, idx, 0);
        RecFillFieldsFromSelection(s);
        return 0;
      }
      if (id == 12) {  // delete
        int idx = RecGetSelectedIndex(s);
        if (idx < 0 || idx >= (int)s->working.size()) return 0;
        int r = MessageBoxW(hwnd, L"确定删除该周期会议？", kAppTitle, MB_YESNO | MB_ICONWARNING);
        if (r != IDYES) return 0;
        s->working.erase(s->working.begin() + idx);
        std::wstring err;
        if (!SaveRecurringMeetings(s->working, &err)) {
          ShowInfoBox(hwnd, err.c_str(), L"保存失败");
          return 0;
        }
        s->app->recurring_meetings = s->working;

        RecRefreshList(s);
        if (!s->working.empty()) {
          int new_idx = idx;
          if (new_idx >= (int)s->working.size()) new_idx = (int)s->working.size() - 1;
          SendMessageW(s->list, LB_SETCURSEL, new_idx, 0);
          RecFillFieldsFromSelection(s);
        } else {
          SendMessageW(s->list, LB_SETCURSEL, (WPARAM)-1, 0);
          SetWindowTextW(s->ed_title, L"");
          SetWindowTextW(s->ed_template, L"");
        }
        return 0;
      }
      if (id == 13) {  // close
        DestroyWindow(hwnd);
        return 0;
      }
      return 0;
    }
    case WM_DESTROY: {
      if (s && s->font) DeleteObject(s->font);
      return 0;
    }
  }
  return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void ShowManageRecurringMeetingsWindow(AppState* app) {
  WNDCLASSW wc{};
  wc.lpfnWndProc = RecWndProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.lpszClassName = L"WorkLogLiteRecurringWnd";
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
  RegisterClassW(&wc);

  RecurringWindowState state{};
  state.app = app;
  state.working = app->recurring_meetings;

  HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME, wc.lpszClassName, L"管理周期会议",
                              WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                              CW_USEDEFAULT, CW_USEDEFAULT, 920, 650,
                              app->hwnd, nullptr, wc.hInstance, &state);
  if (!hwnd) {
    ShowLastErrorBox(app->hwnd, L"CreateWindowEx(Recurring)");
    return;
  }

  EnableWindow(app->hwnd, FALSE);
  while (IsWindow(hwnd)) {
    MSG msg{};
    while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
      if (msg.message == WM_QUIT) {
        PostQuitMessage((int)msg.wParam);
        EnableWindow(app->hwnd, TRUE);
        return;
      }
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
    WaitMessage();
  }
  EnableWindow(app->hwnd, TRUE);
  SetForegroundWindow(app->hwnd);
}

static LRESULT CALLBACK TaskWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
  auto* s = reinterpret_cast<TaskWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
    case WM_CREATE: {
      auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
      s = reinterpret_cast<TaskWindowState*>(cs->lpCreateParams);
      s->hwnd = hwnd;
      SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(s));

      s->font = CreateUiFont(hwnd);

      s->list = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
                                WS_CHILD | WS_VISIBLE | LBS_NOTIFY | WS_VSCROLL | WS_TABSTOP,
                                0, 0, 10, 10, hwnd, (HMENU)1, nullptr, nullptr);

      s->lbl_category = CreateWindowExW(0, L"STATIC", L"类型/分类",
                                        WS_CHILD | WS_VISIBLE,
                                        0, 0, 10, 10, hwnd, (HMENU)20, nullptr, nullptr);
      s->cb_category = CreateWindowExW(0, L"COMBOBOX", L"",
                                       WS_CHILD | WS_VISIBLE | CBS_DROPDOWN | WS_TABSTOP | WS_VSCROLL,
                                       0, 0, 10, 10, hwnd, (HMENU)2, nullptr, nullptr);
      s->lbl_title = CreateWindowExW(0, L"STATIC", L"任务标题",
                                     WS_CHILD | WS_VISIBLE,
                                     0, 0, 10, 10, hwnd, (HMENU)21, nullptr, nullptr);
      s->ed_title = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                    WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | WS_TABSTOP,
                                    0, 0, 10, 10, hwnd, (HMENU)3, nullptr, nullptr);
      s->lbl_basis = CreateWindowExW(0, L"STATIC", L"依据(建议填写)",
                                     WS_CHILD | WS_VISIBLE,
                                     0, 0, 10, 10, hwnd, (HMENU)22, nullptr, nullptr);
      s->ed_basis = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                    WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | WS_TABSTOP,
                                    0, 0, 10, 10, hwnd, (HMENU)15, nullptr, nullptr);
      s->lbl_start = CreateWindowExW(0, L"STATIC", L"开始日期",
                                     WS_CHILD | WS_VISIBLE,
                                     0, 0, 10, 10, hwnd, (HMENU)23, nullptr, nullptr);
      s->re_desc = CreateWindowExW(WS_EX_CLIENTEDGE, L"RICHEDIT50W", L"",
                                   WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_DISABLENOSCROLL | WS_VSCROLL | WS_TABSTOP,
                                   0, 0, 10, 10, hwnd, (HMENU)16, nullptr, nullptr);
      s->dt_start = CreateWindowExW(0, DATETIMEPICK_CLASSW, L"",
                                    WS_CHILD | WS_VISIBLE | DTS_SHORTDATEFORMAT | WS_TABSTOP,
                                    0, 0, 10, 10, hwnd, (HMENU)4, nullptr, nullptr);
      s->lbl_end = CreateWindowExW(0, L"STATIC", L"结束日期",
                                   WS_CHILD | WS_VISIBLE,
                                   0, 0, 10, 10, hwnd, (HMENU)24, nullptr, nullptr);
      s->dt_end = CreateWindowExW(0, DATETIMEPICK_CLASSW, L"",
                                  WS_CHILD | WS_VISIBLE | DTS_SHORTDATEFORMAT | WS_TABSTOP,
                                  0, 0, 10, 10, hwnd, (HMENU)5, nullptr, nullptr);
      s->lbl_status = CreateWindowExW(0, L"STATIC", L"状态(用于看板颜色)",
                                      WS_CHILD | WS_VISIBLE,
                                      0, 0, 10, 10, hwnd, (HMENU)25, nullptr, nullptr);
      s->cb_status = CreateWindowExW(0, L"COMBOBOX", L"",
                                     WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP,
                                     0, 0, 10, 10, hwnd, (HMENU)6, nullptr, nullptr);
      SendMessageW(s->cb_status, CB_ADDSTRING, 0, (LPARAM)L"无");
      SendMessageW(s->cb_status, CB_ADDSTRING, 0, (LPARAM)L"未开始");
      SendMessageW(s->cb_status, CB_ADDSTRING, 0, (LPARAM)L"进行中");
      SendMessageW(s->cb_status, CB_ADDSTRING, 0, (LPARAM)L"阻塞");
      SendMessageW(s->cb_status, CB_ADDSTRING, 0, (LPARAM)L"已完成");
      SendMessageW(s->cb_status, CB_SETCURSEL, 1, 0);

      s->lbl_desc = CreateWindowExW(0, L"STATIC", L"任务说明(支持富文本，可做段落/编号/缩进)",
                                    WS_CHILD | WS_VISIBLE,
                                    0, 0, 10, 10, hwnd, (HMENU)26, nullptr, nullptr);

      s->btn_new = CreateWindowExW(0, L"BUTTON", L"新建",
                                   WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                   0, 0, 10, 10, hwnd, (HMENU)7, nullptr, nullptr);
      s->btn_save = CreateWindowExW(0, L"BUTTON", L"保存",
                                      WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                      0, 0, 10, 10, hwnd, (HMENU)8, nullptr, nullptr);
      s->btn_del = CreateWindowExW(0, L"BUTTON", L"删除",
                                   WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                   0, 0, 10, 10, hwnd, (HMENU)9, nullptr, nullptr);
      s->btn_close = CreateWindowExW(0, L"BUTTON", L"关闭",
                                     WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                     0, 0, 10, 10, hwnd, (HMENU)10, nullptr, nullptr);

      for (HWND h : {s->list, s->lbl_category, s->cb_category, s->lbl_title, s->ed_title, s->lbl_basis, s->ed_basis,
                     s->lbl_start, s->dt_start, s->lbl_end, s->dt_end, s->lbl_status, s->cb_status,
                     s->lbl_desc, s->re_desc,
                     s->btn_new, s->btn_save, s->btn_del, s->btn_close}) {
        SetControlFont(h, s->font);
      }

      TaskPopulateCategoryCombo(s);
      TaskRefreshList(s);
      if (!s->working.empty()) {
        int sel = s->initial_select;
        if (sel < 0 || sel >= (int)s->working.size()) sel = 0;
        SendMessageW(s->list, LB_SETCURSEL, sel, 0);
        TaskFillFieldsFromSelection(s);
      } else {
        SYSTEMTIME now{};
        GetLocalTime(&now);
        DateTime_SetSystemtime(s->dt_start, GDT_VALID, &now);
        DateTime_SetSystemtime(s->dt_end, GDT_VALID, &now);
      }

      SetEditCue(GetComboEditHandle(s->cb_category), L"任务分类，例如：项目/研发/学习");
      SetEditCue(s->ed_title, L"任务标题，例如：重构支付模块");
      SetEditCue(s->ed_basis, L"依据(可选)：会议/需求/故障单/邮件等");
      SetEditCue(s->re_desc, L"任务说明(可选)：目标、范围、验收标准、风险、依赖等");
      return 0;
    }
    case WM_SIZE: {
      int w = LOWORD(lParam);
      int h = HIWORD(lParam);
      int pad = ScalePx(hwnd, 10);
      int row = ScalePx(hwnd, 28);
      int lbl = ScalePx(hwnd, 18);
      int gap = ScalePx(hwnd, 2);
      int btn_h = ScalePx(hwnd, 30);
      int btn_w = ScalePx(hwnd, 92);

      int left_w = ScalePx(hwnd, 300);
      MoveWindow(s->list, pad, pad, left_w, h - pad * 2, TRUE);

      int x = pad + left_w + pad;
      int right_w = w - x - pad;
      int y = pad;

      MoveWindow(s->lbl_category, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      MoveWindow(s->cb_category, x, y, right_w, row * 6, TRUE);
      y += row + pad;

      MoveWindow(s->lbl_title, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      MoveWindow(s->ed_title, x, y, right_w, row, TRUE);
      y += row + pad;

      MoveWindow(s->lbl_basis, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      MoveWindow(s->ed_basis, x, y, right_w, row, TRUE);
      y += row + pad;

      MoveWindow(s->lbl_start, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      MoveWindow(s->dt_start, x, y, right_w, row, TRUE);
      y += row + pad;

      MoveWindow(s->lbl_end, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      MoveWindow(s->dt_end, x, y, right_w, row, TRUE);
      y += row + pad;

      MoveWindow(s->lbl_status, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      MoveWindow(s->cb_status, x, y, right_w, row * 6, TRUE);
      y += row + pad;

      MoveWindow(s->lbl_desc, x, y, right_w, lbl, TRUE);
      y += lbl + gap;

      int desc_h = h - y - (pad + btn_h);
      if (desc_h < ScalePx(hwnd, 120)) desc_h = ScalePx(hwnd, 120);
      MoveWindow(s->re_desc, x, y, right_w, desc_h, TRUE);
      y += desc_h + pad;

      int needed = btn_w * 4 + pad * 3;
      int bw = btn_w;
      if (right_w < needed) {
        bw = (right_w - pad * 3) / 4;
        if (bw < ScalePx(hwnd, 70)) bw = ScalePx(hwnd, 70);
      }
      int bx = x;
      MoveWindow(s->btn_new, bx, y, bw, btn_h, TRUE);
      bx += bw + pad;
      MoveWindow(s->btn_save, bx, y, bw, btn_h, TRUE);
      bx += bw + pad;
      MoveWindow(s->btn_del, bx, y, bw, btn_h, TRUE);
      bx += bw + pad;
      MoveWindow(s->btn_close, bx, y, bw, btn_h, TRUE);
      return 0;
    }

    case WM_CONTEXTMENU: {
      if (!s) return 0;
      HWND src = (HWND)wParam;
      if (src != s->re_desc) return DefWindowProcW(hwnd, msg, wParam, lParam);

      POINT pt{};
      if (lParam == (LPARAM)-1) {
        GetCursorPos(&pt);
      } else {
        pt.x = GET_X_LPARAM(lParam);
        pt.y = GET_Y_LPARAM(lParam);
      }

      HMENU menu = CreatePopupMenu();
      if (!menu) return 0;
      AppendMenuW(menu, MF_STRING, IDM_FMT_BOLD, L"加粗\tCtrl+B");
      AppendMenuW(menu, MF_STRING, IDM_FMT_ITALIC, L"斜体\tCtrl+I");
      AppendMenuW(menu, MF_STRING, IDM_FMT_UNDERLINE, L"下划线\tCtrl+U");
      AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
      AppendMenuW(menu, MF_STRING, IDM_FMT_NUMBERING, L"编号");
      AppendMenuW(menu, MF_STRING, IDM_FMT_BULLET, L"项目符号");
      AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
      AppendMenuW(menu, MF_STRING, IDM_FMT_INDENT_INC, L"增加缩进");
      AppendMenuW(menu, MF_STRING, IDM_FMT_INDENT_DEC, L"减少缩进");
      AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
      AppendMenuW(menu, MF_STRING, IDM_FMT_CLEAR, L"清除格式");

      UINT cmd = TrackPopupMenu(menu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, 0, hwnd, nullptr);
      DestroyMenu(menu);
      if (!cmd) return 0;

      if (cmd == IDM_FMT_BOLD) RichToggleCharEffect(s->re_desc, CFE_BOLD);
      else if (cmd == IDM_FMT_ITALIC) RichToggleCharEffect(s->re_desc, CFE_ITALIC);
      else if (cmd == IDM_FMT_UNDERLINE) RichToggleCharEffect(s->re_desc, CFE_UNDERLINE);
      else if (cmd == IDM_FMT_BULLET) RichToggleBullet(s->re_desc);
      else if (cmd == IDM_FMT_NUMBERING) RichToggleNumbering(s->re_desc);
      else if (cmd == IDM_FMT_INDENT_INC) RichAdjustIndent(s->re_desc, 360);
      else if (cmd == IDM_FMT_INDENT_DEC) RichAdjustIndent(s->re_desc, -360);
      else if (cmd == IDM_FMT_CLEAR) RichClearFormatting(s->re_desc);
      return 0;
    }

    case WM_COMMAND: {
      int id = LOWORD(wParam);
      int code = HIWORD(wParam);
      if (id == 1 && (code == LBN_SELCHANGE || code == LBN_DBLCLK)) {
        TaskFillFieldsFromSelection(s);
        return 0;
      }
      if (id == 7) {  // new
        TaskClearFields(s);
        HWND h = GetComboEditHandle(s->cb_category);
        SetFocus(h ? h : s->cb_category);
        return 0;
      }
      if (id == 8) {  // save (add or update)
        Task t{};
        std::wstring e;
        if (!TaskReadFields(s, &t, &e)) {
          ShowInfoBox(hwnd, e.c_str(), L"提示");
          return 0;
        }
        int idx = TaskGetSelectedIndex(s);
        if (idx >= 0 && idx < (int)s->working.size()) {
          t.id = s->working[(size_t)idx].id;
          s->working[(size_t)idx] = t;
        } else {
          t.id = NewTaskId();
          s->working.push_back(t);
          idx = (int)s->working.size() - 1;
        }

        std::wstring err;
        if (!SaveTasks(s->working, &err)) {
          ShowInfoBox(hwnd, err.c_str(), L"保存失败");
          return 0;
        }
        s->app->tasks = s->working;

        TaskRefreshList(s);
        SendMessageW(s->list, LB_SETCURSEL, idx, 0);
        TaskFillFieldsFromSelection(s);
        return 0;
      }
      if (id == 9) {  // delete
        int idx = TaskGetSelectedIndex(s);
        if (idx < 0 || idx >= (int)s->working.size()) return 0;
        int r = MessageBoxW(hwnd, L"确定删除该长期任务？", kAppTitle, MB_YESNO | MB_ICONWARNING);
        if (r != IDYES) return 0;
        s->working.erase(s->working.begin() + idx);
        std::wstring err;
        if (!SaveTasks(s->working, &err)) {
          ShowInfoBox(hwnd, err.c_str(), L"保存失败");
          return 0;
        }
        s->app->tasks = s->working;

        TaskRefreshList(s);
        if (!s->working.empty()) {
          int new_idx = idx;
          if (new_idx >= (int)s->working.size()) new_idx = (int)s->working.size() - 1;
          SendMessageW(s->list, LB_SETCURSEL, new_idx, 0);
          TaskFillFieldsFromSelection(s);
        } else {
          TaskClearFields(s);
        }
        return 0;
      }
      if (id == 10) {  // close
        DestroyWindow(hwnd);
        return 0;
      }
      return 0;
    }
    case WM_DESTROY: {
      if (s && s->font) DeleteObject(s->font);
      return 0;
    }
  }
  return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void ShowManageTasksWindow(AppState* app, const Task* preinsert) {
  WNDCLASSW wc{};
  wc.lpfnWndProc = TaskWndProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.lpszClassName = L"WorkLogLiteTasksWnd";
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
  RegisterClassW(&wc);

  TaskWindowState state{};
  state.app = app;
  state.working = app->tasks;
  if (preinsert) {
    Task t = *preinsert;
    if (t.id.empty()) t.id = NewTaskId();
    state.working.push_back(std::move(t));
    state.initial_select = (int)state.working.size() - 1;
  }

  HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME, wc.lpszClassName, L"管理长期任务",
                              WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                              CW_USEDEFAULT, CW_USEDEFAULT, 900, 600,
                              app->hwnd, nullptr, wc.hInstance, &state);
  if (!hwnd) {
    ShowLastErrorBox(app->hwnd, L"CreateWindowEx(Tasks)");
    return;
  }

  EnableWindow(app->hwnd, FALSE);
  while (IsWindow(hwnd)) {
    MSG msg{};
    while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
      if (msg.message == WM_QUIT) {
        PostQuitMessage((int)msg.wParam);
        EnableWindow(app->hwnd, TRUE);
        return;
      }
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
    WaitMessage();
  }
  EnableWindow(app->hwnd, TRUE);
  SetForegroundWindow(app->hwnd);
}

static void Layout(AppState* s) {
  RECT rc{};
  GetClientRect(s->hwnd, &rc);
  int w = rc.right - rc.left;
  int h = rc.bottom - rc.top;

  int status_h = 0;
  if (s->status) {
    RECT rs{};
    GetWindowRect(s->status, &rs);
    status_h = rs.bottom - rs.top;
    h -= status_h;
  }

  int pad = ScalePx(s->hwnd, 10);
  int left_w = ScalePx(s->hwnd, 260);
  int top_h = ScalePx(s->hwnd, 240);
  int row_h = ScalePx(s->hwnd, 26);
  int lbl_h = ScalePx(s->hwnd, 18);
  int btn_h = ScalePx(s->hwnd, 30);
  int btn_w = ScalePx(s->hwnd, 72);

  MoveWindow(s->cal, pad, pad, left_w, top_h, TRUE);

  int right_x = pad + left_w + pad;
  int right_w = w - right_x - pad;
  int y = pad;

  MoveWindow(s->st_day, right_x, y, right_w, lbl_h, TRUE);
  y += lbl_h + ScalePx(s->hwnd, 4);

  MoveWindow(s->list, right_x, y, right_w, top_h - lbl_h - ScalePx(s->hwnd, 4), TRUE);
  y += top_h + pad;

  // Labels for row
  int x = right_x;
  MoveWindow(s->lbl_category, x, y, ScalePx(s->hwnd, 140), lbl_h, TRUE);
  x += ScalePx(s->hwnd, 140) + pad;
  MoveWindow(s->lbl_start, x, y, ScalePx(s->hwnd, 70), lbl_h, TRUE);
  x += ScalePx(s->hwnd, 70) + pad;
  MoveWindow(s->lbl_end, x, y, ScalePx(s->hwnd, 70), lbl_h, TRUE);
  x += ScalePx(s->hwnd, 70) + pad;
  MoveWindow(s->lbl_status, x, y, ScalePx(s->hwnd, 90), lbl_h, TRUE);

  // Category / Start / End row
  y += lbl_h + ScalePx(s->hwnd, 2);
  x = right_x;
  // For combobox, the "height" controls dropdown height; keep it tall enough for suggestions.
  MoveWindow(s->cb_category, x, y, ScalePx(s->hwnd, 140), row_h * 10, TRUE);
  x += ScalePx(s->hwnd, 140) + pad;
  MoveWindow(s->ed_start, x, y, ScalePx(s->hwnd, 70), row_h * 10, TRUE);
  x += ScalePx(s->hwnd, 70) + pad;
  MoveWindow(s->ed_end, x, y, ScalePx(s->hwnd, 70), row_h * 10, TRUE);
  x += ScalePx(s->hwnd, 70) + pad;
  MoveWindow(s->cb_status, x, y, ScalePx(s->hwnd, 90), row_h * 10, TRUE);

  // Buttons on the right
  int bx = right_x + right_w - btn_w;
  MoveWindow(s->btn_new, bx, y, btn_w, btn_h, TRUE);
  bx -= pad + btn_w;
  MoveWindow(s->btn_save, bx, y, btn_w, btn_h, TRUE);
  bx -= pad + btn_w;
  MoveWindow(s->btn_preview, bx, y, btn_w, btn_h, TRUE);
  bx -= pad + btn_w;
  MoveWindow(s->btn_del, bx, y, btn_w, btn_h, TRUE);

  y += row_h + pad;

  MoveWindow(s->lbl_title, right_x, y, right_w, lbl_h, TRUE);
  y += lbl_h + ScalePx(s->hwnd, 2);
  MoveWindow(s->ed_title, right_x, y, right_w, row_h, TRUE);
  y += row_h + pad;

  MoveWindow(s->lbl_body, right_x, y, right_w, lbl_h, TRUE);
  y += lbl_h + ScalePx(s->hwnd, 2);
  MoveWindow(s->ed_body, right_x, y, right_w, h - y - pad, TRUE);
}

static void InitListColumns(HWND list) {
  LVCOLUMNW col{};
  col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;

  col.pszText = (LPWSTR)L"时间";
  col.cx = 110;
  col.iSubItem = 0;
  ListView_InsertColumn(list, 0, &col);

  col.pszText = (LPWSTR)L"分类";
  col.cx = 120;
  col.iSubItem = 1;
  ListView_InsertColumn(list, 1, &col);

  col.pszText = (LPWSTR)L"标题";
  col.cx = 420;
  col.iSubItem = 2;
  ListView_InsertColumn(list, 2, &col);

  col.pszText = (LPWSTR)L"状态";
  col.cx = 90;
  col.iSubItem = 3;
  ListView_InsertColumn(list, 3, &col);
}

static HMENU CreateAppMenu() {
  HMENU menu = CreateMenu();

  HMENU view = CreateMenu();
  AppendMenuW(view, MF_STRING, IDM_PREVIEW_ENTRY, L"预览当前条目\tCtrl+P");
  AppendMenuW(view, MF_STRING, IDM_PREVIEW_DAY, L"预览当天(全部)\tCtrl+Shift+P");
  AppendMenuW(menu, MF_POPUP, (UINT_PTR)view, L"查看");

  HMENU fmt = CreateMenu();
  AppendMenuW(fmt, MF_STRING, IDM_FMT_BOLD, L"加粗\tCtrl+B");
  AppendMenuW(fmt, MF_STRING, IDM_FMT_ITALIC, L"斜体\tCtrl+I");
  AppendMenuW(fmt, MF_STRING, IDM_FMT_UNDERLINE, L"下划线\tCtrl+U");
  AppendMenuW(fmt, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(fmt, MF_STRING, IDM_FMT_BULLET, L"项目符号\tCtrl+Shift+8");
  AppendMenuW(fmt, MF_STRING, IDM_FMT_NUMBERING, L"编号\tCtrl+Shift+7");
  AppendMenuW(fmt, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(fmt, MF_STRING, IDM_FMT_ALIGN_LEFT, L"左对齐\tCtrl+Alt+L");
  AppendMenuW(fmt, MF_STRING, IDM_FMT_ALIGN_CENTER, L"居中\tCtrl+Alt+E");
  AppendMenuW(fmt, MF_STRING, IDM_FMT_ALIGN_RIGHT, L"右对齐\tCtrl+Alt+R");
  AppendMenuW(fmt, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(fmt, MF_STRING, IDM_FMT_INDENT_INC, L"增加缩进\tCtrl+Alt+>");
  AppendMenuW(fmt, MF_STRING, IDM_FMT_INDENT_DEC, L"减少缩进\tCtrl+Alt+<");
  AppendMenuW(fmt, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(fmt, MF_STRING, IDM_FMT_CLEAR, L"清除格式");
  AppendMenuW(menu, MF_POPUP, (UINT_PTR)fmt, L"格式");

  HMENU report = CreateMenu();
  AppendMenuW(report, MF_STRING, IDM_REPORT_QUARTER, L"本季度\tCtrl+Q");
  AppendMenuW(report, MF_STRING, IDM_REPORT_YEAR, L"本年度\tCtrl+Y");
  AppendMenuW(report, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(report, MF_STRING, IDM_REPORT_QUARTER_BY_CAT, L"本季度(按分类)\tCtrl+Alt+Q");
  AppendMenuW(report, MF_STRING, IDM_REPORT_YEAR_BY_CAT, L"本年度(按分类)\tCtrl+Alt+Y");
  AppendMenuW(report, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(report, MF_STRING, IDM_EXPORT_QUARTER_CSV, L"导出本季度 CSV...");
  AppendMenuW(report, MF_STRING, IDM_EXPORT_YEAR_CSV, L"导出本年度 CSV...");
  AppendMenuW(menu, MF_POPUP, (UINT_PTR)report, L"报告");

  HMENU tools = CreateMenu();
  AppendMenuW(tools, MF_STRING, IDM_OPEN_DATA_DIR, L"打开数据目录");
  AppendMenuW(tools, MF_STRING, IDM_MANAGE_CATEGORIES, L"管理分类...\tCtrl+K");
  AppendMenuW(tools, MF_STRING, IDM_MANAGE_TASKS, L"管理长期任务...\tCtrl+T");
  AppendMenuW(tools, MF_STRING, IDM_MANAGE_RECURRING_MEETINGS, L"管理周期会议...\tCtrl+R");
  AppendMenuW(tools, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(tools, MF_STRING, IDM_GENERATE_DEMO_DATA, L"生成示例数据(演示)...");
  AppendMenuW(menu, MF_POPUP, (UINT_PTR)tools, L"工具");

  HMENU help = CreateMenu();
  AppendMenuW(help, MF_STRING, IDM_HELP, L"使用说明\tF1");
  AppendMenuW(menu, MF_POPUP, (UINT_PTR)help, L"帮助");
  return menu;
}

static void UpdateEditorFromListSelection(AppState* s) {
  int idx = ListView_GetNextItem(s->list, -1, LVNI_SELECTED);
  if (idx < 0 || idx >= (int)s->day.entries.size()) return;
  FillEditorFromEntry(s, s->day.entries[(size_t)idx], idx);
}

static LRESULT CALLBACK MainWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
  AppState* s = GetState(hwnd);
  switch (msg) {
    case WM_CREATE: {
      auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
      s = reinterpret_cast<AppState*>(cs->lpCreateParams);
      s->hwnd = hwnd;
      SetState(hwnd, s);

      s->font = CreateUiFont(hwnd);

      s->cal = CreateWindowExW(0, MONTHCAL_CLASSW, L"",
                               WS_CHILD | WS_VISIBLE,
                               0, 0, 10, 10, hwnd, (HMENU)IDC_CAL, nullptr, nullptr);

      s->st_day = CreateWindowExW(0, L"STATIC", L"",
                                 WS_CHILD | WS_VISIBLE,
                                 0, 0, 10, 10, hwnd, (HMENU)IDC_ST_DAY, nullptr, nullptr);

      s->list = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"",
                                WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SINGLESEL | LVS_SHOWSELALWAYS,
                                0, 0, 10, 10, hwnd, (HMENU)IDC_LIST, nullptr, nullptr);
      ListView_SetExtendedListViewStyle(s->list, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
      InitListColumns(s->list);

      s->lbl_category = CreateWindowExW(0, L"STATIC", L"分类(必填)",
                                        WS_CHILD | WS_VISIBLE,
                                        0, 0, 10, 10, hwnd, (HMENU)IDC_LBL_CATEGORY, nullptr, nullptr);
      s->lbl_start = CreateWindowExW(0, L"STATIC", L"开始(HH:MM)",
                                     WS_CHILD | WS_VISIBLE,
                                     0, 0, 10, 10, hwnd, (HMENU)IDC_LBL_START, nullptr, nullptr);
      s->lbl_end = CreateWindowExW(0, L"STATIC", L"结束(HH:MM)",
                                   WS_CHILD | WS_VISIBLE,
                                   0, 0, 10, 10, hwnd, (HMENU)IDC_LBL_END, nullptr, nullptr);
      s->lbl_status = CreateWindowExW(0, L"STATIC", L"状态",
                                      WS_CHILD | WS_VISIBLE,
                                      0, 0, 10, 10, hwnd, (HMENU)IDC_LBL_STATUS, nullptr, nullptr);

      s->cb_category = CreateWindowExW(0, L"COMBOBOX", L"",
                                       WS_CHILD | WS_VISIBLE | CBS_DROPDOWN | WS_TABSTOP | WS_VSCROLL,
                                       0, 0, 10, 10, hwnd, (HMENU)IDC_CATEGORY, nullptr, nullptr);

      s->ed_start = CreateWindowExW(0, L"COMBOBOX", L"",
                                    WS_CHILD | WS_VISIBLE | CBS_DROPDOWN | WS_TABSTOP | WS_VSCROLL,
                                    0, 0, 10, 10, hwnd, (HMENU)IDC_START, nullptr, nullptr);
      s->ed_end = CreateWindowExW(0, L"COMBOBOX", L"",
                                  WS_CHILD | WS_VISIBLE | CBS_DROPDOWN | WS_TABSTOP | WS_VSCROLL,
                                  0, 0, 10, 10, hwnd, (HMENU)IDC_END, nullptr, nullptr);
      PopulateTimeCombo(s->ed_start);
      PopulateTimeCombo(s->ed_end);

      s->cb_status = CreateWindowExW(0, L"COMBOBOX", L"",
                                     WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP,
                                     0, 0, 10, 10, hwnd, (HMENU)IDC_CB_STATUS, nullptr, nullptr);
      SendMessageW(s->cb_status, CB_ADDSTRING, 0, (LPARAM)L"无");
      SendMessageW(s->cb_status, CB_ADDSTRING, 0, (LPARAM)L"未开始");
      SendMessageW(s->cb_status, CB_ADDSTRING, 0, (LPARAM)L"进行中");
      SendMessageW(s->cb_status, CB_ADDSTRING, 0, (LPARAM)L"阻塞");
      SendMessageW(s->cb_status, CB_ADDSTRING, 0, (LPARAM)L"已完成");
      SendMessageW(s->cb_status, CB_SETCURSEL, 0, 0);
      EnableWindow(s->cb_status, FALSE);

      s->lbl_title = CreateWindowExW(0, L"STATIC", L"标题(建议填写)",
                                     WS_CHILD | WS_VISIBLE,
                                     0, 0, 10, 10, hwnd, (HMENU)IDC_LBL_TITLE, nullptr, nullptr);
      s->ed_title = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                    WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | WS_TABSTOP,
                                    0, 0, 10, 10, hwnd, (HMENU)IDC_TITLE, nullptr, nullptr);

      s->lbl_body = CreateWindowExW(0, L"STATIC", L"内容(富文本)",
                                    WS_CHILD | WS_VISIBLE,
                                    0, 0, 10, 10, hwnd, (HMENU)IDC_LBL_BODY, nullptr, nullptr);
      s->ed_body = CreateWindowExW(WS_EX_CLIENTEDGE, L"RICHEDIT50W", L"",
                                   WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_DISABLENOSCROLL | ES_NOHIDESEL |
                                       WS_VSCROLL | WS_TABSTOP,
                                   0, 0, 10, 10, hwnd, (HMENU)IDC_BODY, nullptr, nullptr);

      s->btn_new = CreateWindowExW(0, L"BUTTON", L"新增",
                                   WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                   0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_NEW, nullptr, nullptr);
      s->btn_save = CreateWindowExW(0, L"BUTTON", L"保存",
                                    WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                    0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_SAVE, nullptr, nullptr);
      s->btn_preview = CreateWindowExW(0, L"BUTTON", L"预览",
                                       WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                       0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_PREVIEW, nullptr, nullptr);
      s->btn_del = CreateWindowExW(0, L"BUTTON", L"删除",
                                   WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                   0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_DEL, nullptr, nullptr);

      s->status = CreateWindowExW(0, STATUSCLASSNAMEW, nullptr,
                                  WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
                                  0, 0, 0, 0, hwnd, (HMENU)IDC_STATUS, nullptr, nullptr);

      // Fonts
      for (HWND h : {s->cal, s->st_day, s->list, s->lbl_category, s->lbl_start, s->lbl_end, s->lbl_title, s->lbl_body,
                     s->lbl_status, s->cb_category, s->ed_start, s->ed_end, s->cb_status, s->ed_title, s->ed_body, s->status,
                     s->btn_new, s->btn_save, s->btn_preview, s->btn_del}) {
        SetControlFont(h, s->font);
      }

      SetMenu(hwnd, CreateAppMenu());

      // Load categories
      {
        std::wstring cat_err;
        if (!LoadCategories(&s->categories, &cat_err)) {
          // Non-fatal: fall back to defaults in memory.
          s->categories = {L"工作", L"会议", L"目标"};
          NormalizeCategories(&s->categories);
        }
        RefreshCategoryCombo(s);
      }

      // Load long-term tasks (non-fatal).
      {
        std::wstring terr;
        LoadTasks(&s->tasks, &terr);
      }

      // Load recurring meetings (non-fatal).
      {
        std::wstring rerr;
        LoadRecurringMeetings(&s->recurring_meetings, &rerr);
      }

      // Field hints
      SetEditCue(GetComboEditHandle(s->cb_category), L"例如: 工作 / 会议 / 目标 / 沟通 / 研发");
      SetEditCue(GetComboEditHandle(s->ed_start), L"09:30");
      SetEditCue(GetComboEditHandle(s->ed_end), L"10:15");
      SetEditCue(s->ed_title, L"一句话概括今天这件事");
      SetEditCue(s->ed_body, L"富文本正文。快捷键: Ctrl+B/I/U, Ctrl+Shift+7 编号, Ctrl+Shift+8 项目符号, Ctrl+Alt+L/E/R 段落对齐");

      SYSTEMTIME st{};
      GetLocalTime(&st);
      MonthCal_SetCurSel(s->cal, &st);
      LoadSelectedDay(s, st);
      Layout(s);
      UpdateStatusParts(s);
      UpdateStatusBar(s);
      return 0;
    }

    case WM_SIZE: {
      if (s && s->status) SendMessageW(s->status, WM_SIZE, 0, 0);
      if (s) UpdateStatusParts(s);
      if (s) Layout(s);
      if (s) UpdateStatusBar(s);
      return 0;
    }

    case WM_COMMAND: {
      int id = LOWORD(wParam);
      int code = HIWORD(wParam);

      if (id == IDC_BTN_NEW && code == BN_CLICKED) {
        if (!PromptSaveIfDirty(s)) return 0;
        ClearEditor(s);
        return 0;
      }
      if (id == IDC_BTN_SAVE && code == BN_CLICKED) {
        SaveCurrent(s);
        return 0;
      }
      if (id == IDC_BTN_PREVIEW && code == BN_CLICKED) {
        ShowPreviewWindow(s->hwnd, L"预览当前条目", BuildMarkdownForPreviewEntry(s));
        return 0;
      }
      if (id == IDC_BTN_DEL && code == BN_CLICKED) {
        int r = MessageBoxW(hwnd, L"确定删除当前条目？", kAppTitle, MB_YESNO | MB_ICONWARNING);
        if (r == IDYES) DeleteCurrent(s);
        return 0;
      }

      if ((id == IDC_CATEGORY || id == IDC_START || id == IDC_END || id == IDC_CB_STATUS || id == IDC_TITLE || id == IDC_BODY) &&
          (code == CBN_SELCHANGE || code == EN_CHANGE)) {
        SetEditorDirty(s, true);
        UpdateStatusBar(s);
        return 0;
      }

      if (id == IDM_REPORT_QUARTER) {
        DoReport(s, true, ReportMode::ByDate);
        return 0;
      }
      if (id == IDM_REPORT_YEAR) {
        DoReport(s, false, ReportMode::ByDate);
        return 0;
      }
      if (id == IDM_REPORT_QUARTER_BY_CAT) {
        DoReport(s, true, ReportMode::ByCategory);
        return 0;
      }
      if (id == IDM_REPORT_YEAR_BY_CAT) {
        DoReport(s, false, ReportMode::ByCategory);
        return 0;
      }
      if (id == IDM_EXPORT_QUARTER_CSV) {
        DoExportCsv(s, true);
        return 0;
      }
      if (id == IDM_EXPORT_YEAR_CSV) {
        DoExportCsv(s, false);
        return 0;
      }
      if (id == IDM_PREVIEW_ENTRY) {
        ShowPreviewWindow(s->hwnd, L"预览当前条目", BuildMarkdownForPreviewEntry(s));
        return 0;
      }
      if (id == IDM_PREVIEW_DAY) {
        ShowPreviewWindow(s->hwnd, L"预览当天(全部)", BuildMarkdownForPreviewDay(s));
        return 0;
      }
      if (id == IDM_FMT_BOLD) {
        if (GetFocus() == s->ed_body) {
          RichToggleCharEffect(s->ed_body, CFE_BOLD);
          SetEditorDirty(s, true);
          SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
          UpdateStatusBar(s);
        }
        return 0;
      }
      if (id == IDM_FMT_ITALIC) {
        if (GetFocus() == s->ed_body) {
          RichToggleCharEffect(s->ed_body, CFE_ITALIC);
          SetEditorDirty(s, true);
          SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
          UpdateStatusBar(s);
        }
        return 0;
      }
      if (id == IDM_FMT_UNDERLINE) {
        if (GetFocus() == s->ed_body) {
          RichToggleCharEffect(s->ed_body, CFE_UNDERLINE);
          SetEditorDirty(s, true);
          SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
          UpdateStatusBar(s);
        }
        return 0;
      }
      if (id == IDM_FMT_BULLET) {
        if (GetFocus() == s->ed_body) {
          RichToggleBullet(s->ed_body);
          SetEditorDirty(s, true);
          SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
          UpdateStatusBar(s);
        }
        return 0;
      }
      if (id == IDM_FMT_NUMBERING) {
        if (GetFocus() == s->ed_body) {
          RichToggleNumbering(s->ed_body);
          SetEditorDirty(s, true);
          SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
          UpdateStatusBar(s);
        }
        return 0;
      }
      if (id == IDM_FMT_ALIGN_LEFT) {
        if (GetFocus() == s->ed_body) {
          RichSetParaAlignment(s->ed_body, PFA_LEFT);
          SetEditorDirty(s, true);
          SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
          UpdateStatusBar(s);
        }
        return 0;
      }
      if (id == IDM_FMT_ALIGN_CENTER) {
        if (GetFocus() == s->ed_body) {
          RichSetParaAlignment(s->ed_body, PFA_CENTER);
          SetEditorDirty(s, true);
          SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
          UpdateStatusBar(s);
        }
        return 0;
      }
      if (id == IDM_FMT_ALIGN_RIGHT) {
        if (GetFocus() == s->ed_body) {
          RichSetParaAlignment(s->ed_body, PFA_RIGHT);
          SetEditorDirty(s, true);
          SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
          UpdateStatusBar(s);
        }
        return 0;
      }
      if (id == IDM_FMT_INDENT_INC) {
        if (GetFocus() == s->ed_body) {
          RichAdjustIndent(s->ed_body, 360);
          SetEditorDirty(s, true);
          SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
          UpdateStatusBar(s);
        }
        return 0;
      }
      if (id == IDM_FMT_INDENT_DEC) {
        if (GetFocus() == s->ed_body) {
          RichAdjustIndent(s->ed_body, -360);
          SetEditorDirty(s, true);
          SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
          UpdateStatusBar(s);
        }
        return 0;
      }
      if (id == IDM_FMT_CLEAR) {
        if (GetFocus() == s->ed_body) {
          RichClearFormatting(s->ed_body);
          SetEditorDirty(s, true);
          SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
          UpdateStatusBar(s);
        }
        return 0;
      }
      if (id == IDM_OPEN_DATA_DIR) {
        OpenDataDir(s);
        return 0;
      }
      if (id == IDM_MANAGE_CATEGORIES) {
        if (!PromptSaveIfDirty(s)) return 0;
        ShowManageCategoriesWindow(s);
        return 0;
      }
      if (id == IDM_MANAGE_TASKS) {
        if (!PromptSaveIfDirty(s)) return 0;
        ShowManageTasksWindow(s, nullptr);
        // Refresh current day to inject updated tasks.
        LoadSelectedDay(s, s->selected);
        return 0;
      }
      if (id == IDM_MANAGE_RECURRING_MEETINGS) {
        if (!PromptSaveIfDirty(s)) return 0;
        ShowManageRecurringMeetingsWindow(s);
        LoadSelectedDay(s, s->selected);
        return 0;
      }
      if (id == IDM_CTX_INSERT_MEETING_TEMPLATE) {
        int sel = ListView_GetNextItem(s->list, -1, LVNI_SELECTED);
        if (sel < 0 || sel >= (int)s->day.entries.size()) return 0;
        Entry& e = s->day.entries[(size_t)sel];
        if (e.type != EntryType::Meeting) return 0;
        const RecurringMeeting* rm = FindRecurringTemplateByTitle(s, e.title);
        if (!rm || rm->template_plain.empty()) return 0;

        SetEditText(s->ed_body, rm->template_plain);
        SetEditorDirty(s, true);
        UpdateStatusBar(s);
        SetFocus(s->ed_body);
        return 0;
      }
      if (id == IDM_CTX_CREATE_TASK_FROM_MEETING) {
        if (!PromptSaveIfDirty(s)) return 0;
        int sel = ListView_GetNextItem(s->list, -1, LVNI_SELECTED);
        if (sel < 0 || sel >= (int)s->day.entries.size()) return 0;
        const Entry& e = s->day.entries[(size_t)sel];
        if (e.type != EntryType::Meeting) return 0;

        Task t = BuildTaskFromMeeting(s->selected, e);
        ShowManageTasksWindow(s, &t);

        // Refresh caches and current day (task placeholders).
        {
          std::wstring terr;
          LoadTasks(&s->tasks, &terr);
        }
        LoadSelectedDay(s, s->selected);
        return 0;
      }
      if (id == IDM_GENERATE_DEMO_DATA) {
        if (!PromptSaveIfDirty(s)) return 0;
        int r = MessageBoxW(hwnd,
                            L"将追加一批演示数据(会议/长期任务/任务进展/零碎记录)。\n"
                            L"不会删除现有数据，但会写入数据目录。\n\n"
                            L"继续？",
                            kAppTitle, MB_YESNO | MB_ICONQUESTION);
        if (r != IDYES) return 0;

        SYSTEMTIME now{};
        GetLocalTime(&now);
        std::wstring derr;
        if (!GenerateDemoData(now, &derr)) {
          ShowInfoBox(hwnd, derr.c_str(), L"生成失败");
          return 0;
        }

        // Reload caches and refresh current day (so injected placeholders show up).
        {
          std::wstring cat_err;
          LoadCategories(&s->categories, &cat_err);
        }
        {
          std::wstring terr;
          LoadTasks(&s->tasks, &terr);
        }
        {
          std::wstring rerr;
          LoadRecurringMeetings(&s->recurring_meetings, &rerr);
        }
        LoadSelectedDay(s, s->selected);
        ShowInfoBox(hwnd, L"已生成示例数据。建议打开“本季度(按分类)”查看汇总效果。", L"完成");
        return 0;
      }
      if (id == IDM_HELP) {
        ShowHelp(s);
        return 0;
      }
      return 0;
    }

    case WM_CONTEXTMENU: {
      if (!s) return 0;
      HWND src = (HWND)wParam;
      if (src == s->ed_body) {
        POINT pt{};
        if (lParam == (LPARAM)-1) {
          GetCursorPos(&pt);
        } else {
          pt.x = GET_X_LPARAM(lParam);
          pt.y = GET_Y_LPARAM(lParam);
        }
        HMENU menu = CreatePopupMenu();
        if (!menu) return 0;
        AppendMenuW(menu, MF_STRING, IDM_FMT_BOLD, L"加粗\tCtrl+B");
        AppendMenuW(menu, MF_STRING, IDM_FMT_ITALIC, L"斜体\tCtrl+I");
        AppendMenuW(menu, MF_STRING, IDM_FMT_UNDERLINE, L"下划线\tCtrl+U");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, IDM_FMT_NUMBERING, L"编号\tCtrl+Shift+7");
        AppendMenuW(menu, MF_STRING, IDM_FMT_BULLET, L"项目符号\tCtrl+Shift+8");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, IDM_FMT_INDENT_INC, L"增加缩进\tCtrl+Alt+>");
        AppendMenuW(menu, MF_STRING, IDM_FMT_INDENT_DEC, L"减少缩进\tCtrl+Alt+<");
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        AppendMenuW(menu, MF_STRING, IDM_FMT_CLEAR, L"清除格式");
        UINT cmd = TrackPopupMenu(menu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, 0, hwnd, nullptr);
        DestroyMenu(menu);
        if (cmd) SendMessageW(hwnd, WM_COMMAND, cmd, 0);
        return 0;
      }

      if (src != s->list) return DefWindowProcW(hwnd, msg, wParam, lParam);

      if (!PromptSaveIfDirty(s)) return 0;

      POINT pt{};
      if (lParam == (LPARAM)-1) {
        GetCursorPos(&pt);
      } else {
        pt.x = GET_X_LPARAM(lParam);
        pt.y = GET_Y_LPARAM(lParam);
      }

      // Hit test and select item under cursor, so context actions apply to what user right-clicked.
      POINT pt_client = pt;
      ScreenToClient(s->list, &pt_client);
      LVHITTESTINFO ht{};
      ht.pt = pt_client;
      int hit = ListView_HitTest(s->list, &ht);
      if (hit >= 0) {
        ListView_SetItemState(s->list, hit, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
        ListView_EnsureVisible(s->list, hit, FALSE);
        UpdateEditorFromListSelection(s);
      }

      int sel = ListView_GetNextItem(s->list, -1, LVNI_SELECTED);
      if (sel < 0 || sel >= (int)s->day.entries.size()) return 0;
      const Entry& e = s->day.entries[(size_t)sel];

      HMENU menu = CreatePopupMenu();
      if (!menu) return 0;

      if (e.type == EntryType::Meeting) {
        AppendMenuW(menu, MF_STRING, IDM_CTX_CREATE_TASK_FROM_MEETING, L"从此会议创建长期任务...");

        const RecurringMeeting* rm = FindRecurringTemplateByTitle(s, e.title);
        if (rm && !rm->template_plain.empty()) {
          if (e.body_plain.empty()) {
            AppendMenuW(menu, MF_STRING, IDM_CTX_INSERT_MEETING_TEMPLATE, L"插入周期会议模板");
          }
        }
      }

      if (GetMenuItemCount(menu) == 0) {
        DestroyMenu(menu);
        return 0;
      }

      UINT cmd = TrackPopupMenu(menu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, 0, hwnd, nullptr);
      DestroyMenu(menu);
      if (cmd) SendMessageW(hwnd, WM_COMMAND, cmd, 0);
      return 0;
    }

    case WM_NOTIFY: {
      NMHDR* hdr = reinterpret_cast<NMHDR*>(lParam);
      if (!s || !hdr) return 0;

      if (hdr->idFrom == IDC_CAL && hdr->code == MCN_SELCHANGE) {
        if (!PromptSaveIfDirty(s)) {
          MonthCal_SetCurSel(s->cal, &s->selected);
          return 0;
        }
        auto* n = reinterpret_cast<NMSELCHANGE*>(lParam);
        LoadSelectedDay(s, n->stSelStart);
        return 0;
      }

      if (hdr->idFrom == IDC_LIST) {
        if (hdr->code == NM_CUSTOMDRAW) {
          auto* cd = reinterpret_cast<NMLVCUSTOMDRAW*>(lParam);
          if (cd->nmcd.dwDrawStage == CDDS_PREPAINT) return CDRF_NOTIFYITEMDRAW;
          if (cd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
            int idx = (int)cd->nmcd.dwItemSpec;
            if (idx >= 0 && idx < (int)s->day.entries.size()) {
              const auto& e = s->day.entries[(size_t)idx];
              EntryStatus st = EffectiveStatus(s, e);
              COLORREF bk = StatusBkColor(st);
              if (bk != CLR_NONE) {
                cd->clrTextBk = bk;
                cd->clrText = RGB(20, 20, 20);
              }
            }
            return CDRF_DODEFAULT;
          }
          return CDRF_DODEFAULT;
        }
        if (hdr->code == LVN_ITEMCHANGED) {
          auto* p = reinterpret_cast<NMLISTVIEW*>(lParam);
          if ((p->uChanged & LVIF_STATE) && (p->uNewState & LVIS_SELECTED)) {
            if (!PromptSaveIfDirty(s)) {
              // Revert selection change.
              ListView_SetItemState(s->list, p->iItem, 0, LVIS_SELECTED);
              return 0;
            }
            UpdateEditorFromListSelection(s);
          }
          return 0;
        }
        if (hdr->code == NM_DBLCLK) {
          if (!PromptSaveIfDirty(s)) return 0;
          auto* act = reinterpret_cast<NMITEMACTIVATE*>(lParam);
          if (act && act->iItem < 0) {
            // Double click on empty space: start a new entry (common expectation).
            ClearEditor(s);
            HWND h = GetComboEditHandle(s->cb_category);
            SetFocus(h ? h : s->cb_category);
            return 0;
          }
          UpdateEditorFromListSelection(s);
          SetFocus(s->ed_body);
          return 0;
        }
      }
      return 0;
    }

    case WM_CLOSE: {
      if (s && !PromptSaveIfDirty(s)) return 0;
      DestroyWindow(hwnd);
      return 0;
    }

    case WM_DESTROY: {
      if (s && s->font) DeleteObject(s->font);
      PostQuitMessage(0);
      return 0;
    }
  }
  return DefWindowProcW(hwnd, msg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int nCmdShow) {
  INITCOMMONCONTROLSEX icc{};
  icc.dwSize = sizeof(icc);
  icc.dwICC = ICC_DATE_CLASSES | ICC_LISTVIEW_CLASSES | ICC_STANDARD_CLASSES | ICC_BAR_CLASSES;
  InitCommonControlsEx(&icc);

  // RichEdit (msftedit) is a Windows component, loaded on demand.
  LoadLibraryW(L"Msftedit.dll");

  AppState state{};

  WNDCLASSW wc{};
  wc.lpfnWndProc = MainWndProc;
  wc.hInstance = hInstance;
  wc.lpszClassName = L"WorkLogLiteMainWnd";
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);

  if (!RegisterClassW(&wc)) return 0;

  HWND hwnd = CreateWindowExW(0, wc.lpszClassName, kAppTitle,
                              WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                              CW_USEDEFAULT, CW_USEDEFAULT, 1100, 720,
                              nullptr, nullptr, hInstance, &state);
  if (!hwnd) return 0;

  ShowWindow(hwnd, nCmdShow);
  UpdateWindow(hwnd);

  MSG msg{};
    while (GetMessageW(&msg, nullptr, 0, 0) > 0) {
      if (msg.message == WM_KEYDOWN) {
        bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
        bool shift = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        bool alt = (GetKeyState(VK_MENU) & 0x8000) != 0;
        HWND focus = GetFocus();

        if (ctrl && !alt) {
          if (msg.wParam == 'S') {
            SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_SAVE, BN_CLICKED), 0);
            continue;
          }
          if (msg.wParam == 'N') {
            SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_NEW, BN_CLICKED), 0);
            continue;
          }
          if (msg.wParam == 'Q') {
            SendMessageW(hwnd, WM_COMMAND, IDM_REPORT_QUARTER, 0);
            continue;
          }
          if (msg.wParam == 'Y') {
            SendMessageW(hwnd, WM_COMMAND, IDM_REPORT_YEAR, 0);
            continue;
          }
          if (msg.wParam == 'K') {
            SendMessageW(hwnd, WM_COMMAND, IDM_MANAGE_CATEGORIES, 0);
            continue;
          }
          if (msg.wParam == 'P') {
            SendMessageW(hwnd, WM_COMMAND, shift ? IDM_PREVIEW_DAY : IDM_PREVIEW_ENTRY, 0);
            continue;
          }
          if (msg.wParam == 'T') {
            SendMessageW(hwnd, WM_COMMAND, IDM_MANAGE_TASKS, 0);
            continue;
          }
          if (msg.wParam == 'R') {
            SendMessageW(hwnd, WM_COMMAND, IDM_MANAGE_RECURRING_MEETINGS, 0);
            continue;
          }
          if (focus == state.ed_body) {
            if (msg.wParam == 'B') {
              RichToggleCharEffect(state.ed_body, CFE_BOLD);
              SetEditorDirty(&state, true);
              SendMessageW(state.ed_body, EM_SETMODIFY, TRUE, 0);
              UpdateStatusBar(&state);
              continue;
            }
            if (msg.wParam == 'I') {
              RichToggleCharEffect(state.ed_body, CFE_ITALIC);
              SetEditorDirty(&state, true);
              SendMessageW(state.ed_body, EM_SETMODIFY, TRUE, 0);
              UpdateStatusBar(&state);
              continue;
            }
            if (msg.wParam == 'U') {
              RichToggleCharEffect(state.ed_body, CFE_UNDERLINE);
              SetEditorDirty(&state, true);
              SendMessageW(state.ed_body, EM_SETMODIFY, TRUE, 0);
              UpdateStatusBar(&state);
              continue;
            }
            if (shift && msg.wParam == '7') {
              RichToggleNumbering(state.ed_body);
              SetEditorDirty(&state, true);
              SendMessageW(state.ed_body, EM_SETMODIFY, TRUE, 0);
              UpdateStatusBar(&state);
              continue;
            }
            if (shift && msg.wParam == '8') {
              RichToggleBullet(state.ed_body);
              SetEditorDirty(&state, true);
              SendMessageW(state.ed_body, EM_SETMODIFY, TRUE, 0);
              UpdateStatusBar(&state);
              continue;
            }
            // Keep legacy shortcut for bullets.
            if (msg.wParam == 'L') {
              RichToggleBullet(state.ed_body);
              SetEditorDirty(&state, true);
              SendMessageW(state.ed_body, EM_SETMODIFY, TRUE, 0);
              UpdateStatusBar(&state);
              continue;
            }
          }
        }

        if (ctrl && alt && focus == state.ed_body) {
          if (msg.wParam == 'L') {
            RichSetParaAlignment(state.ed_body, PFA_LEFT);
            SetEditorDirty(&state, true);
            SendMessageW(state.ed_body, EM_SETMODIFY, TRUE, 0);
            UpdateStatusBar(&state);
            continue;
          }
          if (msg.wParam == 'E') {
            RichSetParaAlignment(state.ed_body, PFA_CENTER);
            SetEditorDirty(&state, true);
            SendMessageW(state.ed_body, EM_SETMODIFY, TRUE, 0);
            UpdateStatusBar(&state);
            continue;
          }
          if (msg.wParam == 'R') {
            RichSetParaAlignment(state.ed_body, PFA_RIGHT);
            SetEditorDirty(&state, true);
            SendMessageW(state.ed_body, EM_SETMODIFY, TRUE, 0);
            UpdateStatusBar(&state);
            continue;
          }
          if (msg.wParam == VK_OEM_PERIOD) {  // '>'
            RichAdjustIndent(state.ed_body, 360);
            SetEditorDirty(&state, true);
            SendMessageW(state.ed_body, EM_SETMODIFY, TRUE, 0);
            UpdateStatusBar(&state);
            continue;
          }
          if (msg.wParam == VK_OEM_COMMA) {  // '<'
            RichAdjustIndent(state.ed_body, -360);
            SetEditorDirty(&state, true);
            SendMessageW(state.ed_body, EM_SETMODIFY, TRUE, 0);
            UpdateStatusBar(&state);
            continue;
          }
        }

        if (msg.wParam == VK_DELETE) {
          SendMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_BTN_DEL, BN_CLICKED), 0);
          continue;
        }
        if (msg.wParam == VK_F1) {
          SendMessageW(hwnd, WM_COMMAND, IDM_HELP, 0);
          continue;
        }
      }
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
  return 0;
}

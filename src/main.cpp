#include "categories.h"
#include "audio_record.h"
#include "app_resource.h"
#include "demo.h"
#include "color_picker.h"
#include "data_ops.h"
#include "report.h"
#include "storage.h"
#include "tasks.h"
#include "types.h"
#include "win32_util.h"
#include "screenshot_tool.h"

#include <commctrl.h>
#include <commdlg.h>
#include <richedit.h>
#include <shobjidl.h>
#include <shellapi.h>
#include <windowsx.h>
#include <bcrypt.h>

#include <string>
#include <vector>
#include <algorithm>
#include <cmath>
#include <cwctype>
#include <cstring>
#include <map>
#include <sstream>
#include <set>
#include <string_view>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "comdlg32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "bcrypt.lib")

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
  IDM_EXPORT_FULL_DATA = 1012,
  IDM_IMPORT_FULL_DATA = 1013,
  IDM_CLEAR_ALL_DATA = 1014,
  IDM_THEME_MUNG = 1015,
  IDM_PREVIEW_ENTRY = 1020,
  IDM_PREVIEW_DAY = 1021,
  IDM_MANAGE_TASKS = 1030,
  IDM_GENERATE_DEMO_DATA = 1050,
  IDM_VIEW_TASK_PROGRESS = 1062,
  IDM_HELP = 1100,

  // Context menu (ListView)
  IDM_CTX_CREATE_TASK_FROM_MEETING = 20001,
  IDM_CTX_VIEW_TASK_PROGRESS = 20003,
  IDM_CTX_OPEN_TASK_MATERIALS_DIR = 20004,
  IDM_CTX_OPEN_MATERIALS_DIR = 20005,
  IDM_CTX_SET_MATERIALS_DIR = 20006,
  IDM_CTX_CLEAR_MATERIALS_DIR = 20007,

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
  IDC_GRP_OFFICE = 2060,
  IDC_BTN_SCREENSHOT = 2061,
  IDC_BTN_QUICK_REPLY = 2062,
  IDC_BTN_TIMESTAMP = 2063,
  IDC_BTN_FOCUS_TIMER = 2064,
  IDC_BTN_CALC = 2065,
  IDC_BTN_DATA_DIR = 2066,
  IDC_BTN_SCREENSHOT_DIR = 2067,
  IDC_BTN_TASK_MATERIALS_ROOT = 2068,
  IDC_BTN_OPEN_MATERIALS = 2069,
  IDC_BTN_SET_MATERIALS = 2070,
  IDC_BTN_RECORD_SYSTEM = 2071,
  IDC_BTN_RECORD_MIC = 2072,
  IDC_BTN_COLOR_PICKER = 2073,
  IDC_BTN_PASTE_PLAIN = 2074,
  IDC_BTN_DATE_SEPARATOR = 2075,
  IDC_BTN_RECORDINGS_DIR = 2076,
  IDC_STATUS = 2020,
  IDC_CB_STATUS = 2021,
  IDC_LBL_CATEGORY = 2030,
  IDC_LBL_START = 2031,
  IDC_LBL_END = 2032,
  IDC_LBL_STATUS = 2033,
  IDC_LBL_TITLE = 2034,
  IDC_LBL_BODY = 2035,

  // Body formatting toolbar buttons
  IDC_FMT_BOLD = 2040,
  IDC_FMT_ITALIC = 2041,
  IDC_FMT_UNDERLINE = 2042,
  IDC_FMT_NUMBERING = 2043,
  IDC_FMT_BULLET = 2044,
  IDC_FMT_ALIGN_LEFT = 2045,
  IDC_FMT_ALIGN_CENTER = 2046,
  IDC_FMT_ALIGN_RIGHT = 2047,
  IDC_FMT_INDENT_INC = 2048,
  IDC_FMT_INDENT_DEC = 2049,
  IDC_FMT_CLEAR = 2050,
  IDC_FMT_SUPERSCRIPT = 2052,
  IDC_FMT_SUBSCRIPT = 2053,
  IDC_FMT_MARKDOWN = 2054,
  IDC_BTN_BODY_MAXIMIZE = 2051,
};

struct AppState {
  HWND hwnd{};
  HWND cal{};
  HWND st_day{};
  HWND list{};
  HWND grp_office{};
  HWND btn_screenshot{};
  HWND btn_quick_reply{};
  HWND btn_timestamp{};
  HWND btn_focus_timer{};
  HWND btn_calc{};
  HWND btn_data_dir{};
  HWND btn_screenshot_dir{};
  HWND btn_task_materials{};
  HWND btn_open_materials{};
  HWND btn_set_materials{};
  HWND btn_record_system{};
  HWND btn_record_mic{};
  HWND btn_recordings_dir{};
  HWND btn_color_picker{};
  HWND btn_paste_plain{};
  HWND btn_date_separator{};
  HWND lbl_category{};
  HWND lbl_start{};
  HWND lbl_end{};
  HWND lbl_status{};
  HWND lbl_title{};
  HWND lbl_body{};
  HWND fmt_bold{};
  HWND fmt_italic{};
  HWND fmt_underline{};
  HWND fmt_numbering{};
  HWND fmt_bullet{};
  HWND fmt_align_left{};
  HWND fmt_align_center{};
  HWND fmt_align_right{};
  HWND fmt_indent_inc{};
  HWND fmt_indent_dec{};
  HWND fmt_clear{};
  HWND fmt_markdown{};
  HWND fmt_superscript{};
  HWND fmt_subscript{};
  HWND btn_body_maximize{};
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
  bool mung_theme{false};
  HBRUSH brush_window{};
  HBRUSH brush_panel{};
  HBRUSH brush_edit{};

  SYSTEMTIME selected{};
  DayData day{};
  int editing_index{-1};  // -1 => new entry
  bool editor_dirty{false};
  bool editor_maximized{false};
  int body_zoom_percent{100};
  bool ignore_cal_selchange{false};  // avoid repeated prompts when we programmatically revert calendar selection
  bool suppress_dirty_prompt{false};  // avoid re-entrant prompts triggered by programmatic UI updates while saving
  bool suppress_editor_change_tracking{false};  // ignore control notifications during programmatic editor fills
  bool suppress_list_selection_sync{false};  // ignore list selection notifications during programmatic rebuild/restore
  bool close_in_progress{false};

  // Small "toast" messages shown in the status bar (e.g. "copied to clipboard").
  std::wstring toast;
  ULONGLONG toast_until{0};

  // Focus timer (pomodoro-like, offline).
  bool focus_running{false};
  ULONGLONG focus_end_tick{0};
  UINT_PTR focus_timer_id{0};
  UINT_PTR audio_timer_id{0};
  UINT_PTR autosave_timer_id{0};
  AudioRecorder* system_recorder{};
  AudioRecorder* mic_recorder{};
  ULONGLONG system_record_start_tick{0};
  ULONGLONG mic_record_start_tick{0};

  std::vector<std::wstring> categories;
  std::vector<Task> tasks;
  int list_sort_column{0};  // 0=start/date range, 1=category, 3=status
  bool list_sort_desc{false};

  // Optional materials directory selected for the current editor content (used for new entries too).
  std::wstring editor_materials_dir;
};

struct EditorSnapshot {
  int editing_index{-1};
  std::wstring category;
  std::wstring start_time;
  std::wstring end_time;
  int status_sel{0};
  std::wstring title;
  std::wstring body_plain;
  std::string body_rtf_b64;
  std::wstring materials_dir;
};

// Forward declarations used by small UI helpers defined early in this file.
static std::wstring NormalizeNewlinesToCrlf(const std::wstring& s);
static bool CopyToClipboard(HWND hwnd, const std::wstring& text);
static bool WriteUtf8File(const std::wstring& path, const std::wstring& content, std::wstring* err);
static bool ReadFileUtf8Simple(const std::wstring& path, std::wstring* out, std::wstring* err);
static bool EditorMatchesDefaultDraft(AppState* s);
static int CompareDateYmd(const std::wstring& a, const std::wstring& b);
static bool EntryVisibleOnDate(const Entry& e, const SYSTEMTIME& selected, const SYSTEMTIME* source_date = nullptr);
static bool LoadVisibleEntriesForDate(const SYSTEMTIME& selected, DayData* out, std::wstring* err);
static bool SaveEntryToBackingStore(const Entry& e, const SYSTEMTIME& fallback_date, std::wstring* err);
static bool DeleteEntriesFromBackingStore(const std::vector<Entry>& entries, const SYSTEMTIME& fallback_date, std::wstring* err);
static std::vector<int> GetSelectedEntryIndices(AppState* s);
static void RenderMarkdownToRichEdit(HWND rich, HFONT font, const std::wstring& md);
static void OpenDataDir(AppState* s);
static std::wstring ParentDirOf(const std::wstring& path);
static std::wstring JoinPathLoose(const std::wstring& a, const std::wstring& b);
static std::wstring NormalizeDirForCompare(const std::wstring& path);
static void ApplyTheme(AppState* s);
static COLORREF StatusBkColor(EntryStatus st);
static std::wstring FormatWin32ErrorMessage(DWORD err);
static bool ReadClipboardText(HWND hwnd, std::wstring* out);
static std::string RichEditGetRtfBytes(HWND rich);
static void RichEditSetRtfBytes(HWND rich, const std::string& bytes);
static std::wstring GetTaskMaterialsDirById(const std::wstring& task_id);
static std::wstring TaskMaterialsDirResolved(const Task& t);
static std::wstring GetWorkMaterialsDir(const SYSTEMTIME& date, const std::wstring& entry_id);
static bool IsNetworkPathForbidden(const std::wstring& path);
static bool OpenPathInShell(HWND owner, const std::wstring& path, const wchar_t* what);
static int CompareSystemTimeDateOnly(const SYSTEMTIME& a, const SYSTEMTIME& b);
static void TaskInitListColumns(HWND list);
static void ShowUnitCalcWindow(HWND owner);
static void Layout(AppState* s);

static bool IsEditorTrulyDirty(AppState* s);
static void ClearEditor(AppState* s, bool allow_blank = false);
static void SetEditorDirty(AppState* s, bool dirty);
static void FillEditorFromEntry(AppState* s, const Entry& e, int index);
static void ApplyBodyZoom(AppState* s);
static bool IsAudioRecording(AppState* s, AudioRecordSource source);
static std::wstring FormatElapsedClock(ULONGLONG start_tick, ULONGLONG now);
static std::wstring GetDateCtrlYmd(HWND hwnd);
static void SetDateCtrlYmd(HWND hwnd, const std::wstring& ymd);
static void SetDateCtrlOrDefault(HWND hwnd, const std::wstring& ymd, const SYSTEMTIME& fallback);

static void OpenMaterialsDirForEntry(HWND owner, AppState* s, const Entry& e);
static void EnsureMaterialsDirForEntry(AppState* s, const Entry& e);
static EditorSnapshot CaptureEditorSnapshot(AppState* s);
static void RestoreEditorSnapshot(AppState* s, const EditorSnapshot& snap, bool mark_clean);
static void UpdateStatusBar(AppState* s);
static bool AutoSaveCurrent(AppState* s);

static UINT QueryDpiForWindowCompat(HWND hwnd) {
  HMODULE user32 = GetModuleHandleW(L"user32.dll");
  if (!user32) return 96;
  using Fn = UINT(WINAPI*)(HWND);
  auto p = reinterpret_cast<Fn>(GetProcAddress(user32, "GetDpiForWindow"));
  if (!p) return 96;
  return p(hwnd);
}

static UINT QuerySystemDpiCompat() {
  // Prefer Win10's GetDpiForSystem; fallback to primary monitor LOGPIXELSX.
  HMODULE user32 = GetModuleHandleW(L"user32.dll");
  if (user32) {
    using Fn = UINT(WINAPI*)();
    auto p = reinterpret_cast<Fn>(GetProcAddress(user32, "GetDpiForSystem"));
    if (p) return p();
  }
  HDC hdc = GetDC(nullptr);
  int dpi = hdc ? GetDeviceCaps(hdc, LOGPIXELSX) : 96;
  if (hdc) ReleaseDC(nullptr, hdc);
  if (dpi <= 0) dpi = 96;
  return (UINT)dpi;
}

static void AdjustWindowRectExForDpiCompat(RECT* rc, DWORD style, BOOL has_menu, DWORD ex_style, UINT dpi) {
  if (!rc) return;
  HMODULE user32 = GetModuleHandleW(L"user32.dll");
  if (user32) {
    using Fn = BOOL(WINAPI*)(LPRECT, DWORD, BOOL, DWORD, UINT);
    auto p = reinterpret_cast<Fn>(GetProcAddress(user32, "AdjustWindowRectExForDpi"));
    if (p) {
      (void)p(rc, style, has_menu, ex_style, dpi);
      return;
    }
  }
  (void)AdjustWindowRectEx(rc, style, has_menu, ex_style);
}

static void ComputeMainMinClientSize(HWND hwnd, AppState* s, int* out_w, int* out_h) {
  if (out_w) *out_w = 900;
  if (out_h) *out_h = 640;
  if (!hwnd) return;

  int pad = ScalePx(hwnd, 10);
  int left_w = ScalePx(hwnd, 260);
  int top_h = ScalePx(hwnd, 240);
  int row_h = ScalePx(hwnd, 26);
  int lbl_h = ScalePx(hwnd, 18);
  int btn_w = ScalePx(hwnd, 72);

  // Right-side row fields widths (category/start/end/status) + a gap before right-aligned buttons.
  int w_fields = ScalePx(hwnd, 140) + pad + ScalePx(hwnd, 70) + pad + ScalePx(hwnd, 70) + pad + ScalePx(hwnd, 90);
  int w_buttons = 4 * btn_w + 3 * pad;
  int right_min_w = w_fields + pad + w_buttons;

  // Formatting toolbar width (11 buttons).
  int tbw = ScalePx(hwnd, 30);
  int tb_pad = ScalePx(hwnd, 6);
  int w_toolbar = 11 * tbw + 10 * tb_pad;
  if (right_min_w < w_toolbar) right_min_w = w_toolbar;

  int min_client_w = pad + left_w + pad + right_min_w + pad;

  // Height: ensure office buttons fit and body editor has a minimum height.
  int status_h = ScalePx(hwnd, 22);
  if (s && s->status) {
    RECT rs{};
    if (GetWindowRect(s->status, &rs)) status_h = (rs.bottom - rs.top);
  }

  // Compute y (layout's "y" right before ed_body) to keep logic consistent with Layout().
  int y = pad;
  y += lbl_h + ScalePx(hwnd, 4);
  y += top_h + pad;
  y += lbl_h + ScalePx(hwnd, 2);
  y += row_h + pad;
  y += lbl_h + ScalePx(hwnd, 2);
  y += row_h + pad;
  y += lbl_h + ScalePx(hwnd, 2);
  int tbh = ScalePx(hwnd, 26);
  y += tbh + ScalePx(hwnd, 4);

  int min_body_h = ScalePx(hwnd, 160);
  int h_no_status_need = y + pad + min_body_h;

  // Left office panel: number of buttons determines required rows.
  int office_btn_count = 15;  // conservative fallback (current shipped count)
  if (s) {
    office_btn_count = 0;
    for (HWND b : {s->btn_screenshot, s->btn_quick_reply, s->btn_timestamp, s->btn_focus_timer, s->btn_calc,
                   s->btn_data_dir, s->btn_screenshot_dir, s->btn_color_picker, s->btn_paste_plain,
                   s->btn_record_system, s->btn_record_mic, s->btn_task_materials, s->btn_open_materials,
                   s->btn_set_materials, s->btn_date_separator}) {
      if (b) office_btn_count++;
    }
    if (office_btn_count <= 0) office_btn_count = 15;
  }
  int cols = 2;
  int rows = (office_btn_count + cols - 1) / cols;
  if (rows < 1) rows = 1;
  int bh = ScalePx(hwnd, 32);
  int gap = ScalePx(hwnd, 8);
  int office_header = ScalePx(hwnd, 26);
  int office_need = office_header + rows * bh + (rows - 1) * gap + ScalePx(hwnd, 8);
  int office_y = pad + top_h + pad;
  int h_no_status_need2 = office_y + office_need + pad;
  if (h_no_status_need < h_no_status_need2) h_no_status_need = h_no_status_need2;

  int min_client_h = h_no_status_need + status_h;

  if (out_w) *out_w = min_client_w;
  if (out_h) *out_h = min_client_h;
}

static void ApplyMinTrackSize(HWND hwnd, AppState* s, MINMAXINFO* mmi) {
  if (!hwnd || !mmi) return;

  int min_client_w = 0, min_client_h = 0;
  ComputeMainMinClientSize(hwnd, s, &min_client_w, &min_client_h);

  RECT rc{0, 0, min_client_w, min_client_h};
  DWORD style = (DWORD)GetWindowLongPtrW(hwnd, GWL_STYLE);
  DWORD ex_style = (DWORD)GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
  BOOL has_menu = GetMenu(hwnd) != nullptr;
  UINT dpi = QueryDpiForWindowCompat(hwnd);
  AdjustWindowRectExForDpiCompat(&rc, style, has_menu, ex_style, dpi);

  int min_w = rc.right - rc.left;
  int min_h = rc.bottom - rc.top;
  if (min_w < 400) min_w = 400;
  if (min_h < 300) min_h = 300;

  // Avoid an impossible minimum size on small screens: cap to current work area.
  RECT wa{};
  if (SystemParametersInfoW(SPI_GETWORKAREA, 0, &wa, 0)) {
    int ww = wa.right - wa.left;
    int wh = wa.bottom - wa.top;
    if (ww > 0 && min_w > ww) min_w = ww;
    if (wh > 0 && min_h > wh) min_h = wh;
  }
  mmi->ptMinTrackSize.x = min_w;
  mmi->ptMinTrackSize.y = min_h;
}

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

static EditorSnapshot CaptureEditorSnapshot(AppState* s) {
  EditorSnapshot snap{};
  if (!s) return snap;
  snap.editing_index = s->editing_index;
  snap.category = GetWindowTextWString(s->cb_category);
  snap.start_time = GetDateCtrlYmd(s->ed_start);
  snap.end_time = GetDateCtrlYmd(s->ed_end);
  snap.status_sel = s->cb_status ? (int)SendMessageW(s->cb_status, CB_GETCURSEL, 0, 0) : 0;
  snap.title = GetWindowTextWString(s->ed_title);
  snap.body_plain = GetWindowTextWString(s->ed_body);
  if (s->ed_body && !snap.body_plain.empty()) {
    std::string rtf = RichEditGetRtfBytes(s->ed_body);
    if (!rtf.empty()) snap.body_rtf_b64 = Base64Encode(rtf);
  }
  snap.materials_dir = s->editor_materials_dir;
  return snap;
}

static void RestoreEditorSnapshot(AppState* s, const EditorSnapshot& snap, bool mark_clean) {
  if (!s) return;
  s->suppress_editor_change_tracking = true;
  SetWindowTextW(s->cb_category, snap.category.c_str());
  SetDateCtrlOrDefault(s->ed_start, snap.start_time, s->selected);
  SetDateCtrlOrDefault(s->ed_end, snap.end_time, AddDays(s->selected, 1));
  if (s->cb_status) {
    SendMessageW(s->cb_status, CB_SETCURSEL, snap.status_sel, 0);
    EnableWindow(s->cb_status, TRUE);
  }
  SetEditText(s->ed_title, snap.title);
  if (!snap.body_rtf_b64.empty()) {
    std::string rtf;
    if (Base64Decode(snap.body_rtf_b64, &rtf)) {
      RichEditSetRtfBytes(s->ed_body, rtf);
    } else {
      SetEditText(s->ed_body, snap.body_plain);
    }
  } else {
    SetEditText(s->ed_body, snap.body_plain);
  }
  ApplyBodyZoom(s);
  s->editor_materials_dir = snap.materials_dir;
  s->editing_index = snap.editing_index;
  if (mark_clean) {
    if (s->ed_title) SendMessageW(s->ed_title, EM_SETMODIFY, FALSE, 0);
    if (s->ed_body) SendMessageW(s->ed_body, EM_SETMODIFY, FALSE, 0);
    SetEditorDirty(s, false);
  }
  s->suppress_editor_change_tracking = false;
  UpdateStatusBar(s);
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

static void RichToggleScriptEffect(HWND rich, DWORD effect_flag, DWORD other_flag) {
  if (!rich) return;
  CHARFORMAT2W cf{};
  cf.cbSize = sizeof(cf);
  SendMessageW(rich, EM_GETCHARFORMAT, (WPARAM)SCF_SELECTION, (LPARAM)&cf);

  bool enabled = (cf.dwEffects & effect_flag) != 0;
  CHARFORMAT2W out{};
  out.cbSize = sizeof(out);
  out.dwMask = effect_flag | other_flag;
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

static void RichToggleNumberingStyle(HWND rich, WORD numbering_kind, WORD numbering_style) {
  if (!rich) return;
  PARAFORMAT2 pf{};
  pf.cbSize = sizeof(pf);
  SendMessageW(rich, EM_GETPARAFORMAT, 0, (LPARAM)&pf);
  bool is_same_numbered = (pf.wNumbering == numbering_kind && pf.wNumberingStyle == numbering_style);

  PARAFORMAT2 out{};
  out.cbSize = sizeof(out);
  out.dwMask = PFM_NUMBERING | PFM_NUMBERINGSTYLE | PFM_NUMBERINGSTART | PFM_NUMBERINGTAB | PFM_STARTINDENT | PFM_OFFSET;
  if (is_same_numbered) {
    out.wNumbering = 0;
    out.wNumberingStyle = 0;
    out.wNumberingStart = 0;
    out.wNumberingTab = 0;
    out.dxStartIndent = 0;
    out.dxOffset = 0;
  } else {
    // Keep numbering compact; avoid shifting paragraph text too far right.
    out.wNumbering = PFN_ARABIC;
    out.wNumberingStyle = PFNS_PERIOD;
    out.wNumberingStart = 1;
    out.wNumberingTab = 0;
    out.dxStartIndent = 0;
    out.dxOffset = 0;
  }
  SendMessageW(rich, EM_SETPARAFORMAT, 0, (LPARAM)&out);
}

static void RichToggleNumbering(HWND rich) {
  RichToggleNumberingStyle(rich, PFN_ARABIC, PFNS_PERIOD);
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

static std::wstring GetDateCtrlYmd(HWND hwnd) {
  if (!hwnd) return L"";
  SYSTEMTIME st{};
  if (DateTime_GetSystemtime(hwnd, &st) != GDT_VALID) return L"";
  st.wHour = 0;
  st.wMinute = 0;
  st.wSecond = 0;
  st.wMilliseconds = 0;
  return FormatDateYYYYMMDD(st);
}

static void SetDateCtrlYmd(HWND hwnd, const std::wstring& ymd) {
  if (!hwnd) return;
  if (ymd.empty()) {
    DateTime_SetSystemtime(hwnd, GDT_NONE, nullptr);
    return;
  }
  SYSTEMTIME st{};
  if (ParseYYYYMMDD(ymd, &st)) {
    st.wHour = 0;
    st.wMinute = 0;
    st.wSecond = 0;
    st.wMilliseconds = 0;
    DateTime_SetSystemtime(hwnd, GDT_VALID, &st);
  } else {
    DateTime_SetSystemtime(hwnd, GDT_NONE, nullptr);
  }
}

static void SetDateCtrlOrDefault(HWND hwnd, const std::wstring& ymd, const SYSTEMTIME& fallback) {
  if (!ymd.empty()) {
    SetDateCtrlYmd(hwnd, ymd);
    return;
  }
  SYSTEMTIME st = fallback;
  st.wHour = 0;
  st.wMinute = 0;
  st.wSecond = 0;
  st.wMilliseconds = 0;
  DateTime_SetSystemtime(hwnd, GDT_VALID, &st);
}

static void UpdateStatusBar(AppState* s) {
  if (!s->status) return;
  // Part 2 is a static hint; part 0 is dynamic info; part 1 shows editor stats.
  std::wstring hint = L"快捷键: Ctrl+S 保存  Ctrl+N 新增  Ctrl+P 预览  Ctrl+K 分类  F1 帮助";
  SendMessageW(s->status, SB_SETTEXTW, 2, (LPARAM)hint.c_str());

  std::wstring body = s->ed_body ? GetWindowTextWString(s->ed_body) : L"";
  wchar_t stats[128]{};
  wsprintfW(stats, L"字数 %d  缩放 %d%%  状态 %s",
            (int)body.size(),
            s->body_zoom_percent,
            IsEditorTrulyDirty(s) ? L"未保存" : L"已保存");
  SendMessageW(s->status, SB_SETTEXTW, 1, (LPARAM)stats);

  ULONGLONG now = GetTickCount64();
  if (now < s->toast_until && !s->toast.empty()) {
    SendMessageW(s->status, SB_SETTEXTW, 0, (LPARAM)s->toast.c_str());
    return;
  }

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
  if (!s->editor_materials_dir.empty()) info += L"  材料路径已设置";

  if (s->focus_running) {
    ULONGLONG remain = (s->focus_end_tick > now) ? (s->focus_end_tick - now) : 0;
    int sec = (int)(remain / 1000);
    int mm = sec / 60;
    int ss = sec % 60;
    wchar_t buf[64]{};
    wsprintfW(buf, L"  专注 %02d:%02d", mm, ss);
    info += buf;
  }
  if (IsAudioRecording(s, AudioRecordSource::SystemLoopback)) {
    info += L"  系统录音 ";
    info += FormatElapsedClock(s->system_record_start_tick, now);
  }
  if (IsAudioRecording(s, AudioRecordSource::MicrophoneCapture)) {
    info += L"  麦克录音 ";
    info += FormatElapsedClock(s->mic_record_start_tick, now);
  }
  SendMessageW(s->status, SB_SETTEXTW, 0, (LPARAM)info.c_str());
}

static void ShowToast(AppState* s, const std::wstring& msg, int ms = 2200) {
  if (!s) return;
  s->toast = msg;
  s->toast_until = GetTickCount64() + (ULONGLONG)(std::max)(500, ms);
  UpdateStatusBar(s);
}

static void ApplyBodyZoom(AppState* s) {
  if (!s || !s->ed_body) return;
  int zoom = s->body_zoom_percent;
  if (zoom < 50) zoom = 50;
  if (zoom > 300) zoom = 300;
  s->body_zoom_percent = zoom;
  SendMessageW(s->ed_body, EM_SETZOOM, (WPARAM)zoom, 100);
}

static void SetBodyZoom(AppState* s, int zoom_percent) {
  if (!s) return;
  int old_zoom = s->body_zoom_percent;
  s->body_zoom_percent = zoom_percent;
  ApplyBodyZoom(s);
  if (s->body_zoom_percent != old_zoom) {
    UpdateStatusBar(s);
    ShowToast(s, L"正文缩放 " + std::to_wstring(s->body_zoom_percent) + L"%", 1200);
  }
}

static bool AdjustBodyZoomFromWheel(AppState* s, short wheel_delta) {
  if (!s || !s->ed_body) return false;
  int step = (wheel_delta > 0) ? 10 : -10;
  int old_zoom = s->body_zoom_percent;
  int next_zoom = old_zoom + step;
  if (next_zoom < 50) next_zoom = 50;
  if (next_zoom > 300) next_zoom = 300;
  if (next_zoom == old_zoom) return true;
  SetBodyZoom(s, next_zoom);
  return true;
}

static LRESULT CALLBACK BodyEditSubclassProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam,
                                             UINT_PTR subclass_id, DWORD_PTR ref_data) {
  auto* s = reinterpret_cast<AppState*>(ref_data);
  if (msg == WM_MOUSEWHEEL) {
    bool ctrl = (GET_KEYSTATE_WPARAM(wParam) & MK_CONTROL) != 0;
    if (ctrl && s && s->ed_body == hwnd) {
      short delta = GET_WHEEL_DELTA_WPARAM(wParam);
      if (delta != 0 && AdjustBodyZoomFromWheel(s, delta)) return 0;
    }
  }
  if (msg == WM_NCDESTROY) {
    RemoveWindowSubclass(hwnd, BodyEditSubclassProc, subclass_id);
  }
  return DefSubclassProc(hwnd, msg, wParam, lParam);
}

static const wchar_t* AudioSourceCn(AudioRecordSource source) {
  return source == AudioRecordSource::MicrophoneCapture ? L"麦克风" : L"系统声音";
}

static const wchar_t* AudioSourceShortCn(AudioRecordSource source) {
  return source == AudioRecordSource::MicrophoneCapture ? L"麦克" : L"系统";
}

static constexpr ULONGLONG kMaxAudioRecordMs = 5ULL * 60ULL * 60ULL * 1000ULL;

static AudioRecorder** AudioRecorderSlot(AppState* s, AudioRecordSource source) {
  if (!s) return nullptr;
  return source == AudioRecordSource::MicrophoneCapture ? &s->mic_recorder : &s->system_recorder;
}

static ULONGLONG* AudioRecordStartTickSlot(AppState* s, AudioRecordSource source) {
  if (!s) return nullptr;
  return source == AudioRecordSource::MicrophoneCapture ? &s->mic_record_start_tick : &s->system_record_start_tick;
}

static HWND AudioRecordButton(AppState* s, AudioRecordSource source) {
  if (!s) return nullptr;
  return source == AudioRecordSource::MicrophoneCapture ? s->btn_record_mic : s->btn_record_system;
}

static bool IsAudioRecording(AppState* s, AudioRecordSource source) {
  AudioRecorder** slot = AudioRecorderSlot(s, source);
  return slot && AudioRecorderIsRunning(*slot);
}

static bool IsAnyAudioRecording(AppState* s) {
  return IsAudioRecording(s, AudioRecordSource::SystemLoopback) ||
         IsAudioRecording(s, AudioRecordSource::MicrophoneCapture);
}

static std::wstring FormatElapsedClock(ULONGLONG start_tick, ULONGLONG now) {
  if (start_tick == 0 || now < start_tick) return L"00:00";
  int sec = (int)((now - start_tick) / 1000ull);
  int mm = sec / 60;
  int ss = sec % 60;
  wchar_t buf[16]{};
  wsprintfW(buf, L"%02d:%02d", mm, ss);
  return buf;
}

static void UpdateAudioRecordButtons(AppState* s) {
  if (!s) return;
  ULONGLONG now = GetTickCount64();

  if (s->btn_record_system) {
    std::wstring text = IsAudioRecording(s, AudioRecordSource::SystemLoopback)
                            ? (std::wstring(L"停系统 ") + FormatElapsedClock(s->system_record_start_tick, now))
                            : L"录系统声";
    SetWindowTextW(s->btn_record_system, text.c_str());
    InvalidateRect(s->btn_record_system, nullptr, TRUE);
  }
  if (s->btn_record_mic) {
    std::wstring text = IsAudioRecording(s, AudioRecordSource::MicrophoneCapture)
                            ? (std::wstring(L"停麦克 ") + FormatElapsedClock(s->mic_record_start_tick, now))
                            : L"录麦克风";
    SetWindowTextW(s->btn_record_mic, text.c_str());
    InvalidateRect(s->btn_record_mic, nullptr, TRUE);
  }
}

static void EnsureAudioTimerState(AppState* s) {
  if (!s || !s->hwnd) return;
  if (IsAnyAudioRecording(s)) {
    if (!s->audio_timer_id) s->audio_timer_id = SetTimer(s->hwnd, 2, 250, nullptr);
  } else if (s->audio_timer_id) {
    KillTimer(s->hwnd, s->audio_timer_id);
    s->audio_timer_id = 0;
  }
}

static std::wstring GetPathFileNamePart(const std::wstring& path) {
  size_t pos = path.find_last_of(L"\\/");
  return (pos == std::wstring::npos) ? path : path.substr(pos + 1);
}

static std::wstring BuildDefaultAudioSaveName(AudioRecordSource source) {
  SYSTEMTIME st{};
  GetLocalTime(&st);
  wchar_t buf[128]{};
  wsprintfW(buf, source == AudioRecordSource::MicrophoneCapture
                     ? L"WorkLogLite-microphone-audio-%04d-%02d-%02d_%02d-%02d-%02d.wav"
                     : L"WorkLogLite-system-audio-%04d-%02d-%02d_%02d-%02d-%02d.wav",
            (int)st.wYear, (int)st.wMonth, (int)st.wDay,
            (int)st.wHour, (int)st.wMinute, (int)st.wSecond);
  return buf;
}

static bool PromptSaveRecordedAudio(HWND owner,
                                    AudioRecordSource source,
                                    const std::wstring& recorded_path,
                                    std::wstring* out_path,
                                    bool* out_cancelled,
                                    std::wstring* out_err) {
  if (out_path) out_path->clear();
  if (out_cancelled) *out_cancelled = false;
  if (out_err) out_err->clear();

  std::wstring initial_dir = ParentDirOf(recorded_path);
  if (initial_dir.empty()) initial_dir = GetAudioRecordingsDir(source);
  std::wstring initial_name = GetPathFileNamePart(recorded_path);
  if (initial_name.empty()) initial_name = BuildDefaultAudioSaveName(source);

  wchar_t filebuf[MAX_PATH * 4]{};
  wcsncpy_s(filebuf, _countof(filebuf), initial_name.c_str(), _TRUNCATE);

  wchar_t filter[] = L"WAV 音频 (*.wav)\0*.wav\0所有文件 (*.*)\0*.*\0\0";
  OPENFILENAMEW ofn{};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = owner;
  ofn.lpstrFilter = filter;
  ofn.lpstrInitialDir = initial_dir.c_str();
  ofn.lpstrFile = filebuf;
  ofn.nMaxFile = (DWORD)_countof(filebuf);
  ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;
  ofn.lpstrDefExt = L"wav";
  ofn.lpstrTitle = source == AudioRecordSource::MicrophoneCapture ? L"保存麦克风录音" : L"保存系统声音录音";
  if (!GetSaveFileNameW(&ofn)) {
    DWORD dlg_err = CommDlgExtendedError();
    if (dlg_err == 0) {
      if (out_cancelled) *out_cancelled = true;
    } else if (out_err) {
      *out_err = L"打开保存录音对话框失败。(common-dialog error=" + std::to_wstring(dlg_err) + L")";
    }
    return false;
  }
  if (out_path) *out_path = filebuf;
  return true;
}

static bool MoveOrCopyFileReplace(const std::wstring& from, const std::wstring& to, std::wstring* err) {
  if (err) err->clear();
  if (from.empty() || to.empty()) {
    if (err) *err = L"录音文件路径无效。";
    return false;
  }
  if (NormalizeDirForCompare(from) == NormalizeDirForCompare(to)) return true;
  if (MoveFileExW(from.c_str(), to.c_str(),
                  MOVEFILE_REPLACE_EXISTING | MOVEFILE_COPY_ALLOWED | MOVEFILE_WRITE_THROUGH) != 0) {
    return true;
  }
  DWORD le = GetLastError();
  if (err) {
    *err = L"无法保存录音文件：\n" + from + L"\n->\n" + to +
           L"\n(error=" + std::to_wstring(le) + L" " + FormatWin32ErrorMessage(le) + L")";
  }
  return false;
}

static void StopAudioRecorderWithoutPrompt(AppState* s, AudioRecordSource source) {
  if (!s) return;
  AudioRecorder** slot = AudioRecorderSlot(s, source);
  ULONGLONG* start_tick = AudioRecordStartTickSlot(s, source);
  if (!slot || !*slot) return;
  std::wstring path;
  std::wstring err;
  StopAudioRecorder(slot, &path, &err);
  if (start_tick) *start_tick = 0;
}

static void FinishAudioRecording(AppState* s, AudioRecordSource source) {
  if (!s) return;
  AudioRecorder** slot = AudioRecorderSlot(s, source);
  ULONGLONG* start_tick = AudioRecordStartTickSlot(s, source);
  if (!slot || !*slot) return;

  std::wstring temp_path;
  std::wstring err;
  bool ok = StopAudioRecorder(slot, &temp_path, &err);
  if (start_tick) *start_tick = 0;
  UpdateAudioRecordButtons(s);
  EnsureAudioTimerState(s);
  UpdateStatusBar(s);

  if (!ok) {
    ShowInfoBox(s->hwnd, err.empty() ? L"停止录音失败。" : err.c_str(), L"录音失败");
    return;
  }

  std::wstring save_path;
  bool cancelled = false;
  std::wstring save_err;
  if (!PromptSaveRecordedAudio(s->hwnd, source, temp_path, &save_path, &cancelled, &save_err)) {
    if (cancelled) {
      std::wstring folder = ParentDirOf(temp_path);
      if (!folder.empty()) OpenPathInShell(s->hwnd, folder, L"录音目录");
      ShowToast(s, std::wstring(AudioSourceCn(source)) + L"录音已保留到默认目录");
      return;
    }
    std::wstring msg = save_err.empty() ? L"打开保存窗口失败。录音已保留在默认目录。" : save_err + L"\n\n录音已保留在默认目录。";
    ShowInfoBox(s->hwnd, msg.c_str(), L"录音保存");
    return;
  }

  if (!MoveOrCopyFileReplace(temp_path, save_path, &err)) {
    ShowInfoBox(s->hwnd, err.c_str(), L"录音保存失败");
    return;
  }

  std::wstring folder = ParentDirOf(save_path);
  if (!folder.empty()) OpenPathInShell(s->hwnd, folder, L"录音目录");
  (void)OpenPathInShell(s->hwnd, save_path, L"录音文件");
  ShowToast(s, std::wstring(AudioSourceCn(source)) + L"录音已保存");
}

static bool MaybeAutoStopAudioRecording(AppState* s, AudioRecordSource source) {
  if (!s || !IsAudioRecording(s, source)) return false;
  ULONGLONG* start_tick = AudioRecordStartTickSlot(s, source);
  if (!start_tick || *start_tick == 0) return false;
  ULONGLONG now = GetTickCount64();
  if (now - *start_tick < kMaxAudioRecordMs) return false;
  FinishAudioRecording(s, source);
  ShowToast(s, std::wstring(AudioSourceCn(source)) + L"录音已达到 5 小时，已自动停止");
  return true;
}

static void ToggleAudioRecording(AppState* s, AudioRecordSource source) {
  if (!s) return;
  if (IsAudioRecording(s, source)) {
    FinishAudioRecording(s, source);
    return;
  }

  AudioRecorder** slot = AudioRecorderSlot(s, source);
  ULONGLONG* start_tick = AudioRecordStartTickSlot(s, source);
  if (!slot || !start_tick) return;

  std::wstring path;
  std::wstring err;
  if (!StartAudioRecorder(s->hwnd, source, slot, &path, &err)) {
    ShowInfoBox(s->hwnd, err.empty() ? L"启动录音失败。" : err.c_str(), L"录音失败");
    return;
  }

  *start_tick = GetTickCount64();
  UpdateAudioRecordButtons(s);
  EnsureAudioTimerState(s);
  UpdateStatusBar(s);
  ShowToast(s, std::wstring(L"已开始录制") + AudioSourceCn(source));
}

static void DrawOfficeButton(AppState* s, const DRAWITEMSTRUCT* di) {
  if (!di || di->CtlType != ODT_BUTTON) return;
  HDC dc = di->hDC;
  RECT r = di->rcItem;

  bool disabled = (di->itemState & ODS_DISABLED) != 0;
  bool pressed = (di->itemState & ODS_SELECTED) != 0;
  bool focused = (di->itemState & ODS_FOCUS) != 0;

  int control_id = GetDlgCtrlID(di->hwndItem);
  bool recording = s && ((control_id == IDC_BTN_RECORD_SYSTEM && IsAudioRecording(s, AudioRecordSource::SystemLoopback)) ||
                         (control_id == IDC_BTN_RECORD_MIC && IsAudioRecording(s, AudioRecordSource::MicrophoneCapture)));

  bool mung = s && s->mung_theme;
  COLORREF bg = disabled ? (mung ? RGB(224, 231, 217) : RGB(245, 245, 245))
                         : (recording ? (pressed ? RGB(252, 226, 226) : RGB(255, 241, 241))
                                      : (mung ? (pressed ? RGB(202, 217, 192) : RGB(217, 228, 208))
                                              : (pressed ? RGB(225, 240, 255) : RGB(248, 249, 251))));
  COLORREF border = disabled ? (mung ? RGB(170, 182, 160) : RGB(215, 215, 215))
                             : (recording ? (pressed ? RGB(196, 46, 46) : RGB(220, 80, 80))
                                          : (mung ? (pressed ? RGB(108, 133, 96) : RGB(132, 152, 122))
                                                  : (pressed ? RGB(0, 120, 215) : RGB(206, 210, 216))));
  COLORREF text = disabled ? RGB(150, 150, 150) : (recording ? RGB(130, 18, 18) : (mung ? RGB(46, 62, 41) : RGB(25, 25, 25)));

  // Outer rounded rect.
  HBRUSH br = CreateSolidBrush(bg);
  HPEN pen = CreatePen(PS_SOLID, 1, border);
  HGDIOBJ ob = SelectObject(dc, br);
  HGDIOBJ op = SelectObject(dc, pen);
  SetBkMode(dc, TRANSPARENT);

  RECT rr = r;
  InflateRect(&rr, -1, -1);
  RoundRect(dc, rr.left, rr.top, rr.right, rr.bottom, 10, 10);

  // Label.
  wchar_t buf[256]{};
  GetWindowTextW(di->hwndItem, buf, (int)(sizeof(buf) / sizeof(buf[0])));
  SetTextColor(dc, text);
  RECT tr = rr;
  InflateRect(&tr, -8, -2);
  if (recording) {
    RECT dot{tr.left + 6, tr.top + ((tr.bottom - tr.top) - 10) / 2, tr.left + 16, tr.top + ((tr.bottom - tr.top) - 10) / 2 + 10};
    HBRUSH dot_br = CreateSolidBrush(RGB(220, 40, 40));
    HGDIOBJ old_dot = SelectObject(dc, dot_br);
    HPEN dot_pen = CreatePen(PS_SOLID, 1, RGB(220, 40, 40));
    HGDIOBJ old_pen = SelectObject(dc, dot_pen);
    Ellipse(dc, dot.left, dot.top, dot.right, dot.bottom);
    SelectObject(dc, old_pen);
    SelectObject(dc, old_dot);
    DeleteObject(dot_pen);
    DeleteObject(dot_br);
    tr.left += 22;
    DrawTextW(dc, buf, -1, &tr, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
  } else {
    DrawTextW(dc, buf, -1, &tr, DT_CENTER | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS | DT_NOPREFIX);
  }

  if (focused && !disabled) {
    RECT fr = rr;
    InflateRect(&fr, -3, -3);
    DrawFocusRect(dc, &fr);
  }

  SelectObject(dc, op);
  SelectObject(dc, ob);
  DeleteObject(pen);
  DeleteObject(br);
}

static COLORREF ThemeWindowColor(const AppState* s) {
  return (s && s->mung_theme) ? RGB(228, 236, 218) : GetSysColor(COLOR_WINDOW);
}

static COLORREF ThemePanelColor(const AppState* s) {
  return (s && s->mung_theme) ? RGB(217, 228, 208) : GetSysColor(COLOR_BTNFACE);
}

static COLORREF ThemeEditColor(const AppState* s) {
  return (s && s->mung_theme) ? RGB(238, 244, 232) : GetSysColor(COLOR_WINDOW);
}

static COLORREF ThemeTextColor(const AppState* s) {
  return (s && s->mung_theme) ? RGB(46, 62, 41) : RGB(20, 20, 20);
}

static void RecreateThemeBrushes(AppState* s) {
  if (!s) return;
  if (s->brush_window) DeleteObject(s->brush_window);
  if (s->brush_panel) DeleteObject(s->brush_panel);
  if (s->brush_edit) DeleteObject(s->brush_edit);
  s->brush_window = CreateSolidBrush(ThemeWindowColor(s));
  s->brush_panel = CreateSolidBrush(ThemePanelColor(s));
  s->brush_edit = CreateSolidBrush(ThemeEditColor(s));
}

static COLORREF ThemedStatusBkColor(AppState* s, EntryStatus st) {
  if (!s || !s->mung_theme) return StatusBkColor(st);
  switch (st) {
    case EntryStatus::Todo: return RGB(230, 238, 223);
    case EntryStatus::Doing: return RGB(220, 232, 212);
    case EntryStatus::Blocked: return RGB(236, 233, 216);
    case EntryStatus::Done: return RGB(214, 232, 208);
    case EntryStatus::None:
    default: return ThemeEditColor(s);
  }
}

static void UpdateThemeMenuState(AppState* s) {
  if (!s || !s->hwnd) return;
  HMENU menu = GetMenu(s->hwnd);
  if (!menu) return;
  CheckMenuItem(menu, IDM_THEME_MUNG, MF_BYCOMMAND | (s->mung_theme ? MF_CHECKED : MF_UNCHECKED));
  DrawMenuBar(s->hwnd);
}

static void ApplyTheme(AppState* s) {
  if (!s) return;
  RecreateThemeBrushes(s);

  if (s->list) {
    ListView_SetBkColor(s->list, ThemeEditColor(s));
    ListView_SetTextBkColor(s->list, ThemeEditColor(s));
    ListView_SetTextColor(s->list, ThemeTextColor(s));
  }
  if (s->ed_body) SendMessageW(s->ed_body, EM_SETBKGNDCOLOR, 0, ThemeEditColor(s));
  if (s->status) SendMessageW(s->status, SB_SETBKCOLOR, 0, ThemePanelColor(s));

  UpdateThemeMenuState(s);
  InvalidateRect(s->hwnd, nullptr, TRUE);
  if (s->list) InvalidateRect(s->list, nullptr, TRUE);
  if (s->ed_body) InvalidateRect(s->ed_body, nullptr, TRUE);
  if (s->status) InvalidateRect(s->status, nullptr, TRUE);
}

static void UpdateFocusTimerButton(AppState* s) {
  if (!s || !s->btn_focus_timer) return;
  if (!s->focus_running) {
    SetWindowTextW(s->btn_focus_timer, L"专注 25:00");
    return;
  }
  ULONGLONG now = GetTickCount64();
  ULONGLONG remain = (s->focus_end_tick > now) ? (s->focus_end_tick - now) : 0;
  int sec = (int)(remain / 1000);
  int mm = sec / 60;
  int ss = sec % 60;
  wchar_t buf[64]{};
  wsprintfW(buf, L"停止 %02d:%02d", mm, ss);
  SetWindowTextW(s->btn_focus_timer, buf);
}

static void StopFocusTimer(AppState* s) {
  if (!s) return;
  if (s->focus_timer_id) {
    KillTimer(s->hwnd, s->focus_timer_id);
    s->focus_timer_id = 0;
  }
  s->focus_running = false;
  s->focus_end_tick = 0;
  UpdateFocusTimerButton(s);
  UpdateStatusBar(s);
}

static void StartFocusTimer(AppState* s, int minutes) {
  if (!s) return;
  ULONGLONG now = GetTickCount64();
  s->focus_running = true;
  s->focus_end_tick = now + (ULONGLONG)minutes * 60ull * 1000ull;
  if (!s->focus_timer_id) s->focus_timer_id = SetTimer(s->hwnd, 1, 250, nullptr);
  UpdateFocusTimerButton(s);
  UpdateStatusBar(s);
}

static std::vector<std::wstring> DefaultQuickReplies() {
  return {
      L"收到，我看一下。",
      L"已同步，辛苦。",
      L"我这边已处理，麻烦再确认一下。",
      L"这个点我需要再确认一下细节，稍后回复。",
      L"我建议先这样做：\n- \n- ",
      L"今日进展：\n- \n\n风险/阻塞：\n- \n\n下一步：\n- ",
  };
}

static std::wstring GetQuickRepliesFilePath() {
  return JoinPath(GetDataRootDir(), L"quick_replies.txt");
}

static std::wstring GetScratchpadFilePath() {
  return JoinPath(GetDataRootDir(), L"scratchpad.txt");
}

static void OpenScratchpadFile(HWND owner) {
  EnsureDirExists(GetDataRootDir());
  std::wstring path = GetScratchpadFilePath();
  if (!FileExists(path)) {
    std::wstring err;
    if (!WriteUtf8File(path, L"", &err)) {
      ShowInfoBox(owner, err.c_str(), L"便签");
      return;
    }
  }
  HINSTANCE r = ShellExecuteW(nullptr, L"open", L"notepad.exe", path.c_str(), nullptr, SW_SHOWNORMAL);
  if ((INT_PTR)r <= 32) ShowInfoBox(owner, L"无法打开便签文件。", L"便签");
}

static bool OpenPathInShell(HWND owner, const std::wstring& path, const wchar_t* what) {
  if (path.empty()) {
    ShowInfoBox(owner, L"路径为空，无法打开。", what ? what : L"提示");
    return false;
  }
  HINSTANCE r = ShellExecuteW(nullptr, L"open", path.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
  if ((INT_PTR)r <= 32) {
    std::wstring msg = L"无法打开路径：\n" + path;
    ShowInfoBox(owner, msg.c_str(), what ? what : L"提示");
    return false;
  }
  return true;
}

static void OpenSelectedDayFolder(AppState* s) {
  if (!s) return;
  std::wstring path = GetDayFilePath(s->selected);
  std::wstring dir = ParentDirOf(path);
  if (dir.empty()) return;
  EnsureDirExists(dir);
  ShellExecuteW(nullptr, L"open", dir.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
}

static std::vector<std::wstring> ParseQuickRepliesContent(const std::wstring& content) {
  // New format (recommended): blocks separated by a line that is exactly "---".
  // This supports multi-line replies and blank lines.
  bool has_sep = false;
  {
    std::wstringstream ss(content);
    std::wstring line;
    while (std::getline(ss, line)) {
      if (!line.empty() && line.back() == L'\r') line.pop_back();
      if (line == L"---") {
        has_sep = true;
        break;
      }
    }
  }

  std::vector<std::wstring> out;
  if (has_sep) {
    std::wstringstream ss(content);
    std::wstring line;
    std::wstring cur;
    while (std::getline(ss, line)) {
      if (!line.empty() && line.back() == L'\r') line.pop_back();
      if (line == L"---") {
        while (!cur.empty() && (cur.back() == L'\n' || cur.back() == L'\r')) cur.pop_back();
        if (!cur.empty()) out.push_back(cur);
        cur.clear();
        continue;
      }
      cur += line;
      cur += L"\n";
    }
    while (!cur.empty() && (cur.back() == L'\n' || cur.back() == L'\r')) cur.pop_back();
    if (!cur.empty()) out.push_back(cur);
    return out;
  }

  // Legacy format: each non-empty line is a reply (single-line only).
  std::wstringstream ss(content);
  std::wstring line;
  while (std::getline(ss, line)) {
    if (!line.empty() && line.back() == L'\r') line.pop_back();
    if (line.empty()) continue;
    out.push_back(line);
  }
  return out;
}

static std::wstring SerializeQuickReplies(const std::vector<std::wstring>& replies) {
  // Always write in block format to preserve multi-line replies.
  std::wstring out;
  for (size_t i = 0; i < replies.size(); i++) {
    out += replies[i];
    if (out.empty() || out.back() != L'\n') out += L"\n";
    out += L"---\n";
  }
  return out;
}

static bool LoadQuickReplies(std::vector<std::wstring>* out, std::wstring* err) {
  if (!out) return false;
  out->clear();

  EnsureDirExists(GetDataRootDir());
  std::wstring path = GetQuickRepliesFilePath();
  if (!FileExists(path)) {
    *out = DefaultQuickReplies();
    // Best-effort persist.
    std::wstring save_err;
    WriteUtf8File(path, SerializeQuickReplies(*out), &save_err);
    return true;
  }

  std::wstring content;
  if (!ReadFileUtf8Simple(path, &content, err)) return false;

  *out = ParseQuickRepliesContent(content);
  if (out->empty()) *out = DefaultQuickReplies();
  return true;
}

static void InsertTextIntoBody(AppState* s, const std::wstring& text) {
  if (!s || !s->ed_body) return;
  SetFocus(s->ed_body);
  // RichEdit expects CRLF for newlines.
  std::wstring t = NormalizeNewlinesToCrlf(text);
  // Add a trailing newline to avoid sticking to next content.
  if (!t.empty() && t.back() != L'\n') t += L"\r\n";
  SendMessageW(s->ed_body, EM_REPLACESEL, TRUE, (LPARAM)t.c_str());
  SetEditorDirty(s, true);
  SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
  UpdateStatusBar(s);
}

static void PastePlainTextIntoBody(AppState* s) {
  if (!s) return;
  std::wstring text;
  if (!ReadClipboardText(s->hwnd, &text) || text.empty()) {
    ShowInfoBox(s->hwnd, L"剪贴板里没有可粘贴的文本。", L"纯文本粘贴");
    return;
  }
  InsertTextIntoBody(s, text);
  ShowToast(s, L"已按纯文本粘贴");
}

static void InsertDateSeparatorIntoBody(AppState* s) {
  if (!s) return;
  SYSTEMTIME st{};
  GetLocalTime(&st);
  wchar_t line[32]{};
  wsprintfW(line, L"----%04d%02d%02d----", (int)st.wYear, (int)st.wMonth, (int)st.wDay);

  if (!s->ed_body) return;
  SetFocus(s->ed_body);

  CHARRANGE before{};
  SendMessageW(s->ed_body, EM_EXGETSEL, 0, (LPARAM)&before);
  CHARFORMAT2W base_cf{};
  base_cf.cbSize = sizeof(base_cf);
  SendMessageW(s->ed_body, EM_GETCHARFORMAT, (WPARAM)SCF_SELECTION, (LPARAM)&base_cf);

  std::wstring insert_text;
  LONG prefix_len = 0;
  if (before.cpMin > 0) {
    wchar_t prev[2]{};
    TEXTRANGEW tr{};
    tr.chrg.cpMin = before.cpMin - 1;
    tr.chrg.cpMax = before.cpMin;
    tr.lpstrText = prev;
    SendMessageW(s->ed_body, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
    if (prev[0] != L'\n' && prev[0] != L'\r') {
      insert_text += L"\r\n";
      prefix_len = 2;
    }
  }
  insert_text += line;
  insert_text += L"\r\n";

  SendMessageW(s->ed_body, EM_REPLACESEL, TRUE, (LPARAM)insert_text.c_str());
  LONG line_len = (LONG)wcslen(line);
  LONG line_start = before.cpMin + prefix_len;
  LONG line_end = line_start + line_len;
  if (line_start >= 0 && line_end >= line_start) {
    CHARRANGE cr{};
    cr.cpMin = line_start;
    cr.cpMax = line_end;
    SendMessageW(s->ed_body, EM_EXSETSEL, 0, (LPARAM)&cr);
    CHARFORMAT2W bold_cf = base_cf;
    bold_cf.cbSize = sizeof(bold_cf);
    bold_cf.dwMask |= CFM_BOLD | CFM_FACE | CFM_SIZE | CFM_CHARSET | CFM_COLOR;
    bold_cf.dwEffects |= CFE_BOLD;
    SendMessageW(s->ed_body, EM_SETCHARFORMAT, (WPARAM)SCF_SELECTION, (LPARAM)&bold_cf);

    CHARRANGE after{};
    SendMessageW(s->ed_body, EM_EXGETSEL, 0, (LPARAM)&after);
    cr.cpMin = after.cpMax;
    cr.cpMax = after.cpMax;
    SendMessageW(s->ed_body, EM_EXSETSEL, 0, (LPARAM)&cr);

    CHARFORMAT2W clear_cf = base_cf;
    clear_cf.cbSize = sizeof(clear_cf);
    clear_cf.dwMask |= CFM_BOLD | CFM_FACE | CFM_SIZE | CFM_CHARSET | CFM_COLOR;
    clear_cf.dwEffects &= ~CFE_BOLD;
    SendMessageW(s->ed_body, EM_SETCHARFORMAT, (WPARAM)SCF_SELECTION, (LPARAM)&clear_cf);
  }

  SetEditorDirty(s, true);
  SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
  UpdateStatusBar(s);
  ShowToast(s, L"已插入时间分隔符");
}

static void PickBodyTextColor(AppState* s) {
  if (!s) return;
  COLORREF color = RGB(0, 0, 0);
  std::wstring err;
  if (!PickScreenColor(s->hwnd, &color, &err)) {
    if (!err.empty()) ShowInfoBox(s->hwnd, err.c_str(), L"取色器");
    return;
  }
  wchar_t buf[16]{};
  wsprintfW(buf, L"#%02X%02X%02X", GetRValue(color), GetGValue(color), GetBValue(color));
  CopyToClipboard(s->hwnd, buf);
  ShowToast(s, std::wstring(L"颜色已复制 ") + buf);
}

static void ShowQuickReplyMenu(AppState* s) {
  if (!s) return;

  std::vector<std::wstring> replies;
  std::wstring err;
  if (!LoadQuickReplies(&replies, &err)) {
    ShowInfoBox(s->hwnd, err.c_str(), L"读取常用回复失败");
    return;
  }
  if (replies.empty()) return;

  enum : UINT { IDM_QR_BASE = 14000 };
  HMENU menu = CreatePopupMenu();
  if (!menu) return;
  UINT id = IDM_QR_BASE;
  for (const auto& r : replies) {
    if (id >= IDM_QR_BASE + 200) break;
    std::wstring one = r;
    size_t nl = one.find(L'\n');
    if (nl != std::wstring::npos) one = one.substr(0, nl) + L" ...";
    if (one.size() > 60) one = one.substr(0, 60) + L"...";
    AppendMenuW(menu, MF_STRING, id++, one.c_str());
  }
  AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(menu, MF_STRING, IDM_QR_BASE + 999, L"编辑 quick_replies.txt");
  AppendMenuW(menu, MF_STRING, IDM_QR_BASE + 1000, L"打开 quick_replies.txt 所在数据目录");

  POINT pt{};
  GetCursorPos(&pt);
  UINT cmd = TrackPopupMenu(menu, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, 0, s->hwnd, nullptr);
  DestroyMenu(menu);
  if (!cmd) return;

  if (cmd == IDM_QR_BASE + 999) {
    std::wstring path = GetQuickRepliesFilePath();
    ShellExecuteW(nullptr, L"open", L"notepad.exe", path.c_str(), nullptr, SW_SHOWNORMAL);
    return;
  }
  if (cmd == IDM_QR_BASE + 1000) {
    OpenDataDir(s);
    return;
  }

  size_t idx = (size_t)(cmd - IDM_QR_BASE);
  if (idx >= replies.size()) return;

  std::wstring t = replies[idx];
  CopyToClipboard(s->hwnd, NormalizeNewlinesToCrlf(t));
  ShowToast(s, L"常用回复已复制到剪贴板");
}

enum : int {
  IDC_UCALC_VALUE = 17001,
  IDC_UCALC_FROM = 17002,
  IDC_UCALC_TO = 17003,
  IDC_UCALC_RESULT = 17004,
  IDC_UCALC_SIZE = 17005,
  IDC_UCALC_SIZE_UNIT = 17006,
  IDC_UCALC_DURATION = 17007,
  IDC_UCALC_DURATION_UNIT = 17008,
  IDC_UCALC_BW_RESULT = 17009,
  IDC_UCALC_COPY = 17010,
  IDC_UCALC_SYSTEM = 17011,
  IDC_UCALC_CLOSE = 17012,
};

struct UnitCalcUnit {
  const wchar_t* label;
  double bytes;
  bool is_rate;
};

struct UnitCalcTimeUnit {
  const wchar_t* label;
  double seconds;
};

static const UnitCalcUnit kUnitCalcUnits[] = {
    {L"B", 1.0, false},
    {L"KB", 1000.0, false},
    {L"MB", 1000.0 * 1000.0, false},
    {L"GB", 1000.0 * 1000.0 * 1000.0, false},
    {L"TB", 1000.0 * 1000.0 * 1000.0 * 1000.0, false},
    {L"KiB", 1024.0, false},
    {L"MiB", 1024.0 * 1024.0, false},
    {L"GiB", 1024.0 * 1024.0 * 1024.0, false},
    {L"TiB", 1024.0 * 1024.0 * 1024.0 * 1024.0, false},
    {L"bps", 1.0 / 8.0, true},
    {L"Kbps", 1000.0 / 8.0, true},
    {L"Mbps", 1000.0 * 1000.0 / 8.0, true},
    {L"Gbps", 1000.0 * 1000.0 * 1000.0 / 8.0, true},
    {L"B/s", 1.0, true},
    {L"KB/s", 1000.0, true},
    {L"MB/s", 1000.0 * 1000.0, true},
    {L"GB/s", 1000.0 * 1000.0 * 1000.0, true},
};

static const int kUnitCalcDataUnitIndexes[] = {0, 1, 2, 3, 4, 5, 6, 7, 8};

static const UnitCalcTimeUnit kUnitCalcTimeUnits[] = {
    {L"秒", 1.0},
    {L"分钟", 60.0},
    {L"小时", 3600.0},
    {L"天", 86400.0},
};

struct UnitCalcWindowState {
  HWND hwnd{};
  HWND st_convert{};
  HWND st_bandwidth{};
  HWND ed_value{};
  HWND cb_from{};
  HWND cb_to{};
  HWND ed_result{};
  HWND ed_size{};
  HWND cb_size_unit{};
  HWND ed_duration{};
  HWND cb_duration_unit{};
  HWND ed_bw_result{};
  HWND btn_copy{};
  HWND btn_system{};
  HWND btn_close{};
  HFONT font{};
};

static void CenterOwnedWindowSimple(HWND hwnd, HWND owner) {
  if (!hwnd) return;
  RECT rc{};
  RECT wa{};
  if (!GetWindowRect(hwnd, &rc)) return;
  SystemParametersInfoW(SPI_GETWORKAREA, 0, &wa, 0);
  int width = rc.right - rc.left;
  int height = rc.bottom - rc.top;
  int x = wa.left + ((wa.right - wa.left) - width) / 2;
  int y = wa.top + ((wa.bottom - wa.top) - height) / 2;
  if (owner) {
    RECT orc{};
    if (GetWindowRect(owner, &orc)) {
      x = orc.left + ((orc.right - orc.left) - width) / 2;
      y = orc.top + ((orc.bottom - orc.top) - height) / 2;
    }
  }
  if (x < wa.left) x = wa.left;
  if (y < wa.top) y = wa.top;
  if (x + width > wa.right) x = wa.right - width;
  if (y + height > wa.bottom) y = wa.bottom - height;
  SetWindowPos(hwnd, nullptr, x, y, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);
}

static bool TryParseDoubleText(const std::wstring& text, double* out) {
  if (out) *out = 0.0;
  const wchar_t* s = text.c_str();
  while (*s && iswspace(*s)) s++;
  if (!*s) return false;
  wchar_t* end = nullptr;
  double v = wcstod(s, &end);
  if (end == s || !std::isfinite(v)) return false;
  while (*end && iswspace(*end)) end++;
  if (*end) return false;
  if (out) *out = v;
  return true;
}

static std::wstring FormatDoubleCompact(double value, int digits = 4) {
  wchar_t buf[128]{};
  swprintf_s(buf, L"%.*f", digits, value);
  std::wstring out = buf;
  while (!out.empty() && out.back() == L'0') out.pop_back();
  if (!out.empty() && out.back() == L'.') out.pop_back();
  if (out.empty()) out = L"0";
  return out;
}

static void UnitCalcFillUnits(HWND combo, bool data_only) {
  SendMessageW(combo, CB_RESETCONTENT, 0, 0);
  if (data_only) {
    for (int idx : kUnitCalcDataUnitIndexes) {
      LRESULT item = SendMessageW(combo, CB_ADDSTRING, 0, (LPARAM)kUnitCalcUnits[idx].label);
      SendMessageW(combo, CB_SETITEMDATA, (WPARAM)item, (LPARAM)idx);
    }
    return;
  }
  for (int i = 0; i < (int)(sizeof(kUnitCalcUnits) / sizeof(kUnitCalcUnits[0])); i++) {
    LRESULT item = SendMessageW(combo, CB_ADDSTRING, 0, (LPARAM)kUnitCalcUnits[i].label);
    SendMessageW(combo, CB_SETITEMDATA, (WPARAM)item, (LPARAM)i);
  }
}

static int UnitCalcGetComboUnitIndex(HWND combo) {
  int sel = (int)SendMessageW(combo, CB_GETCURSEL, 0, 0);
  if (sel < 0) return -1;
  return (int)SendMessageW(combo, CB_GETITEMDATA, (WPARAM)sel, 0);
}

static void UnitCalcFillTimeUnits(HWND combo) {
  SendMessageW(combo, CB_RESETCONTENT, 0, 0);
  for (int i = 0; i < (int)(sizeof(kUnitCalcTimeUnits) / sizeof(kUnitCalcTimeUnits[0])); i++) {
    LRESULT item = SendMessageW(combo, CB_ADDSTRING, 0, (LPARAM)kUnitCalcTimeUnits[i].label);
    SendMessageW(combo, CB_SETITEMDATA, (WPARAM)item, (LPARAM)i);
  }
}

static int UnitCalcGetTimeUnitIndex(HWND combo) {
  int sel = (int)SendMessageW(combo, CB_GETCURSEL, 0, 0);
  if (sel < 0) return -1;
  return (int)SendMessageW(combo, CB_GETITEMDATA, (WPARAM)sel, 0);
}

static void UpdateUnitCalcResults(UnitCalcWindowState* s) {
  if (!s) return;

  std::wstring convert_result = L"请输入数值并选择单位。";
  double value = 0.0;
  int from_idx = UnitCalcGetComboUnitIndex(s->cb_from);
  int to_idx = UnitCalcGetComboUnitIndex(s->cb_to);
  if (TryParseDoubleText(GetWindowTextWString(s->ed_value), &value) &&
      from_idx >= 0 && to_idx >= 0) {
    const UnitCalcUnit& from = kUnitCalcUnits[from_idx];
    const UnitCalcUnit& to = kUnitCalcUnits[to_idx];
    if (from.is_rate != to.is_rate) {
      convert_result = L"大小单位和速率单位不能直接互转。";
    } else {
      double converted = value * from.bytes / to.bytes;
      convert_result = FormatDoubleCompact(value) + L" " + from.label + L" = " +
                       FormatDoubleCompact(converted) + L" " + to.label;
    }
  }
  SetWindowTextW(s->ed_result, convert_result.c_str());

  std::wstring bw_result = L"请输入文件大小和耗时，计算所需带宽。";
  double size_value = 0.0;
  double duration_value = 0.0;
  int size_idx = UnitCalcGetComboUnitIndex(s->cb_size_unit);
  int duration_idx = UnitCalcGetTimeUnitIndex(s->cb_duration_unit);
  if (TryParseDoubleText(GetWindowTextWString(s->ed_size), &size_value) &&
      TryParseDoubleText(GetWindowTextWString(s->ed_duration), &duration_value) &&
      size_idx >= 0 && duration_idx >= 0) {
    if (size_value < 0.0 || duration_value <= 0.0) {
      bw_result = L"文件大小不能为负数，耗时必须大于 0。";
    } else {
      double total_bytes = size_value * kUnitCalcUnits[size_idx].bytes;
      double total_seconds = duration_value * kUnitCalcTimeUnits[duration_idx].seconds;
      double bytes_per_sec = total_bytes / total_seconds;
      double mb_per_sec = bytes_per_sec / 1000.0 / 1000.0;
      double mib_per_sec = bytes_per_sec / 1024.0 / 1024.0;
      double mbps = bytes_per_sec * 8.0 / 1000.0 / 1000.0;
      double gbps = mbps / 1000.0;
      bw_result = L"需要带宽约:\r\n" +
                  FormatDoubleCompact(mb_per_sec) + L" MB/s\r\n" +
                  FormatDoubleCompact(mib_per_sec) + L" MiB/s\r\n" +
                  FormatDoubleCompact(mbps) + L" Mbps\r\n" +
                  FormatDoubleCompact(gbps) + L" Gbps";
    }
  }
  SetWindowTextW(s->ed_bw_result, bw_result.c_str());
}

static LRESULT CALLBACK UnitCalcWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
  auto* s = reinterpret_cast<UnitCalcWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
    case WM_CREATE: {
      auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
      s = reinterpret_cast<UnitCalcWindowState*>(cs->lpCreateParams);
      s->hwnd = hwnd;
      SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(s));
      s->font = CreateUiFont(hwnd);
      DWORD edit_style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_AUTOHSCROLL;
      DWORD ro_style = WS_CHILD | WS_VISIBLE | WS_TABSTOP | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY | WS_VSCROLL;

      s->st_convert = CreateWindowExW(0, L"STATIC", L"单位换算", WS_CHILD | WS_VISIBLE,
                                      0, 0, 10, 10, hwnd, nullptr, nullptr, nullptr);
      s->ed_value = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"1024", edit_style,
                                    0, 0, 10, 10, hwnd, (HMENU)IDC_UCALC_VALUE, nullptr, nullptr);
      s->cb_from = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP,
                                   0, 0, 10, 10, hwnd, (HMENU)IDC_UCALC_FROM, nullptr, nullptr);
      s->cb_to = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP,
                                 0, 0, 10, 10, hwnd, (HMENU)IDC_UCALC_TO, nullptr, nullptr);
      s->ed_result = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", ro_style,
                                     0, 0, 10, 10, hwnd, (HMENU)IDC_UCALC_RESULT, nullptr, nullptr);

      s->st_bandwidth = CreateWindowExW(0, L"STATIC", L"带宽需求", WS_CHILD | WS_VISIBLE,
                                        0, 0, 10, 10, hwnd, nullptr, nullptr, nullptr);
      s->ed_size = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"1", edit_style,
                                   0, 0, 10, 10, hwnd, (HMENU)IDC_UCALC_SIZE, nullptr, nullptr);
      s->cb_size_unit = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP,
                                        0, 0, 10, 10, hwnd, (HMENU)IDC_UCALC_SIZE_UNIT, nullptr, nullptr);
      s->ed_duration = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"10", edit_style,
                                       0, 0, 10, 10, hwnd, (HMENU)IDC_UCALC_DURATION, nullptr, nullptr);
      s->cb_duration_unit = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP,
                                            0, 0, 10, 10, hwnd, (HMENU)IDC_UCALC_DURATION_UNIT, nullptr, nullptr);
      s->ed_bw_result = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", ro_style,
                                        0, 0, 10, 10, hwnd, (HMENU)IDC_UCALC_BW_RESULT, nullptr, nullptr);

      s->btn_copy = CreateWindowExW(0, L"BUTTON", L"复制结果", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                    0, 0, 10, 10, hwnd, (HMENU)IDC_UCALC_COPY, nullptr, nullptr);
      s->btn_system = CreateWindowExW(0, L"BUTTON", L"系统计算器", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                      0, 0, 10, 10, hwnd, (HMENU)IDC_UCALC_SYSTEM, nullptr, nullptr);
      s->btn_close = CreateWindowExW(0, L"BUTTON", L"关闭", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                     0, 0, 10, 10, hwnd, (HMENU)IDC_UCALC_CLOSE, nullptr, nullptr);

      UnitCalcFillUnits(s->cb_from, false);
      UnitCalcFillUnits(s->cb_to, false);
      UnitCalcFillUnits(s->cb_size_unit, true);
      UnitCalcFillTimeUnits(s->cb_duration_unit);
      SendMessageW(s->cb_from, CB_SETCURSEL, 2, 0);
      SendMessageW(s->cb_to, CB_SETCURSEL, 3, 0);
      SendMessageW(s->cb_size_unit, CB_SETCURSEL, 3, 0);
      SendMessageW(s->cb_duration_unit, CB_SETCURSEL, 1, 0);

      for (HWND h : {s->st_convert, s->ed_value, s->cb_from, s->cb_to, s->ed_result,
                     s->st_bandwidth, s->ed_size, s->cb_size_unit, s->ed_duration, s->cb_duration_unit,
                     s->ed_bw_result, s->btn_copy, s->btn_system, s->btn_close}) {
        SetControlFont(h, s->font);
      }
      UpdateUnitCalcResults(s);
      return 0;
    }
    case WM_SIZE: {
      if (!s) return 0;
      int w = LOWORD(lParam);
      int h = HIWORD(lParam);
      int pad = ScalePx(hwnd, 10);
      int row = ScalePx(hwnd, 28);
      int lbl = ScalePx(hwnd, 18);
      int btn_h = ScalePx(hwnd, 30);
      int unit_w = ScalePx(hwnd, 102);

      int x = pad;
      int y = pad;
      MoveWindow(s->st_convert, x, y, w - 2 * pad, lbl, TRUE);
      y += lbl + ScalePx(hwnd, 6);

      int value_w = w - 2 * pad - unit_w * 2 - pad * 2;
      MoveWindow(s->ed_value, x, y, value_w, row, TRUE);
      MoveWindow(s->cb_from, x + value_w + pad, y, unit_w, row + 220, TRUE);
      MoveWindow(s->cb_to, x + value_w + pad * 2 + unit_w, y, unit_w, row + 220, TRUE);
      y += row + pad;

      int result_h = ScalePx(hwnd, 52);
      MoveWindow(s->ed_result, x, y, w - 2 * pad, result_h, TRUE);
      y += result_h + pad;

      MoveWindow(s->st_bandwidth, x, y, w - 2 * pad, lbl, TRUE);
      y += lbl + ScalePx(hwnd, 6);

      int half_w = (w - 3 * pad) / 2;
      int edit_w = half_w - unit_w - pad;
      MoveWindow(s->ed_size, x, y, edit_w, row, TRUE);
      MoveWindow(s->cb_size_unit, x + edit_w + pad, y, unit_w, row + 220, TRUE);
      MoveWindow(s->ed_duration, x + half_w + pad, y, edit_w, row, TRUE);
      MoveWindow(s->cb_duration_unit, x + half_w + pad + edit_w + pad, y, unit_w, row + 220, TRUE);
      y += row + pad;

      int bw_h = h - y - pad * 2 - btn_h;
      if (bw_h < ScalePx(hwnd, 96)) bw_h = ScalePx(hwnd, 96);
      MoveWindow(s->ed_bw_result, x, y, w - 2 * pad, bw_h, TRUE);
      y += bw_h + pad;

      int btn_w = ScalePx(hwnd, 96);
      int sys_w = ScalePx(hwnd, 110);
      int bx = w - pad - btn_w;
      MoveWindow(s->btn_close, bx, y, btn_w, btn_h, TRUE);
      bx -= pad + sys_w;
      MoveWindow(s->btn_system, bx, y, sys_w, btn_h, TRUE);
      bx -= pad + btn_w;
      MoveWindow(s->btn_copy, bx, y, btn_w, btn_h, TRUE);
      return 0;
    }
    case WM_GETMINMAXINFO: {
      auto* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
      if (!mmi) return 0;
      RECT rc{0, 0, 700, 520};
      AdjustWindowRectEx(&rc, WS_OVERLAPPEDWINDOW, FALSE, WS_EX_DLGMODALFRAME);
      mmi->ptMinTrackSize.x = rc.right - rc.left;
      mmi->ptMinTrackSize.y = rc.bottom - rc.top;
      return 0;
    }
    case WM_COMMAND: {
      int id = LOWORD(wParam);
      int code = HIWORD(wParam);
      if (id == IDC_UCALC_CLOSE) {
        DestroyWindow(hwnd);
        return 0;
      }
      if (id == IDC_UCALC_SYSTEM) {
        HINSTANCE r = ShellExecuteW(nullptr, L"open", L"calc.exe", nullptr, nullptr, SW_SHOWNORMAL);
        if ((INT_PTR)r <= 32) ShowInfoBox(hwnd, L"无法启动系统计算器(calc.exe)。", L"提示");
        return 0;
      }
      if (id == IDC_UCALC_COPY) {
        std::wstring text = L"单位换算:\r\n" + GetWindowTextWString(s->ed_result) +
                            L"\r\n\r\n带宽需求:\r\n" + GetWindowTextWString(s->ed_bw_result);
        if (!CopyToClipboard(hwnd, text)) ShowInfoBox(hwnd, L"复制失败。", L"提示");
        return 0;
      }
      if ((id == IDC_UCALC_VALUE || id == IDC_UCALC_SIZE || id == IDC_UCALC_DURATION) && code == EN_CHANGE) {
        UpdateUnitCalcResults(s);
        return 0;
      }
      if ((id == IDC_UCALC_FROM || id == IDC_UCALC_TO || id == IDC_UCALC_SIZE_UNIT || id == IDC_UCALC_DURATION_UNIT) &&
          code == CBN_SELCHANGE) {
        UpdateUnitCalcResults(s);
        return 0;
      }
      return 0;
    }
    case WM_CLOSE:
      DestroyWindow(hwnd);
      return 0;
    case WM_DESTROY:
      if (s && s->font) DeleteObject(s->font);
      return 0;
  }
  return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void ShowUnitCalcWindow(HWND owner) {
  WNDCLASSW wc{};
  wc.lpfnWndProc = UnitCalcWndProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.lpszClassName = L"WorkLogLiteUnitCalcWnd";
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
  RegisterClassW(&wc);

  UnitCalcWindowState state{};
  HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME, wc.lpszClassName, L"单位计算与带宽需求",
                              WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                              CW_USEDEFAULT, CW_USEDEFAULT, 760, 560,
                              owner, nullptr, wc.hInstance, &state);
  if (!hwnd) {
    ShowLastErrorBox(owner, L"CreateWindowEx(UnitCalc)");
    return;
  }
  CenterOwnedWindowSimple(hwnd, owner);

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

static void SetEditorMaximized(AppState* s, bool maximized) {
  if (!s) return;
  s->editor_maximized = maximized;
  if (s->btn_body_maximize) {
    SetWindowTextW(s->btn_body_maximize, maximized ? L"还原编辑" : L"最大化编辑");
  }
  Layout(s);
  InvalidateRect(s->hwnd, nullptr, TRUE);
}

static void ApplyBodyFormatting(AppState* s, void (*fn)(HWND)) {
  if (!s || !s->ed_body || !fn) return;
  // Ensure caret/selection state is owned by the RichEdit.
  SetFocus(s->ed_body);
  fn(s->ed_body);
  SetEditorDirty(s, true);
  SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
  UpdateStatusBar(s);
}

static void ApplyBodyFormatting2(AppState* s, void (*fn)(HWND, int), int arg) {
  if (!s || !s->ed_body || !fn) return;
  SetFocus(s->ed_body);
  fn(s->ed_body, arg);
  SetEditorDirty(s, true);
  SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
  UpdateStatusBar(s);
}

static void BodyFmtBold(HWND rich) { RichToggleCharEffect(rich, CFE_BOLD); }
static void BodyFmtItalic(HWND rich) { RichToggleCharEffect(rich, CFE_ITALIC); }
static void BodyFmtUnderline(HWND rich) { RichToggleCharEffect(rich, CFE_UNDERLINE); }
static void BodyFmtSuperscript(HWND rich) { RichToggleScriptEffect(rich, CFE_SUPERSCRIPT, CFE_SUBSCRIPT); }
static void BodyFmtSubscript(HWND rich) { RichToggleScriptEffect(rich, CFE_SUBSCRIPT, CFE_SUPERSCRIPT); }
static void BodyFmtBullet(HWND rich) { RichToggleBullet(rich); }
static void BodyFmtNumbering(HWND rich) { RichToggleNumbering(rich); }
static void BodyFmtClear(HWND rich) { RichClearFormatting(rich); }
static void BodyFmtIndent(HWND rich, int twips) { RichAdjustIndent(rich, twips); }
static void BodyFmtAlign(HWND rich, int align) { RichSetParaAlignment(rich, (WORD)align); }

static void ApplyMarkdownStylingToBody(AppState* s) {
  if (!s || !s->ed_body) return;
  std::wstring md = GetWindowTextWString(s->ed_body);
  if (md.empty()) return;
  RenderMarkdownToRichEdit(s->ed_body, s->font, md);
  SetEditorDirty(s, true);
  SendMessageW(s->ed_body, EM_SETMODIFY, TRUE, 0);
  UpdateStatusBar(s);
  ShowToast(s, L"已应用 Markdown 样式");
}

static void UpdateStatusParts(AppState* s) {
  if (!s || !s->status) return;
  RECT rc{};
  GetClientRect(s->hwnd, &rc);
  int w = rc.right - rc.left;
  int hint_w = ScalePx(s->hwnd, 520);
  int stats_w = ScalePx(s->hwnd, 240);
  int first = w - hint_w - stats_w;
  if (first < ScalePx(s->hwnd, 200)) first = ScalePx(s->hwnd, 200);
  int second = first + stats_w;
  if (second > w - ScalePx(s->hwnd, 160)) second = w - ScalePx(s->hwnd, 160);
  int parts[3]{first, second, -1};
  SendMessageW(s->status, SB_SETPARTS, 3, (LPARAM)parts);
}

static const wchar_t* StatusToCN(EntryStatus st) {
  switch (st) {
    case EntryStatus::Todo: return L"无";
    case EntryStatus::Doing: return L"进行中";
    case EntryStatus::Blocked: return L"阻塞";
    case EntryStatus::Done: return L"已完成";
    case EntryStatus::None:
    default: return L"无";
  }
}

static EntryStatus StatusFromComboSel(int idx) {
  switch (idx) {
    case 1: return EntryStatus::Doing;
    case 2: return EntryStatus::Blocked;
    case 3: return EntryStatus::Done;
    default: return EntryStatus::None;
  }
}

static int ComboSelFromStatus(EntryStatus st) {
  switch (st) {
    case EntryStatus::Todo: return 0;
    case EntryStatus::Doing: return 1;
    case EntryStatus::Blocked: return 2;
    case EntryStatus::Done: return 3;
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
  if (!s) return;
  s->editor_dirty = dirty;
  // Simple affordance: append "*" to title bar when dirty.
  std::wstring title = kAppTitle;
  if (dirty) title += L" *";
  SetWindowTextW(s->hwnd, title.c_str());

  bool want_autosave = dirty && s->editing_index >= 0;
  if (want_autosave) {
    if (s->autosave_timer_id) KillTimer(s->hwnd, s->autosave_timer_id);
    s->autosave_timer_id = SetTimer(s->hwnd, 3, 3 * 60 * 1000, nullptr);
  } else if (s->autosave_timer_id) {
    KillTimer(s->hwnd, s->autosave_timer_id);
    s->autosave_timer_id = 0;
  }
}

static Task* FindTaskByIdIn(std::vector<Task>* tasks, const std::wstring& id) {
  if (!tasks) return nullptr;
  for (auto& t : *tasks) {
    if (t.id == id) return &t;
  }
  return nullptr;
}

static bool EditorHasMeaningfulContent(AppState* s) {
  if (!s) return false;
  return !EditorMatchesDefaultDraft(s);
}

static bool EditorMatchesDefaultDraft(AppState* s) {
  if (!s) return true;
  std::wstring category = GetWindowTextWString(s->cb_category);
  std::wstring start_time = GetDateCtrlYmd(s->ed_start);
  std::wstring end_time = GetDateCtrlYmd(s->ed_end);
  std::wstring title = GetWindowTextWString(s->ed_title);
  std::wstring body = GetWindowTextWString(s->ed_body);
  std::wstring default_category = s->categories.empty() ? L"" : s->categories[0];
  std::wstring default_start = FormatDateYYYYMMDD(s->selected);
  std::wstring default_end = FormatDateYYYYMMDD(AddDays(s->selected, 1));
  int status_sel = s->cb_status ? (int)SendMessageW(s->cb_status, CB_GETCURSEL, 0, 0) : 0;
  return category == default_category &&
         start_time == default_start &&
         end_time == default_end &&
         status_sel == 0 &&
         title.empty() &&
         body.empty() &&
         s->editor_materials_dir.empty();
}

static bool EditorMatchesEntry(AppState* s, const Entry& e) {
  Entry cur{};
  cur.category = GetWindowTextWString(s->cb_category);
  cur.start_time = GetDateCtrlYmd(s->ed_start);
  cur.end_time = GetDateCtrlYmd(s->ed_end);
  cur.title = GetWindowTextWString(s->ed_title);
  cur.body_plain = GetWindowTextWString(s->ed_body);
  if (s->ed_body && !cur.body_plain.empty()) {
    std::string rtf = RichEditGetRtfBytes(s->ed_body);
    if (!rtf.empty()) cur.body_rtf_b64 = Base64Encode(rtf);
  }
  cur.materials_dir = s->editor_materials_dir;
  return cur.category == e.category && cur.start_time == e.start_time && cur.end_time == e.end_time &&
         cur.title == e.title && cur.body_plain == e.body_plain && cur.body_rtf_b64 == e.body_rtf_b64 &&
         cur.materials_dir == e.materials_dir;
}

static bool IsEditorTrulyDirty(AppState* s) {
  // RichEdit can be modified by formatting changes without changing plain text.
  bool body_modified = false;
  if (s->ed_body) body_modified = (SendMessageW(s->ed_body, EM_GETMODIFY, 0, 0) != 0);

  if (!s->editor_dirty && !body_modified) return false;
  if (s->editing_index >= 0 && s->editing_index < (int)s->day.entries.size()) {
    return !EditorMatchesEntry(s, s->day.entries[(size_t)s->editing_index]);
  }
  // For new entries, avoid nagging when editor still matches the untouched default draft,
  // even if RichEdit or date controls toggled internal modify state.
  return !EditorMatchesDefaultDraft(s);
}

static void ClearEditor(AppState* s, bool allow_blank) {
  if (!s) return;
  if (!allow_blank) {
    int sel = s->list ? ListView_GetNextItem(s->list, -1, LVNI_SELECTED) : -1;
    if (sel >= 0 && sel < (int)s->day.entries.size()) {
      FillEditorFromEntry(s, s->day.entries[(size_t)sel], sel);
      return;
    }
    if (s->editing_index >= 0 && s->editing_index < (int)s->day.entries.size()) {
      FillEditorFromEntry(s, s->day.entries[(size_t)s->editing_index], s->editing_index);
      return;
    }
  }
  s->suppress_editor_change_tracking = true;
  if (!s->categories.empty()) SetWindowTextW(s->cb_category, s->categories[0].c_str());
  else SetWindowTextW(s->cb_category, L"");
  DateTime_SetSystemtime(s->ed_start, GDT_VALID, &s->selected);
  SYSTEMTIME next = AddDays(s->selected, 1);
  DateTime_SetSystemtime(s->ed_end, GDT_VALID, &next);
  if (s->cb_status) SendMessageW(s->cb_status, CB_SETCURSEL, 0, 0);
  SetEditText(s->ed_title, L"");
  SetEditText(s->ed_body, L"");
  ApplyBodyZoom(s);
  s->editor_materials_dir.clear();
  if (s->ed_title) SendMessageW(s->ed_title, EM_SETMODIFY, FALSE, 0);
  if (s->ed_body) SendMessageW(s->ed_body, EM_SETMODIFY, FALSE, 0);
  s->editing_index = -1;
  SetEditorDirty(s, false);
  EnableWindow(s->cb_category, TRUE);
  EnableWindow(s->ed_title, TRUE);
  if (s->cb_status) EnableWindow(s->cb_status, TRUE);
  s->suppress_editor_change_tracking = false;
  UpdateStatusBar(s);
}

static void FillEditorFromEntry(AppState* s, const Entry& e, int index) {
  if (!s) return;
  s->suppress_editor_change_tracking = true;
  bool is_task = EntryIsTaskProgress(e);
  SetWindowTextW(s->cb_category, e.category.c_str());
  std::wstring start_ymd = e.start_time;
  std::wstring end_ymd = e.end_time;
  if (is_task && !e.task_id.empty()) {
    Task* t = FindTaskById(s, e.task_id);
    if (t) {
      start_ymd = FormatDateYYYYMMDD(t->start);
      end_ymd = FormatDateYYYYMMDD(AddDays(t->end, 1));
    }
  }
  SetDateCtrlOrDefault(s->ed_start, start_ymd, s->selected);
  SetDateCtrlOrDefault(s->ed_end, end_ymd, AddDays(s->selected, 1));
  if (s->cb_status) {
    SendMessageW(s->cb_status, CB_SETCURSEL, ComboSelFromStatus(EffectiveStatus(s, e)), 0);
    EnableWindow(s->cb_status, TRUE);
  }
  SetEditText(s->ed_title, e.title);
  if (!e.body_rtf_b64.empty()) {
    std::string rtf;
    if (Base64Decode(e.body_rtf_b64, &rtf)) {
      RichEditSetRtfBytes(s->ed_body, rtf);
      ApplyBodyZoom(s);
    } else {
      SetEditText(s->ed_body, e.body_plain);
      ApplyBodyZoom(s);
    }
  } else {
    SetEditText(s->ed_body, e.body_plain);
    ApplyBodyZoom(s);
  }

  // Materials dir is always editable even if task progress is read-only in other fields.
  // For task progress, it is treated as a property of the underlying Task.
  if (is_task && !e.task_id.empty()) {
    Task* t = FindTaskById(s, e.task_id);
    s->editor_materials_dir = t ? t->materials_dir : L"";
  } else {
    s->editor_materials_dir = e.materials_dir;
  }
  if (s->ed_title) SendMessageW(s->ed_title, EM_SETMODIFY, FALSE, 0);
  if (s->ed_body) SendMessageW(s->ed_body, EM_SETMODIFY, FALSE, 0);
  s->editing_index = index;
  SetEditorDirty(s, false);
  EnableWindow(s->cb_category, TRUE);
  EnableWindow(s->ed_title, TRUE);
  s->suppress_editor_change_tracking = false;
  UpdateStatusBar(s);
}

static std::wstring TimeRangeText(const Entry& e) {
  if (e.start_time.empty() && e.end_time.empty()) return L"";
  if (!e.start_time.empty() && !e.end_time.empty()) return e.start_time + L" ~ " + e.end_time;
  return e.start_time.empty() ? e.end_time : e.start_time;
}

static std::wstring EntryMaterialsDirText(AppState* s, const Entry& e) {
  if (EntryIsTaskProgress(e) && !e.task_id.empty()) {
    Task* t = FindTaskById(s, e.task_id);
    if (t) return TaskMaterialsDirResolved(*t);
    return GetTaskMaterialsDirById(e.task_id);
  }
  if (!e.materials_dir.empty()) return e.materials_dir;
  if (!e.id.empty()) return GetWorkMaterialsDir(s->selected, e.id);
  return L"";
}

static void UpdateListRowForEntry(AppState* s, int index) {
  if (!s || !s->list) return;
  if (index < 0 || index >= (int)s->day.entries.size()) return;
  const auto& e = s->day.entries[(size_t)index];
  ListView_SetItemText(s->list, index, 0, const_cast<wchar_t*>(TimeRangeText(e).c_str()));
  ListView_SetItemText(s->list, index, 1, const_cast<wchar_t*>(e.category.c_str()));
  ListView_SetItemText(s->list, index, 2, const_cast<wchar_t*>(e.title.c_str()));
  EntryStatus st = EffectiveStatus(s, e);
  ListView_SetItemText(s->list, index, 3, const_cast<wchar_t*>(StatusToCN(st)));
  std::wstring dir = EntryMaterialsDirText(s, e);
  ListView_SetItemText(s->list, index, 4, const_cast<wchar_t*>(dir.c_str()));
}

static int CompareTextNoCase(const std::wstring& a, const std::wstring& b) {
  int r = _wcsicmp(a.c_str(), b.c_str());
  if (r < 0) return -1;
  if (r > 0) return 1;
  return 0;
}

static int CompareEntryStartTime(const Entry& a, const Entry& b) {
  bool a_empty = a.start_time.empty();
  bool b_empty = b.start_time.empty();
  if (a_empty != b_empty) return a_empty ? 1 : -1;
  int r = CompareDateYmd(a.start_time, b.start_time);
  if (r != 0) return r;
  r = CompareDateYmd(a.end_time, b.end_time);
  if (r != 0) return r;
  r = CompareTextNoCase(a.title, b.title);
  if (r != 0) return r;
  return CompareTextNoCase(a.id, b.id);
}

static int CompareEntryCategory(const Entry& a, const Entry& b) {
  int r = CompareTextNoCase(a.category, b.category);
  if (r != 0) return r;
  r = CompareEntryStartTime(a, b);
  if (r != 0) return r;
  return CompareTextNoCase(a.id, b.id);
}

static int CompareEntryStatus(AppState* s, const Entry& a, const Entry& b) {
  int sa = (int)EffectiveStatus(s, a);
  int sb = (int)EffectiveStatus(s, b);
  if (sa != sb) return sa < sb ? -1 : 1;
  int r = CompareEntryStartTime(a, b);
  if (r != 0) return r;
  return CompareTextNoCase(a.id, b.id);
}

static void SortDayEntries(AppState* s) {
  if (!s) return;

  std::wstring editing_id;
  if (s->editing_index >= 0 && s->editing_index < (int)s->day.entries.size()) {
    editing_id = s->day.entries[(size_t)s->editing_index].id;
  }

  auto cmp = [&](const Entry& a, const Entry& b) {
    int r = 0;
    switch (s->list_sort_column) {
      case 1:
        r = CompareEntryCategory(a, b);
        break;
      case 3:
        r = CompareEntryStatus(s, a, b);
        break;
      case 0:
      default:
        r = CompareEntryStartTime(a, b);
        break;
    }
    if (s->list_sort_desc) r = -r;
    return r < 0;
  };
  std::stable_sort(s->day.entries.begin(), s->day.entries.end(), cmp);

  if (!editing_id.empty()) {
    for (int i = 0; i < (int)s->day.entries.size(); i++) {
      if (s->day.entries[(size_t)i].id == editing_id) {
        s->editing_index = i;
        break;
      }
    }
  }
}

static void RefreshList(AppState* s) {
  if (!s || !s->list) return;
  int prev_sel = ListView_GetNextItem(s->list, -1, LVNI_SELECTED);
  std::wstring prev_id;
  if (prev_sel >= 0 && prev_sel < (int)s->day.entries.size()) {
    prev_id = s->day.entries[(size_t)prev_sel].id;
  }

  SortDayEntries(s);

  s->suppress_list_selection_sync = true;
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
    std::wstring dir = EntryMaterialsDirText(s, e);
    ListView_SetItemText(s->list, i, 4, const_cast<wchar_t*>(dir.c_str()));
  }

  int restore_index = -1;
  if (!prev_id.empty()) {
    for (int i = 0; i < (int)s->day.entries.size(); i++) {
      if (s->day.entries[(size_t)i].id == prev_id) {
        restore_index = i;
        break;
      }
    }
  } else if (prev_sel >= 0 && prev_sel < (int)s->day.entries.size()) {
    restore_index = prev_sel;
  }
  if (restore_index >= 0) {
    ListView_SetItemState(s->list, restore_index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_EnsureVisible(s->list, restore_index, FALSE);
  }
  s->suppress_list_selection_sync = false;

  UpdateDayHeader(s);
}

static bool LoadSelectedDay(AppState* s, const SYSTEMTIME& date) {
  std::wstring restore_entry_id;
  bool same_day_reload = CompareSystemTimeDateOnly(s->selected, date) == 0;
  if (same_day_reload && s->editing_index >= 0 && s->editing_index < (int)s->day.entries.size()) {
    restore_entry_id = s->day.entries[(size_t)s->editing_index].id;
  }

  DayData dd{};
  std::wstring err;
  if (!LoadVisibleEntriesForDate(date, &dd, &err)) {
    ShowInfoBox(s->hwnd, err.c_str(), L"读取失败");
    return false;
  }
  s->selected = date;
  s->day = dd;

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
      e.start_time = FormatDateYYYYMMDD(t.start);
      e.end_time = FormatDateYYYYMMDD(AddDays(t.end, 1));
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
  if (!restore_entry_id.empty()) {
    for (int i = 0; i < (int)s->day.entries.size(); ++i) {
      if (s->day.entries[(size_t)i].id == restore_entry_id) {
        s->editing_index = i;
        ListView_SetItemState(s->list, i, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
        ListView_EnsureVisible(s->list, i, FALSE);
        FillEditorFromEntry(s, s->day.entries[(size_t)i], i);
        UpdateDayHeader(s);
        UpdateStatusBar(s);
        return true;
      }
    }
  }
  ClearEditor(s, true);
  UpdateDayHeader(s);
  UpdateStatusBar(s);
  return true;
}

static int CompareDateYmd(const std::wstring& a, const std::wstring& b) {
  SYSTEMTIME sa{};
  SYSTEMTIME sb{};
  if (!ParseYYYYMMDD(a, &sa) || !ParseYYYYMMDD(b, &sb)) return 0;
  if (sa.wYear != sb.wYear) return sa.wYear < sb.wYear ? -1 : 1;
  if (sa.wMonth != sb.wMonth) return sa.wMonth < sb.wMonth ? -1 : 1;
  if (sa.wDay != sb.wDay) return sa.wDay < sb.wDay ? -1 : 1;
  return 0;
}

static bool ResolveEntrySourceDate(const Entry& e, const SYSTEMTIME& fallback_date, SYSTEMTIME* out) {
  if (!out) return false;
  if (!e.source_day_ymd.empty() && ParseYYYYMMDD(e.source_day_ymd, out)) return true;
  *out = fallback_date;
  return true;
}

static bool EndsWithI(const std::wstring& s, const std::wstring& suffix) {
  if (s.size() < suffix.size()) return false;
  size_t off = s.size() - suffix.size();
  for (size_t i = 0; i < suffix.size(); ++i) {
    if (towlower(s[off + i]) != towlower(suffix[i])) return false;
  }
  return true;
}

static bool EntryVisibleOnDate(const Entry& e, const SYSTEMTIME& selected, const SYSTEMTIME* source_date) {
  std::wstring selected_ymd = FormatDateYYYYMMDD(selected);
  SYSTEMTIME start{};
  SYSTEMTIME end{};
  if (!e.start_time.empty() && !e.end_time.empty() && ParseYYYYMMDD(e.start_time, &start) && ParseYYYYMMDD(e.end_time, &end)) {
    return CompareDateYmd(e.start_time, selected_ymd) <= 0 && CompareDateYmd(selected_ymd, e.end_time) < 0;
  }

  SYSTEMTIME source{};
  if (source_date) {
    source = *source_date;
    return CompareSystemTimeDateOnly(source, selected) == 0;
  }
  if (ResolveEntrySourceDate(e, selected, &source)) {
    return CompareSystemTimeDateOnly(source, selected) == 0;
  }
  return false;
}

static bool EnumerateStoredDayDates(std::vector<SYSTEMTIME>* out, std::wstring* err) {
  if (!out) return false;
  out->clear();

  std::set<std::wstring> seen;
  std::wstring root = GetDataRootDir();
  WIN32_FIND_DATAW yfd{};
  HANDLE hy = FindFirstFileW(JoinPathLoose(root, L"*").c_str(), &yfd);
  if (hy == INVALID_HANDLE_VALUE) return true;

  do {
    if (!(yfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) continue;
    std::wstring year = yfd.cFileName;
    if (year == L"." || year == L"..") continue;
    if (year.size() != 4 || year.find_first_not_of(L"0123456789") != std::wstring::npos) continue;

    std::wstring year_dir = JoinPathLoose(root, year);
    WIN32_FIND_DATAW mfd{};
    HANDLE hm = FindFirstFileW(JoinPathLoose(year_dir, L"*").c_str(), &mfd);
    if (hm == INVALID_HANDLE_VALUE) continue;
    do {
      if (!(mfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) continue;
      std::wstring month = mfd.cFileName;
      if (month == L"." || month == L"..") continue;
      if (month.size() != 2 || month.find_first_not_of(L"0123456789") != std::wstring::npos) continue;

      std::wstring month_dir = JoinPathLoose(year_dir, month);
      WIN32_FIND_DATAW ffd{};
      HANDLE hf = FindFirstFileW(JoinPathLoose(month_dir, L"*").c_str(), &ffd);
      if (hf == INVALID_HANDLE_VALUE) continue;
      do {
        if (ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;
        std::wstring name = ffd.cFileName;
        if (name.size() != 14 && name.size() != 15) continue;
        if (!(EndsWithI(name, L".wlr") || EndsWithI(name, L".wlmd"))) continue;
        std::wstring ymd = name.substr(0, 10);
        if (!seen.insert(ymd).second) continue;
        SYSTEMTIME st{};
        if (!ParseYYYYMMDD(ymd, &st)) continue;
        out->push_back(st);
      } while (FindNextFileW(hf, &ffd));
      FindClose(hf);
    } while (FindNextFileW(hm, &mfd));
    FindClose(hm);
  } while (FindNextFileW(hy, &yfd));
  FindClose(hy);

  std::sort(out->begin(), out->end(), [](const SYSTEMTIME& a, const SYSTEMTIME& b) {
    return CompareSystemTimeDateOnly(a, b) < 0;
  });
  return true;
}

static bool LoadVisibleEntriesForDate(const SYSTEMTIME& selected, DayData* out, std::wstring* err) {
  if (!out) return false;
  out->date = selected;
  out->entries.clear();

  std::vector<SYSTEMTIME> stored_dates;
  if (!EnumerateStoredDayDates(&stored_dates, err)) return false;
  for (const auto& source_date : stored_dates) {
    DayData day{};
    std::wstring load_err;
    if (!LoadDayFile(source_date, &day, &load_err)) {
      if (err) *err = load_err;
      return false;
    }
    std::wstring source_ymd = FormatDateYYYYMMDD(source_date);
    for (auto e : day.entries) {
      e.source_day_ymd = source_ymd;
      if (EntryVisibleOnDate(e, selected, &source_date)) {
        out->entries.push_back(std::move(e));
      }
    }
  }
  return true;
}

static bool SaveEntryToBackingStore(const Entry& e, const SYSTEMTIME& fallback_date, std::wstring* err) {
  SYSTEMTIME source_date{};
  ResolveEntrySourceDate(e, fallback_date, &source_date);

  DayData stored{};
  if (!LoadDayFile(source_date, &stored, err)) return false;

  Entry persisted = e;
  persisted.source_day_ymd.clear();
  bool found = false;
  for (auto& existing : stored.entries) {
    if (existing.id == persisted.id) {
      existing = persisted;
      found = true;
      break;
    }
  }
  if (!found) stored.entries.push_back(std::move(persisted));
  stored.date = source_date;
  return SaveDayFile(stored, err);
}

static bool DeleteEntriesFromBackingStore(const std::vector<Entry>& entries, const SYSTEMTIME& fallback_date, std::wstring* err) {
  std::map<std::wstring, std::vector<std::wstring>> ids_by_day;
  std::map<std::wstring, DayData> original_days;
  std::vector<std::wstring> commit_order;

  for (const auto& e : entries) {
    if (e.placeholder && e.source_day_ymd.empty()) continue;
    SYSTEMTIME source_date{};
    ResolveEntrySourceDate(e, fallback_date, &source_date);
    std::wstring source_ymd = FormatDateYYYYMMDD(source_date);
    ids_by_day[source_ymd].push_back(e.id);
  }

  for (const auto& kv : ids_by_day) {
    SYSTEMTIME source_date{};
    if (!ParseYYYYMMDD(kv.first, &source_date)) continue;
    DayData day{};
    if (!LoadDayFile(source_date, &day, err)) {
      for (auto it = commit_order.rbegin(); it != commit_order.rend(); ++it) {
        std::wstring rollback_err;
        (void)SaveDayFile(original_days[*it], &rollback_err);
      }
      return false;
    }
    original_days[kv.first] = day;
    auto& ids = kv.second;
    day.entries.erase(std::remove_if(day.entries.begin(), day.entries.end(), [&](const Entry& item) {
                      return std::find(ids.begin(), ids.end(), item.id) != ids.end();
                    }),
                    day.entries.end());

    bool ok = true;
    if (!day.entries.empty()) {
      ok = SaveDayFile(day, err);
    } else {
      std::wstring wlr_path = GetDayFilePath(source_date);
      std::wstring wlmd_path = JoinPathLoose(ParentDirOf(wlr_path), FormatDateYYYYMMDD(source_date) + L".wlmd");
      BOOL del_wlr = DeleteFileW(wlr_path.c_str());
      DWORD le_wlr = del_wlr ? ERROR_SUCCESS : GetLastError();
      BOOL del_wlmd = DeleteFileW(wlmd_path.c_str());
      DWORD le_wlmd = del_wlmd ? ERROR_SUCCESS : GetLastError();
      ok = ((del_wlr || le_wlr == ERROR_FILE_NOT_FOUND || le_wlr == ERROR_PATH_NOT_FOUND) &&
            (del_wlmd || le_wlmd == ERROR_FILE_NOT_FOUND || le_wlmd == ERROR_PATH_NOT_FOUND));
      if (!ok && err) {
        *err = L"删除日记录文件失败：\n" + wlr_path + L"\n\n" + FormatWin32ErrorMessage(
          (le_wlr != ERROR_SUCCESS && le_wlr != ERROR_FILE_NOT_FOUND && le_wlr != ERROR_PATH_NOT_FOUND) ? le_wlr : le_wlmd);
      }
    }

    if (!ok) {
      auto restore_current = original_days.find(kv.first);
      if (restore_current != original_days.end()) {
        std::wstring rollback_err;
        (void)SaveDayFile(restore_current->second, &rollback_err);
      }
      for (auto it = commit_order.rbegin(); it != commit_order.rend(); ++it) {
        std::wstring rollback_err;
        (void)SaveDayFile(original_days[*it], &rollback_err);
      }
      return false;
    }
    commit_order.push_back(kv.first);
  }
  return true;
}

static bool SaveEditorToModel(AppState* s, bool* out_changed) {
  if (out_changed) *out_changed = false;
  Entry e{};
  if (s->editing_index >= 0 && s->editing_index < (int)s->day.entries.size()) {
    e = s->day.entries[(size_t)s->editing_index];
  }
  e.category = GetWindowTextWString(s->cb_category);
  e.start_time = GetDateCtrlYmd(s->ed_start);
  e.end_time = GetDateCtrlYmd(s->ed_end);
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

  // Status: normal entries store it directly; task progress keeps using the underlying task status.
  bool task_status_changed = false;
  bool task_meta_changed = false;
  if (EntryIsTaskProgress(e) && s->cb_status) {
    int idx = (int)SendMessageW(s->cb_status, CB_GETCURSEL, 0, 0);
    EntryStatus chosen = StatusFromComboSel(idx);
    Task* t = FindTaskById(s, e.task_id);
    if (t) {
      e.start_time = FormatDateYYYYMMDD(t->start);
      e.end_time = FormatDateYYYYMMDD(AddDays(t->end, 1));
      if (t->status != chosen) {
        t->status = chosen;
        task_status_changed = true;
      }
      if (t->category != e.category) {
        t->category = e.category;
        task_meta_changed = true;
      }
      if (t->title != e.title) {
        t->title = e.title;
        task_meta_changed = true;
      }
    }
    // Keep per-day entry status empty; list uses EffectiveStatus(task.status) when entry.status == None.
    e.status = EntryStatus::None;
  } else {
    int idx = s->cb_status ? (int)SendMessageW(s->cb_status, CB_GETCURSEL, 0, 0) : 0;
    e.status = StatusFromComboSel(idx);
  }

  // Materials dir: for task progress it is treated as a property of the Task, not the per-day entry.
  if (EntryIsTaskProgress(e)) {
    e.materials_dir.clear();
  } else {
    e.materials_dir = s->editor_materials_dir;
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

  if (e.start_time.empty() || e.end_time.empty()) {
    ShowInfoBox(s->hwnd, L"请选择开始日期和结束日期。", L"日期范围");
    SetFocus(e.start_time.empty() ? s->ed_start : s->ed_end);
    return false;
  }
  if (CompareDateYmd(e.start_time, e.end_time) >= 0) {
    ShowInfoBox(s->hwnd, L"结束日期必须晚于开始日期。", L"日期范围");
    SetFocus(s->ed_end);
    return false;
  }

  bool has_any = false;
  if (EntryIsTaskProgress(e)) {
    // Task progress placeholders are injected into the day view. Autosave should not
    // materialize them as persisted daily records unless the user actually wrote progress.
    // Task title/category/status are task-level data and are saved separately to tasks.wlt.
    has_any = !e.body_plain.empty() || !e.body_rtf_b64.empty();
  } else {
    has_any = !e.title.empty() || !e.body_plain.empty() || e.status != EntryStatus::None || !e.materials_dir.empty();
  }
  if (!has_any) {
    // Treat as "nothing to save": allow navigation without blocking.
    if (out_changed) *out_changed = task_status_changed || task_meta_changed;
    if (task_status_changed || task_meta_changed) {
      std::wstring terr;
      SaveTasks(s->tasks, &terr);
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
    e.source_day_ymd = FormatDateYYYYMMDD(s->selected);
    s->day.entries.push_back(e);
    s->editing_index = (int)s->day.entries.size() - 1;
  }
  if (out_changed) *out_changed = true;

  // Once user saves, this entry should persist.
  s->day.entries[(size_t)s->editing_index].placeholder = false;
  if (s->day.entries[(size_t)s->editing_index].source_day_ymd.empty()) {
    s->day.entries[(size_t)s->editing_index].source_day_ymd = FormatDateYYYYMMDD(s->selected);
  }
  EnsureMaterialsDirForEntry(s, s->day.entries[(size_t)s->editing_index]);

  // Add category to list if new, and persist categories best-effort.
  if (!CategoriesContains(s->categories, e.category)) {
    std::vector<std::wstring> new_categories = s->categories;
    new_categories.push_back(e.category);
    NormalizeCategories(&new_categories);
    std::wstring cat_err;
    if (!SaveCategories(new_categories, &cat_err)) {
      ShowInfoBox(s->hwnd, cat_err.c_str(), L"保存分类失败");
      return false;
    }
    s->categories = std::move(new_categories);
    RefreshCategoryCombo(s);
  }

  if (task_status_changed || task_meta_changed) {
    std::wstring terr;
    if (!SaveTasks(s->tasks, &terr)) {
      ShowInfoBox(s->hwnd, terr.c_str(), L"保存任务失败");
      return false;
    }
  }
  return true;
}

static bool SaveCurrent(AppState* s, bool rebuild_list = true) {
  // Saving can trigger programmatic list refresh/selection changes which would otherwise re-enter
  // PromptSaveIfDirty and create a prompt loop. Suppress prompts during the save operation.
  struct Guard {
    AppState* s{};
    bool old{};
    ~Guard() {
      if (s) s->suppress_dirty_prompt = old;
    }
  } guard{s, s ? s->suppress_dirty_prompt : false};
  if (s) s->suppress_dirty_prompt = true;

  bool changed = false;
  if (!SaveEditorToModel(s, &changed)) return false;
  if (!changed) {
    SetEditorDirty(s, false);
    if (s->ed_title) SendMessageW(s->ed_title, EM_SETMODIFY, FALSE, 0);
    if (s->ed_body) SendMessageW(s->ed_body, EM_SETMODIFY, FALSE, 0);
    UpdateStatusBar(s);
    return true;
  }
  std::wstring err;
  if (s->editing_index < 0 || s->editing_index >= (int)s->day.entries.size()) return false;
  if (!SaveEntryToBackingStore(s->day.entries[(size_t)s->editing_index], s->selected, &err)) {
    ShowInfoBox(s->hwnd, err.c_str(), L"保存失败");
    return false;
  }

  // Clear dirty flags before any UI churn (RefreshList/selection changes) to avoid prompt loops.
  SetEditorDirty(s, false);
  if (s->ed_title) SendMessageW(s->ed_title, EM_SETMODIFY, FALSE, 0);
  if (s->ed_body) SendMessageW(s->ed_body, EM_SETMODIFY, FALSE, 0);

  if (rebuild_list) {
    (void)LoadSelectedDay(s, s->selected);
  } else if (s->editing_index >= 0 && s->editing_index < (int)s->day.entries.size()) {
    UpdateListRowForEntry(s, s->editing_index);
  }
  SetFocus(s->ed_body ? s->ed_body : s->ed_title);
  ShowToast(s, L"已保存");
  UpdateStatusBar(s);
  return true;
}

static bool AutoSaveCurrent(AppState* s) {
  if (!s) return false;
  if (s->editing_index < 0 || s->editing_index >= (int)s->day.entries.size()) return false;

  Entry original_entry = s->day.entries[(size_t)s->editing_index];
  std::vector<Task> tasks_copy = s->tasks;
  Entry e = original_entry;

  e.category = GetWindowTextWString(s->cb_category);
  e.start_time = GetDateCtrlYmd(s->ed_start);
  e.end_time = GetDateCtrlYmd(s->ed_end);
  e.title = GetWindowTextWString(s->ed_title);
  e.body_plain = GetWindowTextWString(s->ed_body);
  if (s->ed_body) {
    if (!e.body_plain.empty()) {
      std::string rtf = RichEditGetRtfBytes(s->ed_body);
      if (!rtf.empty()) e.body_rtf_b64 = Base64Encode(rtf);
    } else {
      e.body_rtf_b64.clear();
    }
  }

  while (!e.category.empty() && iswspace(e.category.front())) e.category.erase(e.category.begin());
  while (!e.category.empty() && iswspace(e.category.back())) e.category.pop_back();
  if (e.category.empty()) return false;
  if (e.start_time.empty() || e.end_time.empty()) return false;
  if (CompareDateYmd(e.start_time, e.end_time) >= 0) return false;

  bool task_status_changed = false;
  bool task_meta_changed = false;
  if (EntryIsTaskProgress(e) && s->cb_status) {
    int idx = (int)SendMessageW(s->cb_status, CB_GETCURSEL, 0, 0);
    EntryStatus chosen = StatusFromComboSel(idx);
    Task* t = FindTaskByIdIn(&tasks_copy, e.task_id);
    if (t) {
      e.start_time = FormatDateYYYYMMDD(t->start);
      e.end_time = FormatDateYYYYMMDD(AddDays(t->end, 1));
      if (t->status != chosen) {
        t->status = chosen;
        task_status_changed = true;
      }
      if (t->category != e.category) {
        t->category = e.category;
        task_meta_changed = true;
      }
      if (t->title != e.title) {
        t->title = e.title;
        task_meta_changed = true;
      }
    }
    e.status = EntryStatus::None;
    e.materials_dir.clear();
  } else {
    int idx = s->cb_status ? (int)SendMessageW(s->cb_status, CB_GETCURSEL, 0, 0) : 0;
    e.status = StatusFromComboSel(idx);
    e.materials_dir = s->editor_materials_dir;
  }

  bool has_any = false;
  if (EntryIsTaskProgress(e)) {
    has_any = !e.body_plain.empty() || !e.body_rtf_b64.empty();
  } else {
    has_any = !e.title.empty() || !e.body_plain.empty() || e.status != EntryStatus::None || !e.materials_dir.empty();
  }

  bool day_written = false;
  if (has_any) {
    if (e.source_day_ymd.empty()) e.source_day_ymd = FormatDateYYYYMMDD(s->selected);
    e.placeholder = false;
    std::wstring err;
    if (!SaveEntryToBackingStore(e, s->selected, &err)) return false;
    day_written = true;
  }

  if (task_status_changed || task_meta_changed) {
    std::wstring terr;
    if (!SaveTasks(tasks_copy, &terr)) {
      if (day_written) {
        std::wstring rollback_err;
        (void)SaveEntryToBackingStore(original_entry, s->selected, &rollback_err);
      }
      return false;
    }
  }

  s->tasks = std::move(tasks_copy);
  if (has_any) s->day.entries[(size_t)s->editing_index] = e;

  if (!CategoriesContains(s->categories, e.category)) {
    std::vector<std::wstring> new_categories = s->categories;
    new_categories.push_back(e.category);
    NormalizeCategories(&new_categories);
    std::wstring cat_err;
    if (!SaveCategories(new_categories, &cat_err)) return false;
    s->categories = std::move(new_categories);
    RefreshCategoryCombo(s);
  }

  SetEditorDirty(s, false);
  if (s->ed_title) SendMessageW(s->ed_title, EM_SETMODIFY, FALSE, 0);
  if (s->ed_body) SendMessageW(s->ed_body, EM_SETMODIFY, FALSE, 0);
  if (has_any) {
    if (s->list_sort_column == 0 || s->list_sort_column == 1 || s->list_sort_column == 3) {
      RefreshList(s);
      if (s->editing_index >= 0 && s->editing_index < (int)s->day.entries.size()) {
        ListView_SetItemState(s->list, s->editing_index, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
        ListView_EnsureVisible(s->list, s->editing_index, FALSE);
      }
    } else {
      UpdateListRowForEntry(s, s->editing_index);
    }
  }
  UpdateStatusBar(s);
  return true;
}

static void DeleteCurrent(AppState* s) {
  if (!s) return;
  std::vector<int> selected = GetSelectedEntryIndices(s);
  if (selected.empty()) {
    if (s->editing_index >= 0 && s->editing_index < (int)s->day.entries.size()) selected.push_back(s->editing_index);
  }
  if (selected.empty()) return;

  std::sort(selected.begin(), selected.end());
  selected.erase(std::unique(selected.begin(), selected.end()), selected.end());
  int next_idx = selected.front();
  DayData original_day = s->day;
  std::vector<Entry> entries_to_delete;
  entries_to_delete.reserve(selected.size());
  for (int idx : selected) {
    if (idx >= 0 && idx < (int)s->day.entries.size()) entries_to_delete.push_back(s->day.entries[(size_t)idx]);
  }

  for (auto it = selected.rbegin(); it != selected.rend(); ++it) {
    s->day.entries.erase(s->day.entries.begin() + *it);
  }

  std::wstring err;
  if (!DeleteEntriesFromBackingStore(entries_to_delete, s->selected, &err)) {
    s->day = std::move(original_day);
    ShowInfoBox(s->hwnd, err.c_str(), L"删除失败");
    return;
  }

  (void)LoadSelectedDay(s, s->selected);
  if (!s->day.entries.empty()) {
    if (next_idx >= (int)s->day.entries.size()) next_idx = (int)s->day.entries.size() - 1;
    s->editing_index = next_idx;
    ListView_SetItemState(s->list, next_idx, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_EnsureVisible(s->list, next_idx, FALSE);
    FillEditorFromEntry(s, s->day.entries[(size_t)next_idx], next_idx);
    SetFocus(s->ed_body ? s->ed_body : s->ed_title);
  } else {
    ClearEditor(s, true);
    HWND h = GetComboEditHandle(s->cb_category);
    SetFocus(h ? h : s->cb_category);
  }
  ShowToast(s, selected.size() > 1 ? L"已删除选中条目" : L"已删除条目");
}

static void DiscardEditorChanges(AppState* s) {
  if (!s) return;
  // Restore editor UI to match persisted model, so user choosing "No" won't keep a dirty editor
  // and trigger the same prompt repeatedly on the next navigation.
  if (s->editing_index >= 0 && s->editing_index < (int)s->day.entries.size()) {
    FillEditorFromEntry(s, s->day.entries[(size_t)s->editing_index], s->editing_index);
  } else {
    ClearEditor(s, true);
  }
  SetEditorDirty(s, false);
  if (s->ed_title) SendMessageW(s->ed_title, EM_SETMODIFY, FALSE, 0);
  if (s->ed_body) SendMessageW(s->ed_body, EM_SETMODIFY, FALSE, 0);
  UpdateStatusBar(s);
}

static bool PromptSaveIfDirty(AppState* s) {
  if (s && s->suppress_dirty_prompt) return true;
  if (!IsEditorTrulyDirty(s)) return true;
  if (s && !EditorHasMeaningfulContent(s)) {
    DiscardEditorChanges(s);
    return true;
  }
  if (s) {
    return SaveCurrent(s);
  }
  return true;
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

static bool ReadFileUtf8Simple(const std::wstring& path, std::wstring* out, std::wstring* err) {
  if (!out) return false;
  out->clear();
  FILE* f = nullptr;
  _wfopen_s(&f, path.c_str(), L"rb");
  if (!f) {
    if (err) *err = L"无法读取文件: " + path;
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

  // Strip UTF-8 BOM if present.
  if (bytes.size() >= 3 && (unsigned char)bytes[0] == 0xEF && (unsigned char)bytes[1] == 0xBB &&
      (unsigned char)bytes[2] == 0xBF) {
    bytes.erase(0, 3);
  }
  *out = Utf8ToWide(bytes);
  return true;
}

static bool ParseKVLine(const std::wstring& line, std::wstring* out_k, std::wstring* out_v) {
  size_t eq = line.find(L'=');
  if (eq == std::wstring::npos) return false;
  if (out_k) *out_k = line.substr(0, eq);
  if (out_v) *out_v = line.substr(eq + 1);
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

static bool ReadClipboardText(HWND hwnd, std::wstring* out) {
  if (out) out->clear();
  if (!OpenClipboard(hwnd)) return false;
  HANDLE h = GetClipboardData(CF_UNICODETEXT);
  if (!h) {
    CloseClipboard();
    return false;
  }
  const wchar_t* p = reinterpret_cast<const wchar_t*>(GlobalLock(h));
  if (!p) {
    CloseClipboard();
    return false;
  }
  if (out) *out = p;
  GlobalUnlock(h);
  CloseClipboard();
  return true;
}

static HMODULE TryLoadRichEditMsftedit() {
  // Security hardening: remove current directory from DLL search path (mitigates DLL hijacking).
  // This is a process-wide setting and is safe for this app (we only load system DLLs).
  SetDllDirectoryW(L"");

  // Prefer loading msftedit from System32 when supported, fallback to default behavior otherwise.
  const DWORD kLoadSystem32 = 0x00000800;  // LOAD_LIBRARY_SEARCH_SYSTEM32
  HMODULE h = LoadLibraryExW(L"Msftedit.dll", nullptr, kLoadSystem32);
  if (h) return h;
  return LoadLibraryW(L"Msftedit.dll");
}

static Entry GetEditorEntrySnapshot(AppState* s) {
  Entry e{};
  e.category = GetWindowTextWString(s->cb_category);
  e.start_time = GetDateCtrlYmd(s->ed_start);
  e.end_time = GetDateCtrlYmd(s->ed_end);
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
    md += TimeRangeText(e) + L" ";
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

static std::wstring NowTimestampCompact() {
  SYSTEMTIME st{};
  GetLocalTime(&st);
  wchar_t buf[64]{};
  wsprintfW(buf, L"%04d%02d%02d_%02d%02d%02d",
            (int)st.wYear, (int)st.wMonth, (int)st.wDay, (int)st.wHour, (int)st.wMinute, (int)st.wSecond);
  return std::wstring(buf);
}

static std::wstring ParentDirOf(const std::wstring& path) {
  if (path.empty()) return L"";
  std::wstring p = path;
  while (!p.empty() && (p.back() == L'\\' || p.back() == L'/')) p.pop_back();
  size_t pos = p.find_last_of(L"\\/");
  if (pos == std::wstring::npos) return L"";
  return p.substr(0, pos);
}

static bool DirExistsW(const std::wstring& path) {
  DWORD attr = GetFileAttributesW(path.c_str());
  return (attr != INVALID_FILE_ATTRIBUTES) && ((attr & FILE_ATTRIBUTE_DIRECTORY) != 0);
}

static bool IsUncPath(const std::wstring& path) {
  return path.size() >= 2 && ((path[0] == L'\\' && path[1] == L'\\') || (path[0] == L'/' && path[1] == L'/'));
}

static std::wstring NormalizeDirForCompare(const std::wstring& in) {
  // Minimal normalization for "is child" checks: trim trailing slashes and lowercase ASCII.
  std::wstring p = in;
  while (!p.empty() && (p.back() == L'\\' || p.back() == L'/')) p.pop_back();
  for (auto& ch : p) {
    if (ch >= L'A' && ch <= L'Z') ch = (wchar_t)(ch - L'A' + L'a');
    if (ch == L'/') ch = L'\\';
  }
  return p;
}

static bool IsSameOrChildDir(const std::wstring& parent, const std::wstring& child) {
  std::wstring p = NormalizeDirForCompare(parent);
  std::wstring c = NormalizeDirForCompare(child);
  if (p.empty() || c.empty()) return false;
  if (c == p) return true;
  if (c.size() <= p.size()) return false;
  if (c.compare(0, p.size(), p) != 0) return false;
  return c[p.size()] == L'\\';
}

static std::wstring JoinPathLoose(const std::wstring& a, const std::wstring& b) {
  // Avoid MAX_PATH pitfalls from PathCombine; JoinPath in win32_util uses it and can truncate for deep trees.
  if (a.empty()) return b;
  if (b.empty()) return a;
  if (a.back() == L'\\' || a.back() == L'/') return a + b;
  return a + L"\\" + b;
}

static std::wstring FormatWin32ErrorMessage(DWORD err) {
  wchar_t* buf = nullptr;
  DWORD flags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
  DWORD len = FormatMessageW(flags, nullptr, err, 0, (LPWSTR)&buf, 0, nullptr);
  std::wstring msg;
  if (len && buf) msg.assign(buf, buf + len);
  if (buf) LocalFree(buf);
  // Trim trailing newlines
  while (!msg.empty() && (msg.back() == L'\r' || msg.back() == L'\n')) msg.pop_back();
  return msg;
}

static std::wstring NormalizeNewlinesToCrlf(const std::wstring& s) {
  // Windows EDIT control and many clipboard consumers expect CRLF line endings.
  std::wstring out;
  out.reserve(s.size() + 16);
  wchar_t prev = 0;
  for (wchar_t ch : s) {
    if (ch == L'\n' && prev != L'\r') out.push_back(L'\r');
    out.push_back(ch);
    prev = ch;
  }
  return out;
}

static bool PickFolderDialog(HWND owner, const wchar_t* title, std::wstring* out_path) {
  if (out_path) out_path->clear();

  HRESULT hr_init = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
  bool do_uninit = SUCCEEDED(hr_init);

  IFileDialog* dlg = nullptr;
  HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&dlg));
  if (FAILED(hr) || !dlg) {
    if (do_uninit) CoUninitialize();
    return false;
  }

  DWORD opts = 0;
  dlg->GetOptions(&opts);
  dlg->SetOptions(opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST);
  if (title) dlg->SetTitle(title);

  hr = dlg->Show(owner);
  if (hr == HRESULT_FROM_WIN32(ERROR_CANCELLED)) {
    dlg->Release();
    if (do_uninit) CoUninitialize();
    return false;
  }
  if (FAILED(hr)) {
    dlg->Release();
    if (do_uninit) CoUninitialize();
    return false;
  }

  IShellItem* item = nullptr;
  hr = dlg->GetResult(&item);
  if (FAILED(hr) || !item) {
    dlg->Release();
    if (do_uninit) CoUninitialize();
    return false;
  }

  PWSTR psz = nullptr;
  hr = item->GetDisplayName(SIGDN_FILESYSPATH, &psz);
  if (SUCCEEDED(hr) && psz && out_path) *out_path = psz;
  if (psz) CoTaskMemFree(psz);

  item->Release();
  dlg->Release();
  if (do_uninit) CoUninitialize();
  return out_path && !out_path->empty();
}

static bool CopyDirRecursive(const std::wstring& src_dir, const std::wstring& dst_dir, std::wstring* err) {
  if (err) err->clear();
  if (!DirExistsW(src_dir)) {
    if (err) *err = L"Source folder not found: " + src_dir;
    return false;
  }
  if (!EnsureDirExists(dst_dir)) {
    if (err) *err = L"Failed to create folder: " + dst_dir;
    return false;
  }

  std::wstring pattern = JoinPathLoose(src_dir, L"*");
  WIN32_FIND_DATAW fd{};
  HANDLE h = FindFirstFileW(pattern.c_str(), &fd);
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

      // Safety: don't follow reparse points (junctions/symlinks). Data dir is expected to be plain files/folders.
      if ((fd.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0) {
        continue;
      }

      if (is_dir) {
        if (!CopyDirRecursive(sp, dp, err)) {
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

static std::wstring ResolveImportDataRoot(const std::wstring& picked_folder) {
  // Accept either:
  // - a direct data root (contains categories/tasks/... or YYYY/MM/day files)
  // - a wrapper folder that contains a "data" subfolder (common when users zip/copy the whole app folder)
  auto looks_like_data_root = [&](const std::wstring& dir) -> bool {
    if (!DirExistsW(dir)) return false;
    if (FileExists(JoinPathLoose(dir, L"categories.txt"))) return true;
    if (FileExists(JoinPathLoose(dir, L"tasks.wlt"))) return true;
    if (FileExists(JoinPathLoose(dir, L"recurring_meetings.wlrp"))) return true;
    if (FileExists(JoinPathLoose(dir, L"auth.wla"))) return true;

    // Heuristic: any 4-digit year directory.
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
  };

  if (looks_like_data_root(picked_folder)) return picked_folder;
  std::wstring sub = JoinPathLoose(picked_folder, L"data");
  if (looks_like_data_root(sub)) return sub;
  return L"";
}

static bool MoveDirReplace(const std::wstring& from, const std::wstring& to, std::wstring* err) {
  if (err) err->clear();
  if (MoveFileExW(from.c_str(), to.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) != 0) return true;
  DWORD le = GetLastError();
  if (err) *err = L"Move failed:\n" + from + L"\n->\n" + to + L"\n(error=" + std::to_wstring(le) + L" " + FormatWin32ErrorMessage(le) + L")";
  return false;
}

static bool MoveDir(const std::wstring& from, const std::wstring& to, std::wstring* err) {
  if (err) err->clear();
  if (MoveFileExW(from.c_str(), to.c_str(), MOVEFILE_WRITE_THROUGH) != 0) return true;
  DWORD le = GetLastError();
  if (err) *err = L"Move failed:\n" + from + L"\n->\n" + to + L"\n(error=" + std::to_wstring(le) + L" " + FormatWin32ErrorMessage(le) + L")";
  return false;
}

static bool ReloadAppDataAfterFilesystemChange(AppState* s) {
  if (!s) return false;
  std::vector<std::wstring> categories;
  std::wstring cat_err;
  if (!LoadCategories(&categories, &cat_err)) {
    ShowInfoBox(s->hwnd, cat_err.c_str(), L"重新加载分类失败");
    return false;
  }

  std::vector<Task> tasks;
  std::wstring task_err;
  if (!LoadTasks(&tasks, &task_err)) {
    ShowInfoBox(s->hwnd, task_err.c_str(), L"重新加载任务失败");
    return false;
  }

  NormalizeCategories(&categories);
  s->categories = std::move(categories);
  s->tasks = std::move(tasks);
  RefreshCategoryCombo(s);
  return LoadSelectedDay(s, s->selected);
}

static void ExportFullData(AppState* s) {
  std::wstring pick;
  if (!PickFolderDialog(s->hwnd, L"选择导出到的文件夹", &pick)) return;

  if (IsUncPath(pick)) {
    int r = MessageBoxW(s->hwnd,
                        L"你选择的是网络共享路径(UNC)。\n\n"
                        L"导出会通过网络写入数据，可能存在泄露风险。\n\n"
                        L"仍要继续导出吗？",
                        kAppTitle, MB_YESNO | MB_ICONWARNING);
    if (r != IDYES) return;
  }

  std::wstring export_dir;
  std::wstring err;
  if (!ExportWorkLogLiteData(GetDataRootDir(), pick, &export_dir, &err)) {
    ShowInfoBox(s->hwnd, err.c_str(), L"导出失败");
    return;
  }

  std::wstring done = L"已导出到：\n" + export_dir + L"\n\n说明：这是完整数据目录的原样拷贝（包含日数据、分类、长期任务和兼容文件等）。";
  ShowInfoBox(s->hwnd, done.c_str(), L"导出完成");
}

static void ImportFullData(AppState* s) {
  if (!PromptSaveIfDirty(s)) return;

  std::wstring pick;
  if (!PickFolderDialog(s->hwnd, L"选择要导入的数据文件夹", &pick)) return;

  std::wstring src_root;
  if (!ResolveWorkLogLiteImportRoot(pick, &src_root)) {
    ShowInfoBox(s->hwnd, L"所选文件夹看起来不像 WorkLogLite 的数据目录。\n\n请选中：包含 categories.txt / tasks.wlt / YYYY\\MM\\YYYY-MM-DD.wlr 的目录。\n\n也支持选择包含 data 子目录的外层文件夹。", L"导入失败");
    return;
  }

  if (IsUncPath(src_root)) {
    int rr = MessageBoxW(s->hwnd,
                         L"你选择的是网络共享路径(UNC)。\n\n"
                         L"导入会通过网络读取数据。\n\n"
                         L"仍要继续导入吗？",
                         kAppTitle, MB_YESNO | MB_ICONWARNING);
    if (rr != IDYES) return;
  }

  std::wstring dst_root = GetDataRootDir();
  if (NormalizeDirForCompare(src_root) == NormalizeDirForCompare(dst_root)) {
    ShowInfoBox(s->hwnd, L"你选择的正是当前数据目录。\n\n如果你的目标是备份/迁移，请使用“导出全部数据”。", L"导入已取消");
    return;
  }

  int r = MessageBoxW(s->hwnd,
                      L"即将用所选目录的数据“替换”当前数据目录。\n\n"
                      L"- 会覆盖你当前的所有数据文件\n"
                      L"- 会覆盖同名数据文件，包括旧版本留下的 auth.wla\n"
                      L"- 程序会自动创建一个备份目录\n\n"
                      L"确定继续导入吗？",
                      kAppTitle, MB_YESNO | MB_ICONWARNING);
  if (r != IDYES) return;

  std::wstring backup;
  std::wstring err;
  if (!ImportWorkLogLiteData(pick, dst_root, &backup, &err)) {
    ShowInfoBox(s->hwnd, err.c_str(), L"导入失败");
    return;
  }

  if (!ReloadAppDataAfterFilesystemChange(s)) {
    ShowInfoBox(s->hwnd, L"数据已经导入，但界面重新加载失败。请关闭后重新打开程序。", L"导入后重载失败");
    return;
  }

  std::wstring done = L"导入完成。\n\n当前数据目录已替换。";
  if (!backup.empty()) {
    done += L"\n备份目录：\n" + backup;
  }
  done +=
                      L"\n\n提示：如果导入的数据包含旧版本的 auth.wla，它会作为兼容文件一并保留，但当前版本不会启用密码验证。";
  ShowInfoBox(s->hwnd, done.c_str(), L"导入完成");
}

static void ClearAllData(AppState* s) {
  if (!PromptSaveIfDirty(s)) return;

  std::wstring root = GetDataRootDir();
  std::wstring backup;

  std::wstring msg1 =
      L"即将清空本机所有 WorkLogLite 数据(重置)。\n\n"
      L"会清空的数据包括：\n"
      L"- 所有日记录(.wlr)\n"
      L"- 分类(categories.txt)\n"
      L"- 长期任务(tasks.wlt)\n"
      L"- 旧版本兼容文件(auth.wla)\n\n"
      L"当前数据目录：\n" + root + L"\n\n"
      L"程序会先备份整个目录，再重建一个干净的数据目录。\n\n"
      L"确定继续？";
  int r1 = MessageBoxW(s->hwnd, msg1.c_str(), kAppTitle, MB_YESNO | MB_ICONWARNING);
  if (r1 != IDYES) return;

  int r2 = MessageBoxW(s->hwnd, L"最后确认：真的要清空全部数据吗？\n\n提示：你可以先用“导出全部数据”做一份拷贝。",
                       kAppTitle, MB_YESNO | MB_ICONWARNING);
  if (r2 != IDYES) return;

  std::wstring err;
  if (!ClearWorkLogLiteData(root, &backup, &err)) {
    ShowInfoBox(s->hwnd, err.c_str(), L"清空失败");
    return;
  }

  if (!ReloadAppDataAfterFilesystemChange(s)) {
    ShowInfoBox(s->hwnd, L"数据已经清空，但界面重新加载失败。请关闭后重新打开程序。", L"清空后重载失败");
    return;
  }

  std::wstring done =
      L"已清空数据并完成重置。\n\n"
      L"备份目录：\n" + backup + L"\n\n"
      L"提示：清空后会重新建立一个干净的数据目录。";
  ShowInfoBox(s->hwnd, done.c_str(), L"清空完成");
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

static const Task* FindTaskByIdInState(AppState* s, const std::wstring& id) {
  if (!s) return nullptr;
  for (const auto& t : s->tasks) {
    if (t.id == id) return &t;
  }
  return nullptr;
}

static std::wstring SanitizeFileName(const std::wstring& s) {
  std::wstring out = s;
  for (auto& ch : out) {
    if (ch == L'<' || ch == L'>' || ch == L':' || ch == L'"' || ch == L'/' || ch == L'\\' || ch == L'|' ||
        ch == L'?' || ch == L'*') {
      ch = L'_';
    }
  }
  return out;
}

static ReportRange DefaultRangeForTaskProgress(AppState* s, const std::wstring& task_id) {
  ReportRange r{};
  const Task* t = FindTaskByIdInState(s, task_id);
  if (t && t->start.wYear != 0 && t->end.wYear != 0) {
    r.start = t->start;
    r.end = t->end;
    return r;
  }
  SYSTEMTIME now{};
  GetLocalTime(&now);
  r.end = now;
  r.start = AddDays(now, -30);
  return r;
}

static void ShowReportWindow(HWND owner, const std::wstring& title, const std::wstring& md, const std::wstring& default_name);

static void ViewTaskProgressSummary(AppState* s, const std::wstring& task_id) {
  if (!s) return;
  if (task_id.empty()) return;

  ReportRange r = DefaultRangeForTaskProgress(s, task_id);
  if (DaysBetween(r.start, r.end) < 0) {
    SYSTEMTIME tmp = r.start;
    r.start = r.end;
    r.end = tmp;
  }

  int days = DaysBetween(r.start, r.end) + 1;
  if (days > 400) {
    std::wstring msg = L"该任务的汇总范围较大：\n" + FormatDateYYYYMMDD(r.start) + L" ~ " + FormatDateYYYYMMDD(r.end) +
                       L"\n共 " + std::to_wstring(days) + L" 天。\n\n生成汇总可能较慢，是否继续？";
    int rr = MessageBoxW(s->hwnd, msg.c_str(), kAppTitle, MB_YESNO | MB_ICONWARNING);
    if (rr != IDYES) return;
  }

  bool include_empty_days = (days <= 60);
  std::wstring md, err;
  if (!GenerateTaskProgressMarkdown(task_id, r, include_empty_days, &md, &err)) {
    ShowInfoBox(s->hwnd, err.c_str(), L"生成失败");
    return;
  }

  const Task* t = FindTaskByIdInState(s, task_id);
  std::wstring title = L"任务进展汇总";
  if (t && !t->title.empty()) title += L": " + t->title;

  std::wstring file = L"task-progress-" + SanitizeFileName(task_id) + L"-" + FormatDateYYYYMMDD(r.start) + L"-" + FormatDateYYYYMMDD(r.end) + L".md";
  ShowReportWindow(s->hwnd, title, md, file);
}

struct PickTaskWindowState {
  HWND hwnd{};
  HWND list{};
  HWND btn_ok{};
  HWND btn_cancel{};
  HFONT font{};

  AppState* app{};
  std::vector<std::wstring> ids;
  bool ok{false};
  std::wstring out_id;
};

static void PickTaskRefreshList(PickTaskWindowState* s) {
  if (!s || !s->list) return;
  SendMessageW(s->list, LB_RESETCONTENT, 0, 0);
  s->ids.clear();
  if (!s->app) return;
  for (const auto& t : s->app->tasks) {
    if (t.id.empty() || t.title.empty()) continue;
    std::wstring line = t.title;
    if (!t.category.empty()) line = L"[" + t.category + L"] " + line;
    SendMessageW(s->list, LB_ADDSTRING, 0, (LPARAM)line.c_str());
    s->ids.push_back(t.id);
  }
  if (!s->ids.empty()) SendMessageW(s->list, LB_SETCURSEL, 0, 0);
}

static int PickTaskGetSelectedIndex(PickTaskWindowState* s) {
  if (!s || !s->list) return -1;
  return (int)SendMessageW(s->list, LB_GETCURSEL, 0, 0);
}

static LRESULT CALLBACK PickTaskWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
  auto* s = reinterpret_cast<PickTaskWindowState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
    case WM_CREATE: {
      auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
      s = reinterpret_cast<PickTaskWindowState*>(cs->lpCreateParams);
      s->hwnd = hwnd;
      SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(s));
      s->font = CreateUiFont(hwnd);

      s->list = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"",
                                WS_CHILD | WS_VISIBLE | LBS_NOTIFY | WS_VSCROLL | WS_TABSTOP,
                                0, 0, 10, 10, hwnd, (HMENU)1, nullptr, nullptr);
      s->btn_ok = CreateWindowExW(0, L"BUTTON", L"查看汇总",
                                  WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                  0, 0, 10, 10, hwnd, (HMENU)2, nullptr, nullptr);
      s->btn_cancel = CreateWindowExW(0, L"BUTTON", L"取消",
                                      WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                      0, 0, 10, 10, hwnd, (HMENU)3, nullptr, nullptr);

      SetControlFont(s->list, s->font);
      SetControlFont(s->btn_ok, s->font);
      SetControlFont(s->btn_cancel, s->font);

      PickTaskRefreshList(s);
      return 0;
    }
    case WM_SIZE: {
      int w = LOWORD(lParam);
      int h = HIWORD(lParam);
      int pad = ScalePx(hwnd, 10);
      int btn_h = ScalePx(hwnd, 30);
      int btn_w = ScalePx(hwnd, 110);
      int y_btn = h - pad - btn_h;
      MoveWindow(s->list, pad, pad, w - 2 * pad, y_btn - 2 * pad, TRUE);
      int x = w - pad - btn_w;
      MoveWindow(s->btn_cancel, x, y_btn, btn_w, btn_h, TRUE);
      x -= pad + btn_w;
      MoveWindow(s->btn_ok, x, y_btn, btn_w, btn_h, TRUE);
      return 0;
    }
    case WM_COMMAND: {
      int id = LOWORD(wParam);
      int code = HIWORD(wParam);
      if (id == 2) {
        int idx = PickTaskGetSelectedIndex(s);
        if (idx >= 0 && idx < (int)s->ids.size()) {
          s->ok = true;
          s->out_id = s->ids[(size_t)idx];
          DestroyWindow(hwnd);
        }
        return 0;
      }
      if (id == 3) {
        DestroyWindow(hwnd);
        return 0;
      }
      if (id == 1 && code == LBN_DBLCLK) {
        SendMessageW(hwnd, WM_COMMAND, 2, 0);
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

static bool PromptPickTaskId(AppState* app, std::wstring* out_task_id) {
  if (out_task_id) out_task_id->clear();
  if (!app) return false;
  if (app->tasks.empty()) {
    ShowInfoBox(app->hwnd, L"当前没有长期任务。请先在“工具 -> 管理长期任务”中创建任务。", L"提示");
    return false;
  }

  WNDCLASSW wc{};
  wc.lpfnWndProc = PickTaskWndProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.lpszClassName = L"WorkLogLitePickTaskWnd";
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
  RegisterClassW(&wc);

  PickTaskWindowState state{};
  state.app = app;

  HWND hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME, wc.lpszClassName, L"选择任务",
                              WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                              CW_USEDEFAULT, CW_USEDEFAULT, 720, 520,
                              app->hwnd, nullptr, wc.hInstance, &state);
  if (!hwnd) {
    ShowLastErrorBox(app->hwnd, L"CreateWindowEx(PickTask)");
    return false;
  }

  EnableWindow(app->hwnd, FALSE);
  while (IsWindow(hwnd)) {
    MSG msg{};
    while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
      if (msg.message == WM_QUIT) {
        PostQuitMessage((int)msg.wParam);
        EnableWindow(app->hwnd, TRUE);
        return false;
      }
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
    WaitMessage();
  }
  EnableWindow(app->hwnd, TRUE);
  SetForegroundWindow(app->hwnd);

  if (state.ok && out_task_id) *out_task_id = state.out_id;
  return state.ok;
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
  state.markdown = NormalizeNewlinesToCrlf(md);
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
  TryLoadRichEditMsftedit();

  WNDCLASSW wc{};
  wc.lpfnWndProc = PreviewWndProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.lpszClassName = L"WorkLogLitePreviewWnd";
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
  RegisterClassW(&wc);

  PreviewWindowState state{};
  state.markdown = NormalizeNewlinesToCrlf(markdown);

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
  (void)OpenPathInShell(s->hwnd, path, L"导出文件");
}

static void OpenDataDir(AppState* s) {
  std::wstring dir = GetDataRootDir();
  EnsureDirExists(dir);
  (void)OpenPathInShell(s ? s->hwnd : nullptr, dir, L"数据目录");
}

static void ShowHelp(AppState* s) {
  const wchar_t* text =
      L"需要填写什么？\n"
      L"- 分类: 必填，可自定义(例如 工作/会议/目标/沟通/研发)。\n"
      L"- 时间: 可选，格式 HH:MM。\n"
      L"- 标题: 建议填写，用于列表和汇总更好看。\n"
      L"- 内容: 富文本(支持段落/对齐/缩进/编号/项目符号)。\n"
      L"  在正文框内: Ctrl+B 加粗, Ctrl+I 斜体, Ctrl+U 下划线。\n"
      L"  可直接输入 Markdown 后点 MD，或按 Ctrl+Shift+M 应用样式。\n"
      L"  列表: Ctrl+Shift+7 编号, Ctrl+Shift+8 项目符号。\n"
      L"  段落: Ctrl+Alt+L 左对齐, Ctrl+Alt+E 居中, Ctrl+Alt+R 右对齐。\n"
      L"\n"
      L"如何操作？\n"
      L"- 日历点一天: 切换到该日期记录。\n"
      L"- 列表选中一条: 右侧进入编辑。\n"
      L"- 查看 -> 预览: 查看 Markdown 渲染效果。\n"
      L"- 报告 -> 导出 CSV: 直接导出本季度/本年度 CSV，方便 Excel/脚本处理。\n"
      L"- 演示数据: 工具 -> 生成示例数据(演示)。\n"
      L"- 数据兼容: 旧版本的 auth.wla 会在导入导出时一并处理，但当前版本不再启用密码登录。\n"
      L"\n"
      L"快捷键:\n"
      L"- Ctrl+S 保存\n"
      L"- Ctrl+N 新增\n"
      L"- Delete 删除(会二次确认)\n"
      L"- Ctrl+P 预览当前条目\n"
      L"- Ctrl+Shift+P 预览当天全部\n"
      L"- Ctrl+K 管理分类\n"
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
      int action_gap = ScalePx(hwnd, 8);
      int footer_gap = ScalePx(hwnd, 10);

      bool split_actions = w < ScalePx(hwnd, 720);
      int action_rows = split_actions ? 2 : 1;
      int actions_h = action_rows * btn_h + (action_rows - 1) * action_gap;
      int footer_h = btn_h;
      int list_h = h - (pad * 4 + row + actions_h + footer_gap + footer_h);
      if (list_h < ScalePx(hwnd, 160)) list_h = ScalePx(hwnd, 160);

      int y = pad;
      MoveWindow(s->list, pad, y, w - 2 * pad, list_h, TRUE);
      y += list_h + pad;

      MoveWindow(s->edit, pad, y, w - 2 * pad, row, TRUE);
      y += row + action_gap;

      if (!split_actions) {
        int total_actions_w = btn_w * 3 + action_gap * 2;
        int x = w - pad - total_actions_w;
        if (x < pad) x = pad;
        MoveWindow(s->btn_add, x, y, btn_w, btn_h, TRUE);
        x += btn_w + action_gap;
        MoveWindow(s->btn_rename, x, y, btn_w, btn_h, TRUE);
        x += btn_w + action_gap;
        MoveWindow(s->btn_del, x, y, btn_w, btn_h, TRUE);
      } else {
        int half_w = (w - pad * 2 - action_gap) / 2;
        int del_w = w - pad * 2;
        if (half_w < ScalePx(hwnd, 120)) half_w = ScalePx(hwnd, 120);
        MoveWindow(s->btn_add, pad, y, half_w, btn_h, TRUE);
        MoveWindow(s->btn_rename, pad + half_w + action_gap, y, half_w, btn_h, TRUE);
        y += btn_h + action_gap;
        MoveWindow(s->btn_del, pad, y, del_w, btn_h, TRUE);
      }

      int y2 = h - pad - btn_h;
      int save_w = ScalePx(hwnd, 130);
      int cancel_w = ScalePx(hwnd, 110);
      int footer_total_w = save_w + cancel_w + action_gap;
      int x2 = w - pad - footer_total_w;
      if (x2 < pad) x2 = pad;
      MoveWindow(s->btn_save, x2, y2, save_w, btn_h, TRUE);
      MoveWindow(s->btn_cancel, x2 + save_w + action_gap, y2, cancel_w, btn_h, TRUE);
      return 0;
    }
    case WM_GETMINMAXINFO: {
      auto* mmi = reinterpret_cast<MINMAXINFO*>(lParam);
      if (!mmi) return 0;
      mmi->ptMinTrackSize.x = ScalePx(hwnd, 560);
      mmi->ptMinTrackSize.y = ScalePx(hwnd, 420);
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
        SetWindowTextW(s->edit, L"");
        SetFocus(s->edit);
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
                              CW_USEDEFAULT, CW_USEDEFAULT, 760, 560,
                              app->hwnd, nullptr, wc.hInstance, &state);
  if (!hwnd) {
    ShowLastErrorBox(app->hwnd, L"CreateWindowEx(Categories)");
    return;
  }

  CenterOwnedWindowSimple(hwnd, app->hwnd);

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
  HWND lbl_materials{};
  HWND ed_materials{};
  HWND btn_pick_dir{};
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
  HWND btn_open_dir{};
  HWND btn_close{};
  HFONT font{};

  AppState* app{};
  std::vector<Task> working;
  int initial_select{-1};

  // When creating a new task (no list selection), we generate an id early so the user can
  // set/open a materials folder before the first save.
  std::wstring draft_task_id;
};

// Helpers used by the long-term task window (defined later in this file).
static std::wstring NewTaskId();
static std::wstring GetTaskMaterialsDirById(const std::wstring& task_id);
static std::wstring TaskMaterialsDirResolved(const Task& t);

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

static void TaskInitListColumns(HWND list) {
  ListView_SetExtendedListViewStyle(list, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_DOUBLEBUFFER);

  LVCOLUMNW col{};
  col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;

  col.pszText = (LPWSTR)L"任务";
  col.cx = 220;
  col.iSubItem = 0;
  ListView_InsertColumn(list, 0, &col);

  col.pszText = (LPWSTR)L"材料";
  col.cx = 260;
  col.iSubItem = 1;
  ListView_InsertColumn(list, 1, &col);
}

static void TaskRefreshList(TaskWindowState* s) {
  ListView_DeleteAllItems(s->list);
  for (size_t i = 0; i < s->working.size(); ++i) {
    const Task& t = s->working[i];
    std::wstring item = FormatTaskListItem(t);

    LVITEMW lvi{};
    lvi.mask = LVIF_TEXT;
    lvi.iItem = (int)i;
    lvi.iSubItem = 0;
    lvi.pszText = const_cast<LPWSTR>(item.c_str());
    int row = ListView_InsertItem(s->list, &lvi);

    std::wstring materials = TaskMaterialsDirResolved(t);
    ListView_SetItemText(s->list, row, 1, const_cast<LPWSTR>(materials.c_str()));
  }
}

static int TaskGetSelectedIndex(TaskWindowState* s) {
  return ListView_GetNextItem(s->list, -1, LVNI_SELECTED);
}

static void TaskSetSelectedIndex(TaskWindowState* s, int idx) {
  ListView_SetItemState(s->list, -1, 0, LVIS_SELECTED | LVIS_FOCUSED);
  if (idx >= 0 && idx < (int)s->working.size()) {
    ListView_SetItemState(s->list, idx, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
    ListView_EnsureVisible(s->list, idx, FALSE);
  }
}

static void TaskFillFieldsFromSelection(TaskWindowState* s) {
  int idx = TaskGetSelectedIndex(s);
  if (idx < 0 || idx >= (int)s->working.size()) return;
  s->draft_task_id.clear();
  const Task& t = s->working[(size_t)idx];
  SetWindowTextW(s->cb_category, t.category.c_str());
  SetWindowTextW(s->ed_title, t.title.c_str());
  SetWindowTextW(s->ed_basis, t.basis.c_str());
  // Display resolved materials path (fallback to default auto dir if empty).
  if (s->ed_materials) {
    std::wstring dir = TaskMaterialsDirResolved(t);
    SetWindowTextW(s->ed_materials, dir.c_str());
  }
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
  // New task: generate a draft id so materials folder can be opened/selected immediately.
  s->draft_task_id = NewTaskId();
  if (s->ed_materials) {
    std::wstring dir = GetTaskMaterialsDirById(s->draft_task_id);
    SetWindowTextW(s->ed_materials, dir.c_str());
  }
  SetEditText(s->re_desc, L"");
  SendMessageW(s->cb_status, CB_SETCURSEL, 1, 0);
  SYSTEMTIME now{};
  GetLocalTime(&now);
  SYSTEMTIME next = AddDays(now, 1);
  DateTime_SetSystemtime(s->dt_start, GDT_VALID, &now);
  DateTime_SetSystemtime(s->dt_end, GDT_VALID, &next);
  TaskSetSelectedIndex(s, -1);
}

static bool TaskReadFields(TaskWindowState* s, Task* out, std::wstring* err) {
  if (!out) return false;
  out->category = TrimWs(GetWindowTextWString(s->cb_category));
  out->title = TrimWs(GetWindowTextWString(s->ed_title));
  out->basis = TrimWs(GetWindowTextWString(s->ed_basis));
  out->materials_dir = TrimWs(GetWindowTextWString(s->ed_materials));
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

static int CompareSystemTimeDateOnly(const SYSTEMTIME& a, const SYSTEMTIME& b) {
  if (a.wYear != b.wYear) return (a.wYear < b.wYear) ? -1 : 1;
  if (a.wMonth != b.wMonth) return (a.wMonth < b.wMonth) ? -1 : 1;
  if (a.wDay != b.wDay) return (a.wDay < b.wDay) ? -1 : 1;
  return 0;
}

static bool TaskReadFieldsStrict(TaskWindowState* s, Task* out, std::wstring* err) {
  if (!out) return false;
  out->category = TrimWs(GetWindowTextWString(s->cb_category));
  out->title = TrimWs(GetWindowTextWString(s->ed_title));
  out->basis = TrimWs(GetWindowTextWString(s->ed_basis));
  out->materials_dir = TrimWs(GetWindowTextWString(s->ed_materials));
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
  if (CompareSystemTimeDateOnly(ed, st) <= 0) {
    if (err) *err = L"结束日期必须晚于开始日期。";
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

static void TaskOpenMaterialsDir(TaskWindowState* s, bool allow_draft_dir) {
  int idx = TaskGetSelectedIndex(s);
  if (idx < 0 || idx >= (int)s->working.size()) {
    if (allow_draft_dir) {
      std::wstring dir = TrimWs(GetWindowTextWString(s->ed_materials));
      if (!dir.empty()) {
        if (IsNetworkPathForbidden(dir)) {
          ShowInfoBox(s->hwnd, L"材料路径为网络共享(UNC)，出于安全考虑已禁止打开。请改为本地目录。", L"提示");
          return;
        }
        if (!DirExistsW(dir)) EnsureDirExists(dir);
        (void)OpenPathInShell(s->hwnd, dir, L"材料目录");
        return;
      }
    }
    ShowInfoBox(s->hwnd, L"请先在左侧列表选择一个任务，或先点击“新建”。", L"提示");
    return;
  }
  const Task& t = s->working[(size_t)idx];
  std::wstring dir = TaskMaterialsDirResolved(t);
  if (IsNetworkPathForbidden(dir)) {
    ShowInfoBox(s->hwnd, L"材料路径为网络共享(UNC)，出于安全考虑已禁止打开。请改为本地目录。", L"提示");
    return;
  }
  if (!DirExistsW(dir)) {
    ShowInfoBox(s->hwnd, L"材料路径不存在或不是文件夹，请重新设置材料路径。", L"提示");
    return;
  }
  (void)OpenPathInShell(s->hwnd, dir, L"材料目录");
}

static std::wstring NewTaskId() {
  SYSTEMTIME st{};
  GetLocalTime(&st);
  wchar_t buf[64]{};
  wsprintfW(buf, L"t%04d%02d%02d-%02d%02d%02d",
            (int)st.wYear, (int)st.wMonth, (int)st.wDay, (int)st.wHour, (int)st.wMinute, (int)st.wSecond);
  return buf;
}

static std::wstring GetTaskMaterialsRootDir() {
  std::wstring root = JoinPath(GetDataRootDir(), L"task_materials");
  EnsureDirExists(root);
  return root;
}

static std::wstring SanitizeDirNameComponent(std::wstring s) {
  // Windows forbids: < > : " / \ | ? * and control chars.
  // Also trailing dots/spaces are problematic for folders.
  for (auto& ch : s) {
    if (ch < 32) { ch = L'_'; continue; }
    switch (ch) {
      case L'<':
      case L'>':
      case L':':
      case L'"':
      case L'/':
      case L'\\':
      case L'|':
      case L'?':
      case L'*': ch = L'_'; break;
      default: break;
    }
  }
  while (!s.empty() && (s.back() == L' ' || s.back() == L'.')) s.pop_back();
  while (!s.empty() && (s.front() == L' ')) s.erase(s.begin());
  if (s.empty()) s = L"task";
  if (s.size() > 80) s.resize(80);
  return s;
}

static std::wstring GetTaskMaterialsDirById(const std::wstring& task_id) {
  std::wstring safe = SanitizeDirNameComponent(task_id);
  std::wstring dir = JoinPath(GetTaskMaterialsRootDir(), safe);
  EnsureDirExists(dir);
  return dir;
}

static bool StartsWithW(const std::wstring& s, const std::wstring& prefix) {
  return s.size() >= prefix.size() && s.compare(0, prefix.size(), prefix) == 0;
}

static bool IsNetworkPathForbidden(const std::wstring& path) {
  // Block UNC paths (e.g. \\server\share) to avoid implicit network access.
  // Allow extended-length local paths like \\?\C:\...
  if (StartsWithW(path, L"\\\\?\\UNC\\") || StartsWithW(path, L"//?/UNC/")) return true;
  if (path.size() >= 2 && ((path[0] == L'\\' && path[1] == L'\\') || (path[0] == L'/' && path[1] == L'/'))) {
    if (StartsWithW(path, L"\\\\?\\") || StartsWithW(path, L"\\\\.\\")) return false;
    return true;
  }
  return false;
}

static std::wstring TaskMaterialsDirResolved(const Task& t) {
  if (!t.materials_dir.empty()) return t.materials_dir;
  return GetTaskMaterialsDirById(t.id);
}

static bool ConfirmMaterialsPathIfExternal(HWND owner, const std::wstring& picked_dir) {
  std::wstring data_root = GetDataRootDir();
  if (picked_dir.empty()) return true;
  if (IsNetworkPathForbidden(picked_dir)) {
    ShowInfoBox(owner, L"不允许设置为网络共享路径(UNC)。请改为本地磁盘目录。", L"提示");
    return false;
  }
  if (!DirExistsW(picked_dir)) {
    ShowInfoBox(owner, L"所选路径不存在或不是文件夹。", L"提示");
    return false;
  }
  if (!IsSameOrChildDir(data_root, picked_dir)) {
    int r = MessageBoxW(owner,
                        L"该材料路径不在 WorkLogLite 数据目录内。\n\n"
                        L"说明：导出/导入“全量数据”只会拷贝 data 目录，无法携带 data 目录之外的外部文件。\n\n"
                        L"仍要继续使用该路径吗？",
                        kAppTitle, MB_YESNO | MB_ICONQUESTION);
    if (r != IDYES) return false;
  }
  return true;
}

static void SetMaterialsDirForSelected(AppState* s) {
  if (!s) return;

  int sel = ListView_GetNextItem(s->list, -1, LVNI_SELECTED);
  if (sel < 0 || sel >= (int)s->day.entries.size()) {
    // No selection: apply to current editor (useful for new entries before saving).
    std::wstring picked;
    if (!PickFolderDialog(s->hwnd, L"选择材料目录(本地)", &picked)) return;
    if (!ConfirmMaterialsPathIfExternal(s->hwnd, picked)) return;
    s->editor_materials_dir = picked;
    SetEditorDirty(s, true);
    UpdateStatusBar(s);
    ShowToast(s, L"已设置当前编辑内容的材料路径(保存后生效)");
    return;
  }

  Entry& e = s->day.entries[(size_t)sel];
  if (EntryIsTaskProgress(e)) {
    Task* t = FindTaskById(s, e.task_id);
    if (!t) {
      ShowInfoBox(s->hwnd, L"未找到对应的长期任务，无法设置材料路径。", L"提示");
      return;
    }
    std::wstring picked;
    if (!PickFolderDialog(s->hwnd, L"选择长期任务材料目录(本地)", &picked)) return;
    if (!ConfirmMaterialsPathIfExternal(s->hwnd, picked)) return;
    std::wstring old_dir = t->materials_dir;
    t->materials_dir = picked;
    if (s->editing_index == sel) s->editor_materials_dir = picked;
    std::wstring terr;
    if (!SaveTasks(s->tasks, &terr)) {
      t->materials_dir = old_dir;
      if (s->editing_index == sel) s->editor_materials_dir = old_dir;
      ShowInfoBox(s->hwnd, terr.c_str(), L"保存任务失败");
      return;
    }
    RefreshList(s);
    UpdateStatusBar(s);
    ShowToast(s, L"已设置该长期任务的材料路径");
    return;
  }

  std::wstring picked;
  if (!PickFolderDialog(s->hwnd, L"选择材料目录(本地)", &picked)) return;
  if (!ConfirmMaterialsPathIfExternal(s->hwnd, picked)) return;

  e.materials_dir = picked;
  if (s->editing_index == sel) s->editor_materials_dir = picked;
  SetEditorDirty(s, true);
  RefreshList(s);
  UpdateStatusBar(s);
  ShowToast(s, L"已设置该条目的材料路径(保存后写入数据)");
}

static void ClearMaterialsDirForSelected(AppState* s) {
  if (!s) return;
  int sel = ListView_GetNextItem(s->list, -1, LVNI_SELECTED);
  if (sel < 0 || sel >= (int)s->day.entries.size()) {
    s->editor_materials_dir.clear();
    SetEditorDirty(s, true);
    UpdateStatusBar(s);
    ShowToast(s, L"已清除当前编辑内容的材料路径");
    return;
  }
  Entry& e = s->day.entries[(size_t)sel];
  if (EntryIsTaskProgress(e)) {
    Task* t = FindTaskById(s, e.task_id);
    if (!t) return;
    std::wstring old_dir = t->materials_dir;
    t->materials_dir.clear();
    if (s->editing_index == sel) s->editor_materials_dir.clear();
    std::wstring terr;
    if (!SaveTasks(s->tasks, &terr)) {
      t->materials_dir = old_dir;
      if (s->editing_index == sel) s->editor_materials_dir = old_dir;
      ShowInfoBox(s->hwnd, terr.c_str(), L"保存任务失败");
      return;
    }
    RefreshList(s);
    UpdateStatusBar(s);
    ShowToast(s, L"已清除该长期任务的自定义材料路径");
    return;
  }
  e.materials_dir.clear();
  if (s->editing_index == sel) s->editor_materials_dir.clear();
  SetEditorDirty(s, true);
  RefreshList(s);
  UpdateStatusBar(s);
  ShowToast(s, L"已清除该条目的自定义材料路径(保存后写入数据)");
}

static void OpenTaskMaterialsDir(HWND owner, const std::wstring& task_id) {
  if (task_id.empty()) {
    ShowInfoBox(owner, L"任务ID为空，无法打开材料目录。", L"提示");
    return;
  }
  std::wstring dir = GetTaskMaterialsDirById(task_id);
  (void)OpenPathInShell(owner, dir, L"材料目录");
}

static std::wstring DateKeyYmd(const SYSTEMTIME& st) {
  wchar_t buf[32]{};
  wsprintfW(buf, L"%04d-%02d-%02d", (int)st.wYear, (int)st.wMonth, (int)st.wDay);
  return buf;
}

static std::wstring GetWorkMaterialsRootDir() {
  std::wstring root = JoinPath(GetDataRootDir(), L"work_materials");
  EnsureDirExists(root);
  return root;
}

static std::wstring GetWorkMaterialsDir(const SYSTEMTIME& date, const std::wstring& entry_id) {
  std::wstring root = GetWorkMaterialsRootDir();
  std::wstring day = JoinPathLoose(root, DateKeyYmd(date));
  EnsureDirExists(day);
  std::wstring safe_id = SanitizeDirNameComponent(entry_id.empty() ? L"no_id" : entry_id);
  std::wstring dir = JoinPathLoose(day, safe_id);
  EnsureDirExists(dir);
  return dir;
}

static void OpenWorkMaterialsDir(HWND owner, const SYSTEMTIME& date, const std::wstring& entry_id) {
  if (entry_id.empty()) {
    ShowInfoBox(owner, L"该条目尚未保存(没有ID)。请先点击“保存”，再打开材料目录。", L"提示");
    return;
  }
  std::wstring dir = GetWorkMaterialsDir(date, entry_id);
  (void)OpenPathInShell(owner, dir, L"材料目录");
}

static void OpenMaterialsDirForEntry(HWND owner, AppState* s, const Entry& e) {
  if (!s) return;
  if (EntryIsTaskProgress(e) && !e.task_id.empty()) {
    Task* t = FindTaskById(s, e.task_id);
    if (t) {
      std::wstring dir = TaskMaterialsDirResolved(*t);
      if (IsNetworkPathForbidden(dir)) {
        ShowInfoBox(owner, L"材料路径为网络共享(UNC)，出于安全考虑已禁止打开。请改为本地目录。", L"提示");
        return;
      }
      if (!DirExistsW(dir)) {
        ShowInfoBox(owner, L"材料路径不存在或不是文件夹。请重新设置材料路径。", L"提示");
        return;
      }
      (void)OpenPathInShell(owner, dir, L"材料目录");
      return;
    }
    OpenTaskMaterialsDir(owner, e.task_id);
    return;
  }
  if (!e.materials_dir.empty()) {
    if (IsNetworkPathForbidden(e.materials_dir)) {
      ShowInfoBox(owner, L"材料路径为网络共享(UNC)，出于安全考虑已禁止打开。请改为本地目录。", L"提示");
      return;
    }
    if (!DirExistsW(e.materials_dir)) {
      ShowInfoBox(owner, L"材料路径不存在或不是文件夹。请重新设置材料路径。", L"提示");
      return;
    }
    (void)OpenPathInShell(owner, e.materials_dir, L"材料目录");
    return;
  }
  OpenWorkMaterialsDir(owner, s->selected, e.id);
}

static void EnsureMaterialsDirForEntry(AppState* s, const Entry& e) {
  if (!s) return;
  if (EntryIsTaskProgress(e) && !e.task_id.empty()) {
    (void)GetTaskMaterialsDirById(e.task_id);
    return;
  }
  if (!e.materials_dir.empty()) return;
  if (!e.id.empty()) {
    (void)GetWorkMaterialsDir(s->selected, e.id);
  }
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
  t.status = EntryStatus::None;
  return t;
}

static void TaskPopulateCategoryCombo(TaskWindowState* s) {
  SendMessageW(s->cb_category, CB_RESETCONTENT, 0, 0);
  for (const auto& c : s->app->categories) {
    SendMessageW(s->cb_category, CB_ADDSTRING, 0, (LPARAM)c.c_str());
  }
  if (!s->app->categories.empty()) SetWindowTextW(s->cb_category, s->app->categories[0].c_str());
}

#if 0  // Recurring meetings removed (kept disabled for history; data file may still exist for old users).
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

      int needed = btn_w * 5 + pad * 4;
      int bw = btn_w;
      if (right_w < needed) {
        bw = (right_w - pad * 4) / 5;
        if (bw < ScalePx(hwnd, 70)) bw = ScalePx(hwnd, 70);
      }
      int bx = x;
      MoveWindow(s->btn_new, bx, y, bw, btn_h, TRUE);
      bx += bw + pad;
      MoveWindow(s->btn_save, bx, y, bw, btn_h, TRUE);
      bx += bw + pad;
      MoveWindow(s->btn_open_dir, bx, y, bw, btn_h, TRUE);
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

#endif  // 0

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
      s->lbl_materials = CreateWindowExW(0, L"STATIC", L"材料目录(本地，可导出时建议放在 data 内)",
                                         WS_CHILD | WS_VISIBLE,
                                         0, 0, 10, 10, hwnd, (HMENU)27, nullptr, nullptr);
      s->ed_materials = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                        WS_CHILD | WS_VISIBLE | ES_AUTOHSCROLL | ES_READONLY,
                                        0, 0, 10, 10, hwnd, (HMENU)17, nullptr, nullptr);
      s->btn_pick_dir = CreateWindowExW(0, L"BUTTON", L"设置路径",
                                        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                        0, 0, 10, 10, hwnd, (HMENU)12, nullptr, nullptr);
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
      SendMessageW(s->cb_status, CB_ADDSTRING, 0, (LPARAM)L"进行中");
      SendMessageW(s->cb_status, CB_ADDSTRING, 0, (LPARAM)L"阻塞");
      SendMessageW(s->cb_status, CB_ADDSTRING, 0, (LPARAM)L"已完成");
      SendMessageW(s->cb_status, CB_SETCURSEL, 0, 0);

      s->lbl_desc = CreateWindowExW(0, L"STATIC", L"任务说明(支持富文本，可做段落/编号/缩进)",
                                    WS_CHILD | WS_VISIBLE,
                                    0, 0, 10, 10, hwnd, (HMENU)26, nullptr, nullptr);

      s->btn_new = CreateWindowExW(0, L"BUTTON", L"新建",
                                   WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                   0, 0, 10, 10, hwnd, (HMENU)7, nullptr, nullptr);
      s->btn_save = CreateWindowExW(0, L"BUTTON", L"保存",
                                      WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                      0, 0, 10, 10, hwnd, (HMENU)8, nullptr, nullptr);
      s->btn_open_dir = CreateWindowExW(0, L"BUTTON", L"材料目录",
                                        WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                        0, 0, 10, 10, hwnd, (HMENU)11, nullptr, nullptr);
      s->btn_del = CreateWindowExW(0, L"BUTTON", L"删除",
                                   WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                   0, 0, 10, 10, hwnd, (HMENU)9, nullptr, nullptr);
      s->btn_close = CreateWindowExW(0, L"BUTTON", L"关闭",
                                      WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_TABSTOP,
                                      0, 0, 10, 10, hwnd, (HMENU)10, nullptr, nullptr);

      for (HWND h : {s->list, s->lbl_category, s->cb_category, s->lbl_title, s->ed_title, s->lbl_basis, s->ed_basis,
                     s->lbl_materials, s->ed_materials, s->btn_pick_dir,
                     s->lbl_start, s->dt_start, s->lbl_end, s->dt_end, s->lbl_status, s->cb_status,
                     s->lbl_desc, s->re_desc,
                     s->btn_new, s->btn_save, s->btn_open_dir, s->btn_del, s->btn_close}) {
        SetControlFont(h, s->font);
      }

      TaskPopulateCategoryCombo(s);
      TaskRefreshList(s);
      if (!s->working.empty()) {
        int sel = s->initial_select;
        if (sel < 0 || sel >= (int)s->working.size()) sel = 0;
        TaskSetSelectedIndex(s, sel);
        TaskFillFieldsFromSelection(s);
      } else {
        SYSTEMTIME now{};
        GetLocalTime(&now);
        SYSTEMTIME next = AddDays(now, 1);
        DateTime_SetSystemtime(s->dt_start, GDT_VALID, &now);
        DateTime_SetSystemtime(s->dt_end, GDT_VALID, &next);
      }

      SetEditCue(GetComboEditHandle(s->cb_category), L"任务分类，例如：项目/研发/学习");
      SetEditCue(s->ed_title, L"任务标题，例如：重构支付模块");
      SetEditCue(s->ed_basis, L"依据(可选)：会议/需求/故障单/邮件等");
      SetEditCue(s->ed_materials, L"(未设置时，保存会自动创建默认目录)");
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

      int left_w = ScalePx(hwnd, 380);
      MoveWindow(s->list, pad, pad, left_w, h - pad * 2, TRUE);
      int task_col_w = left_w * 45 / 100;
      int materials_col_w = left_w - task_col_w - ScalePx(hwnd, 24);
      if (materials_col_w < ScalePx(hwnd, 140)) materials_col_w = ScalePx(hwnd, 140);
      ListView_SetColumnWidth(s->list, 0, task_col_w);
      ListView_SetColumnWidth(s->list, 1, materials_col_w);

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

      MoveWindow(s->lbl_materials, x, y, right_w, lbl, TRUE);
      y += lbl + gap;
      // Materials row: readonly edit + set button + open button.
      int mbtn_w = ScalePx(hwnd, 92);
      int mbtn_gap = pad;
      int edit_w = right_w - (mbtn_w * 2 + mbtn_gap);
      if (edit_w < ScalePx(hwnd, 120)) edit_w = ScalePx(hwnd, 120);
      MoveWindow(s->ed_materials, x, y, edit_w, row, TRUE);
      MoveWindow(s->btn_pick_dir, x + edit_w + mbtn_gap, y, mbtn_w, row, TRUE);
      MoveWindow(s->btn_open_dir, x + edit_w + mbtn_gap + mbtn_w, y, mbtn_w, row, TRUE);
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
      if (src != s->re_desc && src != s->list) return DefWindowProcW(hwnd, msg, wParam, lParam);

      POINT pt{};
      if (lParam == (LPARAM)-1) {
        GetCursorPos(&pt);
      } else {
        pt.x = GET_X_LPARAM(lParam);
        pt.y = GET_Y_LPARAM(lParam);
      }

      if (src == s->list) {
        int idx = TaskGetSelectedIndex(s);
        bool has_sel = (idx >= 0 && idx < (int)s->working.size());

        HMENU m = CreatePopupMenu();
        if (!m) return 0;
        AppendMenuW(m, MF_STRING | (has_sel ? 0 : MF_GRAYED), 1001, L"打开材料目录");
        AppendMenuW(m, MF_STRING, 1002, L"打开任务材料根目录");

        UINT cmd = TrackPopupMenu(m, TPM_RIGHTBUTTON | TPM_RETURNCMD, pt.x, pt.y, 0, hwnd, nullptr);
        DestroyMenu(m);
        if (!cmd) return 0;

        if (cmd == 1001 && has_sel) {
          const Task& t = s->working[(size_t)idx];
          std::wstring dir = TaskMaterialsDirResolved(t);
          if (IsNetworkPathForbidden(dir)) {
            ShowInfoBox(hwnd, L"材料路径为网络共享(UNC)，出于安全考虑已禁止打开。请改为本地目录。", L"提示");
            return 0;
          }
          if (!DirExistsW(dir)) {
            ShowInfoBox(hwnd, L"材料路径不存在或不是文件夹。请重新设置材料路径。", L"提示");
            return 0;
          }
          (void)OpenPathInShell(hwnd, dir, L"材料目录");
        } else if (cmd == 1002) {
          std::wstring root = GetTaskMaterialsRootDir();
          (void)OpenPathInShell(hwnd, root, L"材料根目录");
        }
        return 0;
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

    case WM_NOTIFY: {
      LPNMHDR hdr = reinterpret_cast<LPNMHDR>(lParam);
      if (hdr && hdr->hwndFrom == s->list) {
        if (hdr->code == LVN_ITEMCHANGED) {
          auto* nmlv = reinterpret_cast<NMLISTVIEW*>(lParam);
          if ((nmlv->uChanged & LVIF_STATE) &&
              !(nmlv->uOldState & LVIS_SELECTED) &&
              (nmlv->uNewState & LVIS_SELECTED)) {
            TaskFillFieldsFromSelection(s);
          }
          return 0;
        }
        if (hdr->code == NM_DBLCLK) {
          auto* act = reinterpret_cast<LPNMITEMACTIVATE>(lParam);
          if (act->iItem >= 0) {
            TaskSetSelectedIndex(s, act->iItem);
            if (act->iSubItem == 1) {
              TaskOpenMaterialsDir(s, false);
            } else {
              TaskFillFieldsFromSelection(s);
            }
          }
          return 0;
        }
      }
      break;
    }
    case WM_COMMAND: {
      int id = LOWORD(wParam);
      if (id == 7) {  // new
        TaskClearFields(s);
        HWND h = GetComboEditHandle(s->cb_category);
        SetFocus(h ? h : s->cb_category);
        return 0;
      }
      if (id == 12) {  // pick materials dir
        std::wstring picked;
        if (!PickFolderDialog(hwnd, L"选择材料目录(本地)", &picked)) return 0;
        if (!ConfirmMaterialsPathIfExternal(hwnd, picked)) return 0;
        SetWindowTextW(s->ed_materials, picked.c_str());
        return 0;
      }
      if (id == 8) {  // save (add or update)
        Task t{};
        std::wstring e;
        if (!TaskReadFieldsStrict(s, &t, &e)) {
          ShowInfoBox(hwnd, e.c_str(), L"提示");
          return 0;
        }
        int idx = TaskGetSelectedIndex(s);
        if (idx >= 0 && idx < (int)s->working.size()) {
          t.id = s->working[(size_t)idx].id;
          // Preserve materials dir if user didn't change it (or old data had it empty).
          if (t.materials_dir.empty()) t.materials_dir = s->working[(size_t)idx].materials_dir;
          s->working[(size_t)idx] = t;
        } else {
          t.id = !s->draft_task_id.empty() ? s->draft_task_id : NewTaskId();
          // New task must have a materials directory. If user didn't set one, auto-create the default under data.
          if (t.materials_dir.empty()) t.materials_dir = GetTaskMaterialsDirById(t.id);
          s->working.push_back(t);
          idx = (int)s->working.size() - 1;
          s->draft_task_id.clear();
        }

        std::wstring err;
        if (!SaveTasks(s->working, &err)) {
          ShowInfoBox(hwnd, err.c_str(), L"保存失败");
          return 0;
        }

        // Ensure each long-term task has a stable materials folder.
        // We intentionally do NOT delete it when the task is deleted, to avoid data loss.
        if (t.materials_dir.empty()) (void)GetTaskMaterialsDirById(t.id);
        s->app->tasks = s->working;

        TaskRefreshList(s);
        TaskSetSelectedIndex(s, idx);
        TaskFillFieldsFromSelection(s);
        return 0;
      }
      if (id == 11) {  // open materials dir
        int idx = TaskGetSelectedIndex(s);
        if (idx < 0 || idx >= (int)s->working.size()) {
          // No selection: open draft task dir (new task) if present.
          std::wstring dir = TrimWs(GetWindowTextWString(s->ed_materials));
          if (!dir.empty()) {
            if (IsNetworkPathForbidden(dir)) {
              ShowInfoBox(hwnd, L"材料路径为网络共享(UNC)，出于安全考虑已禁止打开。请改为本地目录。", L"提示");
              return 0;
            }
            if (!DirExistsW(dir)) EnsureDirExists(dir);
            (void)OpenPathInShell(hwnd, dir, L"材料目录");
            return 0;
          }
          ShowInfoBox(hwnd, L"请先在左侧列表选择一个任务，或先点击“新建”。", L"提示");
          return 0;
        }
        const Task& t = s->working[(size_t)idx];
        std::wstring dir = TaskMaterialsDirResolved(t);
        if (IsNetworkPathForbidden(dir)) {
          ShowInfoBox(hwnd, L"材料路径为网络共享(UNC)，出于安全考虑已禁止打开。请改为本地目录。", L"提示");
          return 0;
        }
        if (!DirExistsW(dir)) {
          ShowInfoBox(hwnd, L"材料路径不存在或不是文件夹。请重新设置材料路径。", L"提示");
          return 0;
        }
        (void)OpenPathInShell(hwnd, dir, L"材料目录");
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
          TaskSetSelectedIndex(s, new_idx);
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
  int left_w = ScalePx(s->hwnd, 272);
  int top_h = ScalePx(s->hwnd, 248);
  int row_h = ScalePx(s->hwnd, 28);
  int lbl_h = ScalePx(s->hwnd, 20);
  int btn_h = ScalePx(s->hwnd, 36);
  int btn_w = ScalePx(s->hwnd, 96);
  int body_btn_w = ScalePx(s->hwnd, 106);

  auto set_vis = [](HWND hwnd, bool visible) {
    if (hwnd) ShowWindow(hwnd, visible ? SW_SHOW : SW_HIDE);
  };
  auto set_office_visible = [&](bool visible) {
    set_vis(s->grp_office, visible);
    for (HWND b : {s->btn_screenshot, s->btn_quick_reply, s->btn_timestamp, s->btn_focus_timer, s->btn_calc,
                   s->btn_data_dir, s->btn_screenshot_dir, s->btn_task_materials, s->btn_open_materials,
                   s->btn_set_materials, s->btn_color_picker, s->btn_paste_plain, s->btn_record_system,
                   s->btn_record_mic, s->btn_recordings_dir, s->btn_date_separator}) {
      set_vis(b, visible);
    }
  };

  bool maximized = s->editor_maximized;
  set_vis(s->cal, !maximized);
  set_vis(s->st_day, !maximized);
  set_vis(s->list, !maximized);
  set_office_visible(!maximized);
  set_vis(s->lbl_category, !maximized);
  set_vis(s->lbl_start, !maximized);
  set_vis(s->lbl_end, !maximized);
  set_vis(s->lbl_status, !maximized);
  set_vis(s->cb_category, !maximized);
  set_vis(s->ed_start, !maximized);
  set_vis(s->ed_end, !maximized);
  set_vis(s->cb_status, !maximized);
  set_vis(s->lbl_title, TRUE);
  set_vis(s->ed_title, TRUE);
  set_vis(s->lbl_body, TRUE);
  set_vis(s->btn_body_maximize, TRUE);
  set_vis(s->btn_new, TRUE);
  set_vis(s->btn_save, TRUE);
  set_vis(s->btn_preview, TRUE);
  set_vis(s->btn_del, TRUE);
  for (HWND b : {s->fmt_bold, s->fmt_italic, s->fmt_underline, s->fmt_superscript, s->fmt_subscript,
                 s->fmt_numbering, s->fmt_bullet, s->fmt_markdown,
                 s->fmt_align_left, s->fmt_align_center, s->fmt_align_right, s->fmt_indent_dec, s->fmt_indent_inc,
                 s->fmt_clear}) {
    set_vis(b, TRUE);
  }

  if (maximized) {
    int action_gap = ScalePx(s->hwnd, 10);
    int action_w = btn_w * 4 + action_gap * 3;
    int y = pad;
    int action_x = w - pad - action_w;
    if (action_x < pad) action_x = pad;
    int bx = action_x;
    MoveWindow(s->btn_new, bx, y, btn_w, btn_h, TRUE);
    bx += btn_w + action_gap;
    MoveWindow(s->btn_save, bx, y, btn_w, btn_h, TRUE);
    bx += btn_w + action_gap;
    MoveWindow(s->btn_preview, bx, y, btn_w, btn_h, TRUE);
    bx += btn_w + action_gap;
    MoveWindow(s->btn_del, bx, y, btn_w, btn_h, TRUE);

    y += btn_h + ScalePx(s->hwnd, 6);
    MoveWindow(s->lbl_title, pad, y, w - pad * 2, lbl_h, TRUE);
    y += lbl_h + ScalePx(s->hwnd, 2);
    MoveWindow(s->ed_title, pad, y, w - pad * 2, row_h, TRUE);
    y += row_h + pad;

    int body_label_w = (std::max)(ScalePx(s->hwnd, 160), w - pad * 3 - body_btn_w);
    MoveWindow(s->lbl_body, pad, y, body_label_w, lbl_h, TRUE);
    MoveWindow(s->btn_body_maximize, w - pad - body_btn_w, y - ScalePx(s->hwnd, 4), body_btn_w, btn_h, TRUE);
    y += btn_h + ScalePx(s->hwnd, 4);

    int tbh = ScalePx(s->hwnd, 26);
    int tbw = ScalePx(s->hwnd, 30);
    int tb_pad = ScalePx(s->hwnd, 6);
    int tx = pad;
    int ty = y;
    int wrap_right = w - pad;
    for (HWND b : {s->fmt_bold, s->fmt_italic, s->fmt_underline, s->fmt_superscript, s->fmt_subscript,
                   s->fmt_numbering, s->fmt_bullet, s->fmt_markdown,
                   s->fmt_align_left, s->fmt_align_center, s->fmt_align_right, s->fmt_indent_dec, s->fmt_indent_inc,
                   s->fmt_clear}) {
      if (!b) continue;
      if (tx + tbw > wrap_right) {
        tx = pad;
        ty += tbh + tb_pad;
      }
      MoveWindow(b, tx, ty, tbw, tbh, TRUE);
      tx += tbw + tb_pad;
    }
    y = ty + tbh + ScalePx(s->hwnd, 4);
    MoveWindow(s->ed_body, pad, y, w - pad * 2, h - y - pad, TRUE);
    return;
  }

  MoveWindow(s->cal, pad, pad, left_w, top_h, TRUE);

  // Left-bottom office panel.
  int office_y = pad + top_h + pad;
  int office_h = h - office_y - pad;
  if (office_h < ScalePx(s->hwnd, 120)) office_h = ScalePx(s->hwnd, 120);
  MoveWindow(s->grp_office, pad, office_y, left_w, office_h, TRUE);
  int inner = ScalePx(s->hwnd, 10);
  int gap = ScalePx(s->hwnd, 8);
  int oy0 = office_y + ScalePx(s->hwnd, 26);
  int ox0 = pad + inner;
  int cols = 2;
  int bh = ScalePx(s->hwnd, 34);
  int bw = (left_w - 2 * inner - gap) / cols;
  std::vector<HWND> btns = {
    s->btn_screenshot, s->btn_screenshot_dir,
    s->btn_record_system, s->btn_record_mic,
    s->btn_recordings_dir, s->btn_color_picker,
    s->btn_quick_reply, s->btn_timestamp,
    s->btn_paste_plain, s->btn_date_separator,
    s->btn_task_materials, s->btn_focus_timer,
    s->btn_calc, s->btn_open_materials,
    s->btn_set_materials, s->btn_data_dir
  };
  for (size_t i = 0; i < btns.size(); i++) {
    HWND b = btns[i];
    if (!b) continue;
    int row = (int)(i / cols);
    int col = (int)(i % cols);
    int x = ox0 + col * (bw + gap);
    int yb = oy0 + row * (bh + gap);
    MoveWindow(b, x, yb, bw, bh, TRUE);
  }

  int right_x = pad + left_w + pad;
  int right_w = w - right_x - pad;
  int y = pad;
  int action_gap = ScalePx(s->hwnd, 10);
  int action_w = btn_w * 4 + action_gap * 3;
  int pref_cat_w = ScalePx(s->hwnd, 160);
  int pref_date_w = ScalePx(s->hwnd, 130);
  int pref_status_w = ScalePx(s->hwnd, 110);
  int min_cat_w = ScalePx(s->hwnd, 120);
  int min_date_w = ScalePx(s->hwnd, 118);
  int min_status_w = ScalePx(s->hwnd, 90);
  int min_meta_w = min_cat_w + min_date_w + min_date_w + min_status_w + pad * 3;
  bool split_actions = right_w < (min_meta_w + pad + action_w);

  MoveWindow(s->st_day, right_x, y, right_w, lbl_h, TRUE);
  y += lbl_h + ScalePx(s->hwnd, 4);

  MoveWindow(s->list, right_x, y, right_w, top_h - lbl_h - ScalePx(s->hwnd, 4), TRUE);
  y += top_h + pad;

  int fields_w = split_actions ? right_w : (right_w - action_w - pad);
  if (fields_w < min_meta_w) fields_w = min_meta_w;
  int cat_w = pref_cat_w;
  int start_w = pref_date_w;
  int end_w = pref_date_w;
  int status_w = pref_status_w;
  int pref_meta_w = pref_cat_w + pref_date_w + pref_date_w + pref_status_w + pad * 3;
  if (fields_w < pref_meta_w) {
    int available = fields_w - pad * 3;
    int base_total = min_cat_w + min_date_w + min_date_w + min_status_w;
    int extra = available - base_total;
    if (extra < 0) extra = 0;
    int extra_total = (pref_cat_w - min_cat_w) + (pref_date_w - min_date_w) + (pref_date_w - min_date_w) + (pref_status_w - min_status_w);
    auto share_extra = [&](int pref, int minv) {
      if (extra_total <= 0) return minv;
      return minv + extra * (pref - minv) / extra_total;
    };
    cat_w = share_extra(pref_cat_w, min_cat_w);
    start_w = share_extra(pref_date_w, min_date_w);
    end_w = share_extra(pref_date_w, min_date_w);
    status_w = available - cat_w - start_w - end_w;
    if (status_w < min_status_w) status_w = min_status_w;
  }

  // Labels for row
  int x = right_x;
  MoveWindow(s->lbl_category, x, y, cat_w, lbl_h, TRUE);
  x += cat_w + pad;
  MoveWindow(s->lbl_start, x, y, start_w, lbl_h, TRUE);
  x += start_w + pad;
  MoveWindow(s->lbl_end, x, y, end_w, lbl_h, TRUE);
  x += end_w + pad;
  MoveWindow(s->lbl_status, x, y, status_w, lbl_h, TRUE);

  // Category / Start / End row
  y += lbl_h + ScalePx(s->hwnd, 2);
  x = right_x;
  // For combobox, the "height" controls dropdown height; keep it tall enough for suggestions.
  MoveWindow(s->cb_category, x, y, cat_w, row_h * 10, TRUE);
  x += cat_w + pad;
  MoveWindow(s->ed_start, x, y, start_w, row_h, TRUE);
  x += start_w + pad;
  MoveWindow(s->ed_end, x, y, end_w, row_h, TRUE);
  x += end_w + pad;
  MoveWindow(s->cb_status, x, y, status_w, row_h * 10, TRUE);

  int actions_y = y;
  if (!split_actions) {
    int bx = right_x + right_w - action_w;
    if (bx < right_x) bx = right_x;
    MoveWindow(s->btn_new, bx, actions_y, btn_w, btn_h, TRUE);
    bx += btn_w + action_gap;
    MoveWindow(s->btn_save, bx, actions_y, btn_w, btn_h, TRUE);
    bx += btn_w + action_gap;
    MoveWindow(s->btn_preview, bx, actions_y, btn_w, btn_h, TRUE);
    bx += btn_w + action_gap;
    MoveWindow(s->btn_del, bx, actions_y, btn_w, btn_h, TRUE);
    y += row_h + pad;
  } else {
    y += row_h + ScalePx(s->hwnd, 8);
    int bx = right_x + right_w - action_w;
    if (bx < right_x) bx = right_x;
    MoveWindow(s->btn_new, bx, y, btn_w, btn_h, TRUE);
    bx += btn_w + action_gap;
    MoveWindow(s->btn_save, bx, y, btn_w, btn_h, TRUE);
    bx += btn_w + action_gap;
    MoveWindow(s->btn_preview, bx, y, btn_w, btn_h, TRUE);
    bx += btn_w + action_gap;
    MoveWindow(s->btn_del, bx, y, btn_w, btn_h, TRUE);
    y += btn_h + pad;
  }

  MoveWindow(s->lbl_title, right_x, y, right_w, lbl_h, TRUE);
  y += lbl_h + ScalePx(s->hwnd, 2);
  MoveWindow(s->ed_title, right_x, y, right_w, row_h, TRUE);
  y += row_h + pad;

  int body_label_w = (std::max)(ScalePx(s->hwnd, 160), right_w - body_btn_w - pad);
  MoveWindow(s->lbl_body, right_x, y, body_label_w, lbl_h, TRUE);
  MoveWindow(s->btn_body_maximize, right_x + right_w - body_btn_w, y - ScalePx(s->hwnd, 4), body_btn_w, btn_h, TRUE);
  y += btn_h + ScalePx(s->hwnd, 2);

  // Formatting toolbar row (compact).
  int tbh = ScalePx(s->hwnd, 26);
  int tbw = ScalePx(s->hwnd, 30);
  int tb_pad = ScalePx(s->hwnd, 6);
  int tx = right_x;
  int ty = y;
  int wrap_right = right_x + right_w;
  for (HWND b : {s->fmt_bold, s->fmt_italic, s->fmt_underline, s->fmt_superscript, s->fmt_subscript,
                 s->fmt_numbering, s->fmt_bullet, s->fmt_markdown,
                 s->fmt_align_left, s->fmt_align_center, s->fmt_align_right, s->fmt_indent_dec, s->fmt_indent_inc,
                 s->fmt_clear}) {
    if (!b) continue;
    if (tx + tbw > wrap_right) {
      tx = right_x;
      ty += tbh + tb_pad;
    }
    MoveWindow(b, tx, ty, tbw, tbh, TRUE);
    tx += tbw + tb_pad;
  }
  y = ty + tbh + ScalePx(s->hwnd, 4);

  MoveWindow(s->ed_body, right_x, y, right_w, h - y - pad, TRUE);
}

static void InitListColumns(HWND list) {
  LVCOLUMNW col{};
  col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;

  col.pszText = (LPWSTR)L"日期范围";
  col.cx = 170;
  col.iSubItem = 0;
  ListView_InsertColumn(list, 0, &col);

  col.pszText = (LPWSTR)L"分类";
  col.cx = 120;
  col.iSubItem = 1;
  ListView_InsertColumn(list, 1, &col);

  col.pszText = (LPWSTR)L"标题";
  col.cx = 300;
  col.iSubItem = 2;
  ListView_InsertColumn(list, 2, &col);

  col.pszText = (LPWSTR)L"状态";
  col.cx = 90;
  col.iSubItem = 3;
  ListView_InsertColumn(list, 3, &col);

  col.pszText = (LPWSTR)L"目录";
  col.cx = 280;
  col.iSubItem = 4;
  ListView_InsertColumn(list, 4, &col);
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
  AppendMenuW(tools, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(tools, MF_STRING, IDM_EXPORT_FULL_DATA, L"导出全部数据...");
  AppendMenuW(tools, MF_STRING, IDM_IMPORT_FULL_DATA, L"导入全部数据...");
  AppendMenuW(tools, MF_STRING, IDM_CLEAR_ALL_DATA, L"清空全部数据(重置)...");
  AppendMenuW(tools, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(tools, MF_STRING, IDM_THEME_MUNG, L"绿豆沙主题");
  AppendMenuW(tools, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(tools, MF_STRING, IDM_MANAGE_CATEGORIES, L"管理分类...\tCtrl+K");
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

static std::vector<int> GetSelectedEntryIndices(AppState* s) {
  std::vector<int> out;
  if (!s || !s->list) return out;
  int idx = -1;
  while ((idx = ListView_GetNextItem(s->list, idx, LVNI_SELECTED)) >= 0) {
    if (idx < (int)s->day.entries.size()) out.push_back(idx);
  }
  return out;
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
                                WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL | LVS_REPORT | LVS_SHOWSELALWAYS,
                                0, 0, 10, 10, hwnd, (HMENU)IDC_LIST, nullptr, nullptr);
      ListView_SetExtendedListViewStyle(s->list, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_DOUBLEBUFFER);
      InitListColumns(s->list);

      s->lbl_category = CreateWindowExW(0, L"STATIC", L"分类(必填)",
                                        WS_CHILD | WS_VISIBLE,
                                        0, 0, 10, 10, hwnd, (HMENU)IDC_LBL_CATEGORY, nullptr, nullptr);
      s->lbl_start = CreateWindowExW(0, L"STATIC", L"开始日期",
                                     WS_CHILD | WS_VISIBLE,
                                     0, 0, 10, 10, hwnd, (HMENU)IDC_LBL_START, nullptr, nullptr);
      s->lbl_end = CreateWindowExW(0, L"STATIC", L"结束日期",
                                   WS_CHILD | WS_VISIBLE,
                                   0, 0, 10, 10, hwnd, (HMENU)IDC_LBL_END, nullptr, nullptr);
      s->lbl_status = CreateWindowExW(0, L"STATIC", L"状态",
                                      WS_CHILD | WS_VISIBLE,
                                      0, 0, 10, 10, hwnd, (HMENU)IDC_LBL_STATUS, nullptr, nullptr);

      s->cb_category = CreateWindowExW(0, L"COMBOBOX", L"",
                                       WS_CHILD | WS_VISIBLE | CBS_DROPDOWN | WS_TABSTOP | WS_VSCROLL,
                                       0, 0, 10, 10, hwnd, (HMENU)IDC_CATEGORY, nullptr, nullptr);

      s->ed_start = CreateWindowExW(0, DATETIMEPICK_CLASSW, L"",
                                    WS_CHILD | WS_VISIBLE | DTS_SHORTDATEFORMAT | WS_TABSTOP,
                                    0, 0, 10, 10, hwnd, (HMENU)IDC_START, nullptr, nullptr);
      s->ed_end = CreateWindowExW(0, DATETIMEPICK_CLASSW, L"",
                                  WS_CHILD | WS_VISIBLE | DTS_SHORTDATEFORMAT | WS_TABSTOP,
                                  0, 0, 10, 10, hwnd, (HMENU)IDC_END, nullptr, nullptr);

      s->cb_status = CreateWindowExW(0, L"COMBOBOX", L"",
                                     WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP,
                                     0, 0, 10, 10, hwnd, (HMENU)IDC_CB_STATUS, nullptr, nullptr);
      SendMessageW(s->cb_status, CB_ADDSTRING, 0, (LPARAM)L"无");
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

      // Formatting toolbar (kept right above the body to make formatting discoverable).
      DWORD tool_btn_style = WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_TABSTOP;
      s->fmt_bold = CreateWindowExW(0, L"BUTTON", L"B",
                                    tool_btn_style,
                                    0, 0, 10, 10, hwnd, (HMENU)IDC_FMT_BOLD, nullptr, nullptr);
      s->fmt_italic = CreateWindowExW(0, L"BUTTON", L"I",
                                      tool_btn_style,
                                      0, 0, 10, 10, hwnd, (HMENU)IDC_FMT_ITALIC, nullptr, nullptr);
      s->fmt_underline = CreateWindowExW(0, L"BUTTON", L"U",
                                         tool_btn_style,
                                         0, 0, 10, 10, hwnd, (HMENU)IDC_FMT_UNDERLINE, nullptr, nullptr);
      s->fmt_superscript = CreateWindowExW(0, L"BUTTON", L"x2",
                                           tool_btn_style,
                                           0, 0, 10, 10, hwnd, (HMENU)IDC_FMT_SUPERSCRIPT, nullptr, nullptr);
      s->fmt_subscript = CreateWindowExW(0, L"BUTTON", L"x_",
                                         tool_btn_style,
                                         0, 0, 10, 10, hwnd, (HMENU)IDC_FMT_SUBSCRIPT, nullptr, nullptr);
      s->fmt_numbering = CreateWindowExW(0, L"BUTTON", L"1.",
                                         tool_btn_style,
                                         0, 0, 10, 10, hwnd, (HMENU)IDC_FMT_NUMBERING, nullptr, nullptr);
      s->fmt_bullet = CreateWindowExW(0, L"BUTTON", L"•",
                                      tool_btn_style,
                                      0, 0, 10, 10, hwnd, (HMENU)IDC_FMT_BULLET, nullptr, nullptr);
      s->fmt_markdown = CreateWindowExW(0, L"BUTTON", L"MD",
                                        tool_btn_style,
                                        0, 0, 10, 10, hwnd, (HMENU)IDC_FMT_MARKDOWN, nullptr, nullptr);
      s->fmt_align_left = CreateWindowExW(0, L"BUTTON", L"L",
                                          tool_btn_style,
                                          0, 0, 10, 10, hwnd, (HMENU)IDC_FMT_ALIGN_LEFT, nullptr, nullptr);
      s->fmt_align_center = CreateWindowExW(0, L"BUTTON", L"C",
                                            tool_btn_style,
                                            0, 0, 10, 10, hwnd, (HMENU)IDC_FMT_ALIGN_CENTER, nullptr, nullptr);
      s->fmt_align_right = CreateWindowExW(0, L"BUTTON", L"R",
                                           tool_btn_style,
                                           0, 0, 10, 10, hwnd, (HMENU)IDC_FMT_ALIGN_RIGHT, nullptr, nullptr);
      s->fmt_indent_inc = CreateWindowExW(0, L"BUTTON", L">",
                                          tool_btn_style,
                                          0, 0, 10, 10, hwnd, (HMENU)IDC_FMT_INDENT_INC, nullptr, nullptr);
      s->fmt_indent_dec = CreateWindowExW(0, L"BUTTON", L"<",
                                          tool_btn_style,
                                          0, 0, 10, 10, hwnd, (HMENU)IDC_FMT_INDENT_DEC, nullptr, nullptr);
      s->fmt_clear = CreateWindowExW(0, L"BUTTON", L"清",
                                     tool_btn_style,
                                     0, 0, 10, 10, hwnd, (HMENU)IDC_FMT_CLEAR, nullptr, nullptr);
      s->btn_body_maximize = CreateWindowExW(0, L"BUTTON", L"最大化编辑",
                                             tool_btn_style,
                                             0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_BODY_MAXIMIZE, nullptr, nullptr);

      s->ed_body = CreateWindowExW(WS_EX_CLIENTEDGE, L"RICHEDIT50W", L"",
                                   WS_CHILD | WS_VISIBLE | ES_MULTILINE | ES_AUTOVSCROLL | ES_DISABLENOSCROLL | ES_NOHIDESEL |
                                       WS_VSCROLL | WS_TABSTOP,
                                   0, 0, 10, 10, hwnd, (HMENU)IDC_BODY, nullptr, nullptr);

      s->btn_new = CreateWindowExW(0, L"BUTTON", L"新增",
                                   tool_btn_style,
                                   0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_NEW, nullptr, nullptr);
      s->btn_save = CreateWindowExW(0, L"BUTTON", L"保存",
                                    tool_btn_style,
                                    0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_SAVE, nullptr, nullptr);
      s->btn_preview = CreateWindowExW(0, L"BUTTON", L"预览",
                                       tool_btn_style,
                                       0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_PREVIEW, nullptr, nullptr);
      s->btn_del = CreateWindowExW(0, L"BUTTON", L"删除",
                                   tool_btn_style,
                                   0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_DEL, nullptr, nullptr);

      s->status = CreateWindowExW(0, STATUSCLASSNAMEW, nullptr,
                                  WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP,
                                  0, 0, 0, 0, hwnd, (HMENU)IDC_STATUS, nullptr, nullptr);

      // Left-bottom "office" helpers.
      s->grp_office = CreateWindowExW(0, L"BUTTON", L"工作区",
                                      WS_CHILD | WS_VISIBLE | BS_GROUPBOX,
                                      0, 0, 10, 10, hwnd, (HMENU)IDC_GRP_OFFICE, nullptr, nullptr);
      DWORD office_btn_style = WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_TABSTOP;
      s->btn_screenshot = CreateWindowExW(0, L"BUTTON", L"截图",
                                          office_btn_style,
                                          0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_SCREENSHOT, nullptr, nullptr);
      s->btn_quick_reply = CreateWindowExW(0, L"BUTTON", L"常用回复",
                                           office_btn_style,
                                           0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_QUICK_REPLY, nullptr, nullptr);
      s->btn_timestamp = CreateWindowExW(0, L"BUTTON", L"复制时间戳",
                                         office_btn_style,
                                         0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_TIMESTAMP, nullptr, nullptr);
      s->btn_focus_timer = CreateWindowExW(0, L"BUTTON", L"专注 25:00",
                                           office_btn_style,
                                           0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_FOCUS_TIMER, nullptr, nullptr);
      s->btn_calc = CreateWindowExW(0, L"BUTTON", L"计算器",
                                    office_btn_style,
                                    0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_CALC, nullptr, nullptr);
      s->btn_data_dir = CreateWindowExW(0, L"BUTTON", L"数据目录",
                                        office_btn_style,
                                        0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_DATA_DIR, nullptr, nullptr);
      s->btn_screenshot_dir = CreateWindowExW(0, L"BUTTON", L"截图目录",
                                              office_btn_style,
                                              0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_SCREENSHOT_DIR, nullptr, nullptr);
      s->btn_task_materials = CreateWindowExW(0, L"BUTTON", L"便签",
                                              office_btn_style,
                                              0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_TASK_MATERIALS_ROOT, nullptr, nullptr);
      s->btn_open_materials = CreateWindowExW(0, L"BUTTON", L"日志目录",
                                              office_btn_style,
                                              0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_OPEN_MATERIALS, nullptr, nullptr);
      s->btn_set_materials = CreateWindowExW(0, L"BUTTON", L"材料根目录",
                                             office_btn_style,
                                             0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_SET_MATERIALS, nullptr, nullptr);
      s->btn_color_picker = CreateWindowExW(0, L"BUTTON", L"取色器",
                                            office_btn_style,
                                            0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_COLOR_PICKER, nullptr, nullptr);
      s->btn_paste_plain = CreateWindowExW(0, L"BUTTON", L"纯贴",
                                           office_btn_style,
                                           0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_PASTE_PLAIN, nullptr, nullptr);
      s->btn_date_separator = CreateWindowExW(0, L"BUTTON", L"时间分隔",
                                              office_btn_style,
                                              0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_DATE_SEPARATOR, nullptr, nullptr);
      s->btn_record_system = CreateWindowExW(0, L"BUTTON", L"录系统声",
                                             office_btn_style,
                                             0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_RECORD_SYSTEM, nullptr, nullptr);
      s->btn_record_mic = CreateWindowExW(0, L"BUTTON", L"录麦克风",
                                          office_btn_style,
                                          0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_RECORD_MIC, nullptr, nullptr);
      s->btn_recordings_dir = CreateWindowExW(0, L"BUTTON", L"录音目录",
                                              office_btn_style,
                                              0, 0, 10, 10, hwnd, (HMENU)IDC_BTN_RECORDINGS_DIR, nullptr, nullptr);

      // Fonts
      for (HWND h : {s->cal, s->st_day, s->list, s->lbl_category, s->lbl_start, s->lbl_end, s->lbl_title, s->lbl_body,
                     s->lbl_status, s->cb_category, s->ed_start, s->ed_end, s->cb_status, s->ed_title,
                     s->fmt_bold, s->fmt_italic, s->fmt_underline, s->fmt_superscript, s->fmt_subscript,
                     s->fmt_numbering, s->fmt_bullet, s->fmt_markdown,
                     s->fmt_align_left, s->fmt_align_center, s->fmt_align_right, s->fmt_indent_inc,
                     s->fmt_indent_dec, s->fmt_clear,
                     s->btn_body_maximize,
                     s->ed_body, s->status,
                       s->btn_new, s->btn_save, s->btn_preview, s->btn_del,
                     s->grp_office, s->btn_screenshot, s->btn_quick_reply, s->btn_timestamp, s->btn_focus_timer, s->btn_calc,
                      s->btn_data_dir, s->btn_screenshot_dir, s->btn_task_materials, s->btn_open_materials, s->btn_set_materials,
                      s->btn_color_picker, s->btn_paste_plain, s->btn_date_separator,
                      s->btn_record_system, s->btn_record_mic, s->btn_recordings_dir}) {
        SetControlFont(h, s->font);
      }

      SetMenu(hwnd, CreateAppMenu());
      ApplyTheme(s);

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
      // Ensure every long-term task has a stable materials directory, including after import/upgrade.
      for (const auto& t : s->tasks) {
        if (!t.id.empty()) (void)GetTaskMaterialsDirById(t.id);
      }


      // Field hints
      SetEditCue(GetComboEditHandle(s->cb_category), L"例如: 工作 / 会议 / 目标 / 沟通 / 研发");
      SetEditCue(s->ed_title, L"一句话概括今天这件事");
      SetEditCue(s->ed_body, L"富文本正文。可直接输入 Markdown 后点 MD。快捷键: Ctrl+B/I/U, Ctrl+滚轮缩放, Ctrl+0 重置, Ctrl+Shift+7 编号, Ctrl+Shift+8 项目符号, Ctrl+Alt+L/E/R 段落对齐");
      SetWindowSubclass(s->ed_body, BodyEditSubclassProc, 1, (DWORD_PTR)s);
      ApplyBodyZoom(s);
      UpdateAudioRecordButtons(s);

      SYSTEMTIME st{};
      GetLocalTime(&st);
      MonthCal_SetCurSel(s->cal, &st);
      LoadSelectedDay(s, st);
      Layout(s);
      UpdateStatusParts(s);
      UpdateStatusBar(s);
      return 0;
    }

    case WM_GETMINMAXINFO: {
      ApplyMinTrackSize(hwnd, s, reinterpret_cast<MINMAXINFO*>(lParam));
      return 0;
    }

    case WM_SIZE: {
      if (s && s->status) SendMessageW(s->status, WM_SIZE, 0, 0);
      if (s) UpdateStatusParts(s);
      if (s) Layout(s);
      if (s) UpdateStatusBar(s);
      return 0;
    }

    case WM_ERASEBKGND: {
      if (s && s->brush_window) {
        RECT rc{};
        GetClientRect(hwnd, &rc);
        FillRect((HDC)wParam, &rc, s->brush_window);
        return 1;
      }
      break;
    }

    case WM_CTLCOLORSTATIC:
    case WM_CTLCOLOREDIT:
    case WM_CTLCOLORLISTBOX:
    case WM_CTLCOLORBTN: {
      if (!s || !s->mung_theme) break;
      HDC dc = (HDC)wParam;
      HWND ctl = (HWND)lParam;
      SetTextColor(dc, ThemeTextColor(s));

      if (msg == WM_CTLCOLOREDIT || msg == WM_CTLCOLORLISTBOX) {
        SetBkColor(dc, ThemeEditColor(s));
        return (LRESULT)s->brush_edit;
      }

      COLORREF bg = ThemeWindowColor(s);
      HBRUSH br = s->brush_window;
      if (msg == WM_CTLCOLORBTN || ctl == s->grp_office) {
        bg = ThemePanelColor(s);
        br = s->brush_panel;
      }
      SetBkColor(dc, bg);
      return (LRESULT)br;
    }

    case WM_TIMER: {
      if (!s) return 0;
      if (s->focus_running && wParam == s->focus_timer_id) {
        ULONGLONG now = GetTickCount64();
        if (now >= s->focus_end_tick) {
          StopFocusTimer(s);
          MessageBeep(MB_ICONASTERISK);
          ShowToast(s, L"专注计时结束");
          MessageBoxW(hwnd, L"专注计时结束。", kAppTitle, MB_OK | MB_ICONINFORMATION);
          return 0;
        }
        UpdateFocusTimerButton(s);
        UpdateStatusBar(s);
        return 0;
      }
      if (s->audio_timer_id && wParam == s->audio_timer_id) {
        if (MaybeAutoStopAudioRecording(s, AudioRecordSource::SystemLoopback)) return 0;
        if (MaybeAutoStopAudioRecording(s, AudioRecordSource::MicrophoneCapture)) return 0;
        UpdateAudioRecordButtons(s);
        UpdateStatusBar(s);
        return 0;
      }
      if (s->autosave_timer_id && wParam == s->autosave_timer_id) {
        KillTimer(s->hwnd, s->autosave_timer_id);
        s->autosave_timer_id = 0;
        if (s->editing_index >= 0 && IsEditorTrulyDirty(s)) {
          if (AutoSaveCurrent(s)) ShowToast(s, L"已自动保存");
        }
        return 0;
      }
      return 0;
    }

    case WM_MOUSEWHEEL: {
      if (!s) return DefWindowProcW(hwnd, msg, wParam, lParam);
      bool ctrl = (GET_KEYSTATE_WPARAM(wParam) & MK_CONTROL) != 0;
      if (!ctrl) return DefWindowProcW(hwnd, msg, wParam, lParam);
      HWND focus = GetFocus();
      if (focus != s->ed_body) return DefWindowProcW(hwnd, msg, wParam, lParam);
      short delta = GET_WHEEL_DELTA_WPARAM(wParam);
      if (delta == 0) return 0;
      AdjustBodyZoomFromWheel(s, delta);
      UpdateStatusBar(s);
      return 0;
    }

    case WM_DRAWITEM: {
      int id = (int)wParam;
      if (id == IDC_BTN_SCREENSHOT || id == IDC_BTN_QUICK_REPLY || id == IDC_BTN_TIMESTAMP || id == IDC_BTN_FOCUS_TIMER ||
          id == IDC_BTN_CALC || id == IDC_BTN_DATA_DIR || id == IDC_BTN_SCREENSHOT_DIR || id == IDC_BTN_TASK_MATERIALS_ROOT ||
          id == IDC_BTN_OPEN_MATERIALS || id == IDC_BTN_SET_MATERIALS || id == IDC_BTN_COLOR_PICKER ||
          id == IDC_BTN_PASTE_PLAIN || id == IDC_BTN_DATE_SEPARATOR || id == IDC_BTN_RECORD_SYSTEM || id == IDC_BTN_RECORD_MIC ||
          id == IDC_BTN_RECORDINGS_DIR || id == IDC_FMT_BOLD || id == IDC_FMT_ITALIC || id == IDC_FMT_UNDERLINE ||
          id == IDC_FMT_SUPERSCRIPT || id == IDC_FMT_SUBSCRIPT || id == IDC_FMT_NUMBERING || id == IDC_FMT_BULLET ||
          id == IDC_FMT_MARKDOWN || id == IDC_FMT_ALIGN_LEFT || id == IDC_FMT_ALIGN_CENTER || id == IDC_FMT_ALIGN_RIGHT ||
          id == IDC_FMT_INDENT_INC || id == IDC_FMT_INDENT_DEC || id == IDC_FMT_CLEAR || id == IDC_BTN_BODY_MAXIMIZE ||
          id == IDC_BTN_NEW || id == IDC_BTN_SAVE || id == IDC_BTN_PREVIEW || id == IDC_BTN_DEL) {
        DrawOfficeButton(s, reinterpret_cast<const DRAWITEMSTRUCT*>(lParam));
        return TRUE;
      }
      break;
    }

    case WM_COMMAND: {
      int id = LOWORD(wParam);
      int code = HIWORD(wParam);

      if (id == IDC_BTN_NEW && code == BN_CLICKED) {
        if (!PromptSaveIfDirty(s)) return 0;
        ClearEditor(s, true);
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
        std::vector<int> selected = GetSelectedEntryIndices(s);
        int count = (int)selected.size();
        if (count <= 0 && s->editing_index >= 0 && s->editing_index < (int)s->day.entries.size()) count = 1;
        std::wstring prompt = count > 1 ? (L"确定删除当前选中的 " + std::to_wstring(count) + L" 条记录？")
                                        : L"确定删除当前条目？";
        int r = MessageBoxW(hwnd, prompt.c_str(), kAppTitle, MB_YESNO | MB_ICONWARNING);
        if (r == IDYES) DeleteCurrent(s);
        return 0;
      }
      if (id == IDC_BTN_SCREENSHOT && code == BN_CLICKED) {
        StartScreenshotTool(s->hwnd);
        return 0;
      }
      if (id == IDC_BTN_QUICK_REPLY && code == BN_CLICKED) {
        ShowQuickReplyMenu(s);
        return 0;
      }
      if (id == IDC_BTN_TIMESTAMP && code == BN_CLICKED) {
        SYSTEMTIME st{};
        GetLocalTime(&st);
        wchar_t buf[64]{};
        wsprintfW(buf, L"%04d-%02d-%02d %02d:%02d", (int)st.wYear, (int)st.wMonth, (int)st.wDay, (int)st.wHour, (int)st.wMinute);
        CopyToClipboard(s->hwnd, buf);
        ShowToast(s, L"时间戳已复制到剪贴板");
        return 0;
      }
      if (id == IDC_BTN_FOCUS_TIMER && code == BN_CLICKED) {
        if (s->focus_running) {
          StopFocusTimer(s);
          ShowToast(s, L"已停止专注计时");
        } else {
          StartFocusTimer(s, 25);
          ShowToast(s, L"已开始专注计时 25 分钟");
        }
        return 0;
      }
      if (id == IDC_BTN_CALC && code == BN_CLICKED) {
        ShowUnitCalcWindow(s->hwnd);
        return 0;
      }
      if (id == IDC_BTN_DATA_DIR && code == BN_CLICKED) {
        OpenDataDir(s);
        return 0;
      }
      if (id == IDC_BTN_SCREENSHOT_DIR && code == BN_CLICKED) {
        std::wstring dir = JoinPath(GetDataRootDir(), L"screenshots");
        EnsureDirExists(dir);
        OpenPathInShell(s->hwnd, dir, L"截图目录");
        return 0;
      }
      if (id == IDC_BTN_RECORDINGS_DIR && code == BN_CLICKED) {
        std::wstring sys_dir = GetAudioRecordingsDir(AudioRecordSource::SystemLoopback);
        std::wstring mic_dir = GetAudioRecordingsDir(AudioRecordSource::MicrophoneCapture);
        EnsureDirExists(sys_dir);
        EnsureDirExists(mic_dir);
        std::wstring dir = ParentDirOf(sys_dir);
        if (dir.empty()) dir = GetDataRootDir();
        OpenPathInShell(s->hwnd, dir, L"录音目录");
        return 0;
      }
      if (id == IDC_BTN_TASK_MATERIALS_ROOT && code == BN_CLICKED) {
        OpenScratchpadFile(s->hwnd);
        return 0;
      }
      if (id == IDC_BTN_OPEN_MATERIALS && code == BN_CLICKED) {
        OpenSelectedDayFolder(s);
        return 0;
      }
      if (id == IDC_BTN_SET_MATERIALS && code == BN_CLICKED) {
        std::wstring dir = GetWorkMaterialsRootDir();
        OpenPathInShell(s->hwnd, dir, L"材料根目录");
        return 0;
      }
      if (id == IDC_BTN_COLOR_PICKER && code == BN_CLICKED) {
        PickBodyTextColor(s);
        return 0;
      }
      if (id == IDC_BTN_PASTE_PLAIN && code == BN_CLICKED) {
        PastePlainTextIntoBody(s);
        return 0;
      }
      if (id == IDC_BTN_DATE_SEPARATOR && code == BN_CLICKED) {
        InsertDateSeparatorIntoBody(s);
        return 0;
      }
      if (id == IDC_BTN_RECORD_SYSTEM && code == BN_CLICKED) {
        ToggleAudioRecording(s, AudioRecordSource::SystemLoopback);
        return 0;
      }
      if (id == IDC_BTN_RECORD_MIC && code == BN_CLICKED) {
        ToggleAudioRecording(s, AudioRecordSource::MicrophoneCapture);
        return 0;
      }

      if (code == BN_CLICKED) {
        if (id == IDC_FMT_BOLD) {
          ApplyBodyFormatting(s, BodyFmtBold);
          return 0;
        }
        if (id == IDC_FMT_ITALIC) {
          ApplyBodyFormatting(s, BodyFmtItalic);
          return 0;
        }
        if (id == IDC_FMT_UNDERLINE) {
          ApplyBodyFormatting(s, BodyFmtUnderline);
          return 0;
        }
        if (id == IDC_FMT_SUPERSCRIPT) {
          ApplyBodyFormatting(s, BodyFmtSuperscript);
          return 0;
        }
        if (id == IDC_FMT_SUBSCRIPT) {
          ApplyBodyFormatting(s, BodyFmtSubscript);
          return 0;
        }
        if (id == IDC_FMT_NUMBERING) {
          ApplyBodyFormatting(s, BodyFmtNumbering);
          return 0;
        }
        if (id == IDC_FMT_BULLET) {
          ApplyBodyFormatting(s, BodyFmtBullet);
          return 0;
        }
        if (id == IDC_FMT_MARKDOWN) {
          ApplyMarkdownStylingToBody(s);
          return 0;
        }
        if (id == IDC_FMT_ALIGN_LEFT) {
          ApplyBodyFormatting2(s, BodyFmtAlign, (int)PFA_LEFT);
          return 0;
        }
        if (id == IDC_FMT_ALIGN_CENTER) {
          ApplyBodyFormatting2(s, BodyFmtAlign, (int)PFA_CENTER);
          return 0;
        }
        if (id == IDC_FMT_ALIGN_RIGHT) {
          ApplyBodyFormatting2(s, BodyFmtAlign, (int)PFA_RIGHT);
          return 0;
        }
        if (id == IDC_FMT_INDENT_INC) {
          ApplyBodyFormatting2(s, BodyFmtIndent, 360);
          return 0;
        }
        if (id == IDC_FMT_INDENT_DEC) {
          ApplyBodyFormatting2(s, BodyFmtIndent, -360);
          return 0;
        }
        if (id == IDC_FMT_CLEAR) {
          ApplyBodyFormatting(s, BodyFmtClear);
          return 0;
        }
        if (id == IDC_BTN_BODY_MAXIMIZE) {
          SetEditorMaximized(s, !s->editor_maximized);
          SetFocus(s->ed_body);
          return 0;
        }
      }

      if ((id == IDC_CATEGORY || id == IDC_START || id == IDC_END || id == IDC_CB_STATUS || id == IDC_TITLE || id == IDC_BODY) &&
          (code == CBN_SELCHANGE || code == EN_CHANGE)) {
        if (s && s->suppress_editor_change_tracking) return 0;
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
        ApplyBodyFormatting(s, BodyFmtBold);
        return 0;
      }
      if (id == IDM_FMT_ITALIC) {
        ApplyBodyFormatting(s, BodyFmtItalic);
        return 0;
      }
      if (id == IDM_FMT_UNDERLINE) {
        ApplyBodyFormatting(s, BodyFmtUnderline);
        return 0;
      }
      if (id == IDM_FMT_BULLET) {
        ApplyBodyFormatting(s, BodyFmtBullet);
        return 0;
      }
      if (id == IDM_FMT_NUMBERING) {
        ApplyBodyFormatting(s, BodyFmtNumbering);
        return 0;
      }
      if (id == IDM_FMT_ALIGN_LEFT) {
        ApplyBodyFormatting2(s, BodyFmtAlign, (int)PFA_LEFT);
        return 0;
      }
      if (id == IDM_FMT_ALIGN_CENTER) {
        ApplyBodyFormatting2(s, BodyFmtAlign, (int)PFA_CENTER);
        return 0;
      }
      if (id == IDM_FMT_ALIGN_RIGHT) {
        ApplyBodyFormatting2(s, BodyFmtAlign, (int)PFA_RIGHT);
        return 0;
      }
      if (id == IDM_FMT_INDENT_INC) {
        ApplyBodyFormatting2(s, BodyFmtIndent, 360);
        return 0;
      }
      if (id == IDM_FMT_INDENT_DEC) {
        ApplyBodyFormatting2(s, BodyFmtIndent, -360);
        return 0;
      }
      if (id == IDM_FMT_CLEAR) {
        ApplyBodyFormatting(s, BodyFmtClear);
        return 0;
      }
      if (id == IDM_OPEN_DATA_DIR) {
        OpenDataDir(s);
        return 0;
      }
      if (id == IDM_EXPORT_FULL_DATA) {
        ExportFullData(s);
        return 0;
      }
      if (id == IDM_IMPORT_FULL_DATA) {
        ImportFullData(s);
        return 0;
      }
      if (id == IDM_CLEAR_ALL_DATA) {
        ClearAllData(s);
        return 0;
      }
      if (id == IDM_THEME_MUNG) {
        s->mung_theme = !s->mung_theme;
        ApplyTheme(s);
        return 0;
      }
      if (id == IDM_MANAGE_CATEGORIES) {
        if (!PromptSaveIfDirty(s)) return 0;
        ShowManageCategoriesWindow(s);
        return 0;
      }
      if (id == IDM_CTX_OPEN_MATERIALS_DIR) {
        int sel = ListView_GetNextItem(s->list, -1, LVNI_SELECTED);
        if (sel < 0 || sel >= (int)s->day.entries.size()) return 0;
        const Entry& e = s->day.entries[(size_t)sel];
        OpenMaterialsDirForEntry(s->hwnd, s, e);
        return 0;
      }
      if (id == IDM_CTX_SET_MATERIALS_DIR) {
        SetMaterialsDirForSelected(s);
        return 0;
      }
      if (id == IDM_CTX_CLEAR_MATERIALS_DIR) {
        ClearMaterialsDirForSelected(s);
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
      }

      int sel = ListView_GetNextItem(s->list, -1, LVNI_SELECTED);
      // If selection change was blocked (e.g. user canceled the save prompt), don't show a menu that would act on
      // a different item than the one the user right-clicked.
      if (hit >= 0 && sel != hit) return 0;
      if (sel < 0 || sel >= (int)s->day.entries.size()) return 0;
      HMENU menu = CreatePopupMenu();
      if (!menu) return 0;

      AppendMenuW(menu, MF_STRING, IDM_CTX_OPEN_MATERIALS_DIR, L"打开材料目录");
      AppendMenuW(menu, MF_STRING, IDM_CTX_SET_MATERIALS_DIR, L"设置存放路径...");
      AppendMenuW(menu, MF_STRING, IDM_CTX_CLEAR_MATERIALS_DIR, L"清除存放路径(默认)");

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
        if (s->ignore_cal_selchange) {
          s->ignore_cal_selchange = false;
          return 0;
        }
        if (!PromptSaveIfDirty(s)) {
          s->ignore_cal_selchange = true;
          MonthCal_SetCurSel(s->cal, &s->selected);
          return 0;
        }
        auto* n = reinterpret_cast<NMSELCHANGE*>(lParam);
        LoadSelectedDay(s, n->stSelStart);
        return 0;
      }

      if ((hdr->idFrom == IDC_START || hdr->idFrom == IDC_END) && hdr->code == DTN_DATETIMECHANGE) {
        if (s->suppress_editor_change_tracking) return 0;
        SetEditorDirty(s, true);
        UpdateStatusBar(s);
        return 0;
      }

      if (hdr->idFrom == IDC_LIST) {
        if (hdr->code == LVN_COLUMNCLICK) {
          auto* lv = reinterpret_cast<NMLISTVIEW*>(lParam);
          if (!lv) return 0;
          int col = lv->iSubItem;
          if (col == 0 || col == 1 || col == 3) {
            if (s->list_sort_column == col) s->list_sort_desc = !s->list_sort_desc;
            else {
              s->list_sort_column = col;
              s->list_sort_desc = false;
            }
            RefreshList(s);
          }
          return 0;
        }
        if (hdr->code == NM_CUSTOMDRAW) {
          auto* cd = reinterpret_cast<NMLVCUSTOMDRAW*>(lParam);
          if (cd->nmcd.dwDrawStage == CDDS_PREPAINT) return CDRF_NOTIFYITEMDRAW;
          if (cd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
            int idx = (int)cd->nmcd.dwItemSpec;
            if (idx >= 0 && idx < (int)s->day.entries.size()) {
              const auto& e = s->day.entries[(size_t)idx];
              EntryStatus st = EffectiveStatus(s, e);
                COLORREF bk = ThemedStatusBkColor(s, st);
                if (bk != CLR_NONE) {
                  cd->clrTextBk = bk;
                  cd->clrText = ThemeTextColor(s);
                }
            }
            return CDRF_DODEFAULT;
          }
          return CDRF_DODEFAULT;
        }
        if (hdr->code == LVN_ITEMCHANGING) {
          if (s->suppress_list_selection_sync) return FALSE;
          auto* p = reinterpret_cast<NMLISTVIEW*>(lParam);
          if ((p->uChanged & LVIF_STATE) &&
              (p->uNewState & LVIS_SELECTED) && !(p->uOldState & LVIS_SELECTED)) {
            if (!PromptSaveIfDirty(s)) {
              // Returning TRUE cancels the selection change without triggering extra state churn.
              return TRUE;
            }
          }
          return FALSE;
        }
        if (hdr->code == LVN_ITEMCHANGED) {
          if (s->suppress_list_selection_sync) return 0;
          auto* p = reinterpret_cast<NMLISTVIEW*>(lParam);
          if ((p->uChanged & LVIF_STATE) && (p->uNewState & LVIS_SELECTED)) {
            UpdateEditorFromListSelection(s);
          }
          return 0;
        }
        if (hdr->code == NM_DBLCLK) {
          auto* act = reinterpret_cast<NMITEMACTIVATE*>(lParam);
          if (act && act->iItem < 0) {
            // Double click on empty space: start a new entry (common expectation).
            if (!PromptSaveIfDirty(s)) return 0;
            ClearEditor(s, true);
            HWND h = GetComboEditHandle(s->cb_category);
            SetFocus(h ? h : s->cb_category);
            return 0;
          }
          if (act && act->iItem >= 0 && act->iItem < (int)s->day.entries.size() && act->iSubItem == 4) {
            const Entry& e = s->day.entries[(size_t)act->iItem];
            OpenMaterialsDirForEntry(s->hwnd, s, e);
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
      if (s && s->close_in_progress) return 0;
      if (s) s->close_in_progress = true;
      if (s && !PromptSaveIfDirty(s)) {
        s->close_in_progress = false;
        return 0;
      }
      DestroyWindow(hwnd);
      return 0;
    }

    case WM_DESTROY: {
      if (s) StopFocusTimer(s);
      if (s) {
        if (s->autosave_timer_id) {
          KillTimer(s->hwnd, s->autosave_timer_id);
          s->autosave_timer_id = 0;
        }
        StopAudioRecorderWithoutPrompt(s, AudioRecordSource::SystemLoopback);
        StopAudioRecorderWithoutPrompt(s, AudioRecordSource::MicrophoneCapture);
        EnsureAudioTimerState(s);
        if (s->brush_window) DeleteObject(s->brush_window);
        if (s->brush_panel) DeleteObject(s->brush_panel);
        if (s->brush_edit) DeleteObject(s->brush_edit);
      }
      if (s && s->font) DeleteObject(s->font);
      PostQuitMessage(0);
      return 0;
    }
  }
  return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static void TryEnablePerMonitorDpiAwareness() {
  // Fix DPI virtualization issues (e.g. screenshot overlay "zooming" on high-DPI displays).
  // Use dynamic lookup to keep compatibility with older Windows versions.
  HMODULE user32 = GetModuleHandleW(L"user32.dll");
  if (user32) {
    using FnSetCtx = BOOL(WINAPI*)(HANDLE);
    auto pSetCtx = reinterpret_cast<FnSetCtx>(GetProcAddress(user32, "SetProcessDpiAwarenessContext"));
    if (pSetCtx) {
      // DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 == (HANDLE)-4
      if (pSetCtx(reinterpret_cast<HANDLE>(-4))) return;
      // DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE == (HANDLE)-3
      if (pSetCtx(reinterpret_cast<HANDLE>(-3))) return;
    }
  }

  // Windows 8.1: SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE=2)
  HMODULE shcore = LoadLibraryW(L"shcore.dll");
  if (shcore) {
    using FnSetAw = HRESULT(WINAPI*)(int);
    auto pSetAw = reinterpret_cast<FnSetAw>(GetProcAddress(shcore, "SetProcessDpiAwareness"));
    if (pSetAw) {
      (void)pSetAw(2);
      FreeLibrary(shcore);
      return;
    }
    FreeLibrary(shcore);
  }

  // Vista+: system DPI aware.
  if (user32) {
    using FnSetAware = BOOL(WINAPI*)();
    auto p = reinterpret_cast<FnSetAware>(GetProcAddress(user32, "SetProcessDPIAware"));
    if (p) (void)p();
  }
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int nCmdShow) {
  TryEnablePerMonitorDpiAwareness();

  INITCOMMONCONTROLSEX icc{};
  icc.dwSize = sizeof(icc);
  icc.dwICC = ICC_DATE_CLASSES | ICC_LISTVIEW_CLASSES | ICC_STANDARD_CLASSES | ICC_BAR_CLASSES;
  InitCommonControlsEx(&icc);

  // RichEdit (msftedit) is a Windows component, loaded on demand.
  TryLoadRichEditMsftedit();

  AppState state{};

  HICON app_icon = (HICON)LoadImageW(hInstance, MAKEINTRESOURCEW(IDI_APP_ICON), IMAGE_ICON, 0, 0, LR_DEFAULTSIZE);
  HICON app_icon_sm = (HICON)LoadImageW(hInstance, MAKEINTRESOURCEW(IDI_APP_ICON), IMAGE_ICON,
                                        GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), 0);

  WNDCLASSEXW wc{};
  wc.cbSize = sizeof(wc);
  wc.lpfnWndProc = MainWndProc;
  wc.hInstance = hInstance;
  wc.lpszClassName = L"WorkLogLiteMainWnd";
  wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
  wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
  wc.hIcon = app_icon;
  wc.hIconSm = app_icon_sm ? app_icon_sm : app_icon;

  if (!RegisterClassExW(&wc)) return 0;

  // Size the initial window based on system DPI so that scaled controls are visible without requiring a manual resize.
  UINT sys_dpi = QuerySystemDpiCompat();
  int init_w = MulDiv(1100, (int)sys_dpi, 96);
  int init_h = MulDiv(760, (int)sys_dpi, 96);
  RECT wa{};
  if (SystemParametersInfoW(SPI_GETWORKAREA, 0, &wa, 0)) {
    int ww = wa.right - wa.left;
    int wh = wa.bottom - wa.top;
    if (init_w > ww) init_w = ww;
    if (init_h > wh) init_h = wh;
    if (init_w < 800) init_w = 800;
    if (init_h < 560) init_h = 560;
  }

  HWND hwnd = CreateWindowExW(0, wc.lpszClassName, kAppTitle,
                              WS_OVERLAPPEDWINDOW | WS_VISIBLE,
                              CW_USEDEFAULT, CW_USEDEFAULT, init_w, init_h,
                              nullptr, nullptr, hInstance, &state);
  if (!hwnd) return 0;
  if (app_icon) SendMessageW(hwnd, WM_SETICON, ICON_BIG, (LPARAM)app_icon);
  if (wc.hIconSm) SendMessageW(hwnd, WM_SETICON, ICON_SMALL, (LPARAM)wc.hIconSm);

  // Ensure initial size meets our computed minimum (DPI aware). Clamp to work area to avoid going off-screen.
  {
    int min_cw = 0, min_ch = 0;
    ComputeMainMinClientSize(hwnd, &state, &min_cw, &min_ch);

    RECT cc{};
    GetClientRect(hwnd, &cc);
    int cur_cw = cc.right - cc.left;
    int cur_ch = cc.bottom - cc.top;

    if (cur_cw < min_cw || cur_ch < min_ch) {
      RECT wr{0, 0, (cur_cw < min_cw ? min_cw : cur_cw), (cur_ch < min_ch ? min_ch : cur_ch)};
      DWORD style = (DWORD)GetWindowLongPtrW(hwnd, GWL_STYLE);
      DWORD ex_style = (DWORD)GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
      BOOL has_menu = GetMenu(hwnd) != nullptr;
      UINT dpi = QueryDpiForWindowCompat(hwnd);
      AdjustWindowRectExForDpiCompat(&wr, style, has_menu, ex_style, dpi);
      int want_w = wr.right - wr.left;
      int want_h = wr.bottom - wr.top;

      RECT wa2{};
      if (SystemParametersInfoW(SPI_GETWORKAREA, 0, &wa2, 0)) {
        int ww = wa2.right - wa2.left;
        int wh = wa2.bottom - wa2.top;
        if (want_w > ww) want_w = ww;
        if (want_h > wh) want_h = wh;
      }

      SetWindowPos(hwnd, nullptr, 0, 0, want_w, want_h, SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
    }
  }

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
          if (focus == state.ed_body) {
            if (shift && msg.wParam == 'M') {
              ApplyMarkdownStylingToBody(&state);
              continue;
            }
            if (msg.wParam == '0') {
              SetBodyZoom(&state, 100);
              continue;
            }
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

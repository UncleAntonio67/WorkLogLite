#include "color_picker.h"

#include <windowsx.h>

#include <string>

namespace {

struct ColorPickerState {
  HWND hwnd{};
  HWND owner{};
  HBITMAP bitmap{};
  HDC memdc{};
  int vx{};
  int vy{};
  int vw{};
  int vh{};
  POINT cursor{};
  COLORREF color{RGB(0, 0, 0)};
  bool accepted{false};
};

static const wchar_t* kColorPickerClass = L"WorkLogLiteColorPickerWindow";

template <typename T>
static void DeleteGdiObject(T obj) {
  if (obj) DeleteObject(obj);
}

static bool CaptureVirtualScreen(ColorPickerState* s) {
  if (!s) return false;
  s->vx = GetSystemMetrics(SM_XVIRTUALSCREEN);
  s->vy = GetSystemMetrics(SM_YVIRTUALSCREEN);
  s->vw = GetSystemMetrics(SM_CXVIRTUALSCREEN);
  s->vh = GetSystemMetrics(SM_CYVIRTUALSCREEN);
  if (s->vw <= 0 || s->vh <= 0) return false;

  HDC screen = GetDC(nullptr);
  if (!screen) return false;
  HDC mem = CreateCompatibleDC(screen);
  HBITMAP bmp = CreateCompatibleBitmap(screen, s->vw, s->vh);
  if (!mem || !bmp) {
    if (bmp) DeleteObject(bmp);
    if (mem) DeleteDC(mem);
    ReleaseDC(nullptr, screen);
    return false;
  }

  SelectObject(mem, bmp);
  BOOL ok = BitBlt(mem, 0, 0, s->vw, s->vh, screen, s->vx, s->vy, SRCCOPY | CAPTUREBLT);
  ReleaseDC(nullptr, screen);
  if (!ok) {
    DeleteObject(bmp);
    DeleteDC(mem);
    return false;
  }

  s->bitmap = bmp;
  s->memdc = mem;
  return true;
}

static void UpdateColorUnderCursor(ColorPickerState* s) {
  if (!s || !s->memdc) return;
  int x = s->cursor.x - s->vx;
  int y = s->cursor.y - s->vy;
  if (x < 0) x = 0;
  if (y < 0) y = 0;
  if (x >= s->vw) x = s->vw - 1;
  if (y >= s->vh) y = s->vh - 1;
  s->color = GetPixel(s->memdc, x, y);
}

static std::wstring ColorToHex(COLORREF color) {
  wchar_t buf[16]{};
  wsprintfW(buf, L"#%02X%02X%02X", GetRValue(color), GetGValue(color), GetBValue(color));
  return std::wstring(buf);
}

static void DrawOverlay(ColorPickerState* s, HDC dc) {
  if (!s || !dc || !s->memdc) return;
  BitBlt(dc, 0, 0, s->vw, s->vh, s->memdc, 0, 0, SRCCOPY);

  const int cx = s->cursor.x - s->vx;
  const int cy = s->cursor.y - s->vy;

  HPEN cross = CreatePen(PS_SOLID, 1, RGB(255, 255, 255));
  HPEN cross_dark = CreatePen(PS_SOLID, 1, RGB(25, 25, 25));
  HGDIOBJ old_pen = SelectObject(dc, cross_dark);
  MoveToEx(dc, cx - 12, cy, nullptr);
  LineTo(dc, cx + 13, cy);
  MoveToEx(dc, cx, cy - 12, nullptr);
  LineTo(dc, cx, cy + 13);
  SelectObject(dc, cross);
  MoveToEx(dc, cx - 11, cy, nullptr);
  LineTo(dc, cx + 12, cy);
  MoveToEx(dc, cx, cy - 11, nullptr);
  LineTo(dc, cx, cy + 12);
  SelectObject(dc, old_pen);
  DeleteObject(cross);
  DeleteObject(cross_dark);

  RECT panel{cx + 18, cy + 18, cx + 210, cy + 78};
  if (panel.right > s->vw - 8) OffsetRect(&panel, -(panel.right - (s->vw - 8)), 0);
  if (panel.bottom > s->vh - 8) OffsetRect(&panel, 0, -(panel.bottom - (s->vh - 8)));
  if (panel.left < 8) OffsetRect(&panel, 8 - panel.left, 0);
  if (panel.top < 8) OffsetRect(&panel, 0, 8 - panel.top);

  HBRUSH bg = CreateSolidBrush(RGB(250, 250, 250));
  HPEN border = CreatePen(PS_SOLID, 1, RGB(190, 190, 190));
  HGDIOBJ old_br = SelectObject(dc, bg);
  old_pen = SelectObject(dc, border);
  RoundRect(dc, panel.left, panel.top, panel.right, panel.bottom, 10, 10);
  SelectObject(dc, old_pen);
  SelectObject(dc, old_br);
  DeleteObject(bg);
  DeleteObject(border);

  RECT swatch{panel.left + 10, panel.top + 10, panel.left + 54, panel.top + 50};
  HBRUSH sw = CreateSolidBrush(s->color);
  FrameRect(dc, &swatch, (HBRUSH)GetStockObject(BLACK_BRUSH));
  InflateRect(&swatch, -1, -1);
  FillRect(dc, &swatch, sw);
  DeleteObject(sw);

  SetBkMode(dc, TRANSPARENT);
  SetTextColor(dc, RGB(25, 25, 25));
  std::wstring hex = ColorToHex(s->color);
  RECT tr{panel.left + 64, panel.top + 12, panel.right - 10, panel.bottom - 10};
  DrawTextW(dc, hex.c_str(), -1, &tr, DT_LEFT | DT_TOP | DT_SINGLELINE);
  tr.top += 22;
  DrawTextW(dc, L"�������  Escȡ��", -1, &tr, DT_LEFT | DT_TOP | DT_SINGLELINE);
}

static LRESULT CALLBACK ColorPickerWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
  auto* s = reinterpret_cast<ColorPickerState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));
  switch (msg) {
    case WM_CREATE: {
      auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
      s = reinterpret_cast<ColorPickerState*>(cs->lpCreateParams);
      s->hwnd = hwnd;
      SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(s));
      if (!CaptureVirtualScreen(s)) return -1;
      GetCursorPos(&s->cursor);
      UpdateColorUnderCursor(s);
      SetCapture(hwnd);
      SetCursor(LoadCursorW(nullptr, IDC_CROSS));
      return 0;
    }
    case WM_MOUSEMOVE: {
      if (!s) return 0;
      POINT pt{GET_X_LPARAM(lParam) + s->vx, GET_Y_LPARAM(lParam) + s->vy};
      s->cursor = pt;
      UpdateColorUnderCursor(s);
      InvalidateRect(hwnd, nullptr, FALSE);
      return 0;
    }
    case WM_LBUTTONDOWN:
      if (s) {
        s->accepted = true;
        DestroyWindow(hwnd);
      }
      return 0;
    case WM_RBUTTONDOWN:
    case WM_KEYDOWN:
      if (wParam == VK_ESCAPE || msg == WM_RBUTTONDOWN) {
        if (s) s->accepted = false;
        DestroyWindow(hwnd);
      }
      return 0;
    case WM_PAINT: {
      PAINTSTRUCT ps{};
      HDC dc = BeginPaint(hwnd, &ps);
      if (s) DrawOverlay(s, dc);
      EndPaint(hwnd, &ps);
      return 0;
    }
    case WM_ERASEBKGND:
      return 1;
    case WM_DESTROY:
      if (GetCapture() == hwnd) ReleaseCapture();
      if (s) {
        if (s->memdc) DeleteDC(s->memdc);
        if (s->bitmap) DeleteObject(s->bitmap);
        s->memdc = nullptr;
        s->bitmap = nullptr;
      }
      return 0;
  }
  return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static ATOM EnsureColorPickerClass(HINSTANCE inst) {
  static ATOM atom = 0;
  if (atom) return atom;
  WNDCLASSW wc{};
  wc.lpfnWndProc = ColorPickerWndProc;
  wc.hInstance = inst;
  wc.hCursor = LoadCursorW(nullptr, IDC_CROSS);
  wc.lpszClassName = kColorPickerClass;
  wc.hbrBackground = (HBRUSH)GetStockObject(BLACK_BRUSH);
  atom = RegisterClassW(&wc);
  return atom;
}

}  // namespace

bool PickScreenColor(HWND owner, COLORREF* out_color) {
  if (out_color) *out_color = RGB(0, 0, 0);

  HINSTANCE inst = (HINSTANCE)GetModuleHandleW(nullptr);
  if (!EnsureColorPickerClass(inst)) return false;

  ColorPickerState state{};
  state.owner = owner;
  state.vx = GetSystemMetrics(SM_XVIRTUALSCREEN);
  state.vy = GetSystemMetrics(SM_YVIRTUALSCREEN);
  state.vw = GetSystemMetrics(SM_CXVIRTUALSCREEN);
  state.vh = GetSystemMetrics(SM_CYVIRTUALSCREEN);

  HWND hwnd = CreateWindowExW(WS_EX_TOPMOST | WS_EX_TOOLWINDOW, kColorPickerClass, L"",
                              WS_POPUP | WS_VISIBLE,
                              state.vx, state.vy, state.vw, state.vh,
                              owner, nullptr, inst, &state);
  if (!hwnd) return false;

  if (owner) EnableWindow(owner, FALSE);
  while (IsWindow(hwnd)) {
    MSG msg{};
    while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
      if (msg.message == WM_QUIT) {
        PostQuitMessage((int)msg.wParam);
        if (owner) EnableWindow(owner, TRUE);
        return false;
      }
      TranslateMessage(&msg);
      DispatchMessageW(&msg);
    }
    WaitMessage();
  }
  if (owner) {
    EnableWindow(owner, TRUE);
    SetForegroundWindow(owner);
  }

  if (state.accepted) {
    if (out_color) *out_color = state.color;
    return true;
  }
  return false;
}

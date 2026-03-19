#include "screenshot_tool.h"

#include "storage.h"
#include "win32_util.h"

#include <commdlg.h>
#include <windowsx.h>
#include <algorithm>
#include <cstdlib>
#include <string>
#include <vector>

namespace {

struct Op {
  enum class Type { Pen, Rect };
  Type type{Type::Pen};
  COLORREF color{RGB(231, 76, 60)};  // red
  int width{3};
  std::vector<POINT> pts;  // pen: polyline points (crop-local). rect: [p0, p1]
};

static RECT NormalizeRect(POINT a, POINT b) {
  RECT r{};
  r.left = (a.x < b.x) ? a.x : b.x;
  r.right = (a.x < b.x) ? b.x : a.x;
  r.top = (a.y < b.y) ? a.y : b.y;
  r.bottom = (a.y < b.y) ? b.y : a.y;
  return r;
}

static int RectW(const RECT& r) { return r.right - r.left; }
static int RectH(const RECT& r) { return r.bottom - r.top; }

static std::wstring GetNowStamp() {
  SYSTEMTIME st{};
  GetLocalTime(&st);
  wchar_t buf[64]{};
  wsprintfW(buf, L"%04d%02d%02d_%02d%02d%02d", (int)st.wYear, (int)st.wMonth, (int)st.wDay, (int)st.wHour, (int)st.wMinute,
            (int)st.wSecond);
  return buf;
}

static bool WriteBmpFile(const std::wstring& path, HBITMAP bmp, std::wstring* err) {
  if (!bmp) {
    if (err) *err = L"无效位图";
    return false;
  }

  BITMAP bm{};
  if (GetObjectW(bmp, sizeof(bm), &bm) == 0) {
    if (err) *err = L"GetObjectW 失败";
    return false;
  }

  // Force 32bpp top-down DIB.
  BITMAPINFO bmi{};
  bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
  bmi.bmiHeader.biWidth = bm.bmWidth;
  bmi.bmiHeader.biHeight = -bm.bmHeight;
  bmi.bmiHeader.biPlanes = 1;
  bmi.bmiHeader.biBitCount = 32;
  bmi.bmiHeader.biCompression = BI_RGB;

  std::vector<unsigned char> pixels;
  pixels.resize((size_t)bm.bmWidth * (size_t)bm.bmHeight * 4u);

  HDC dc = GetDC(nullptr);
  int got = GetDIBits(dc, bmp, 0, (UINT)bm.bmHeight, pixels.data(), &bmi, DIB_RGB_COLORS);
  ReleaseDC(nullptr, dc);
  if (got == 0) {
    if (err) *err = L"GetDIBits 失败";
    return false;
  }

  BITMAPFILEHEADER bfh{};
  bfh.bfType = 0x4D42;  // 'BM'
  bfh.bfOffBits = sizeof(BITMAPFILEHEADER) + sizeof(BITMAPINFOHEADER);
  bfh.bfSize = (DWORD)(bfh.bfOffBits + pixels.size());

  HANDLE h = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
  if (h == INVALID_HANDLE_VALUE) {
    if (err) *err = L"无法写入文件";
    return false;
  }

  DWORD written = 0;
  BOOL ok = TRUE;
  ok = ok && WriteFile(h, &bfh, sizeof(bfh), &written, nullptr);
  ok = ok && WriteFile(h, &bmi.bmiHeader, sizeof(bmi.bmiHeader), &written, nullptr);
  ok = ok && WriteFile(h, pixels.data(), (DWORD)pixels.size(), &written, nullptr);
  CloseHandle(h);

  if (!ok) {
    if (err) *err = L"写入BMP失败";
    return false;
  }
  return true;
}

static HBITMAP CreateBitmapCopy(HBITMAP src) {
  if (!src) return nullptr;
  BITMAP bm{};
  if (GetObjectW(src, sizeof(bm), &bm) == 0) return nullptr;

  HDC sdc = CreateCompatibleDC(nullptr);
  HDC ddc = CreateCompatibleDC(nullptr);
  if (!sdc || !ddc) {
    if (sdc) DeleteDC(sdc);
    if (ddc) DeleteDC(ddc);
    return nullptr;
  }

  HDC screen = GetDC(nullptr);
  HBITMAP dst = screen ? CreateCompatibleBitmap(screen, bm.bmWidth, bm.bmHeight) : nullptr;
  if (screen) ReleaseDC(nullptr, screen);
  if (!dst) {
    DeleteDC(sdc);
    DeleteDC(ddc);
    return nullptr;
  }

  HGDIOBJ os = SelectObject(sdc, src);
  HGDIOBJ od = SelectObject(ddc, dst);
  BitBlt(ddc, 0, 0, bm.bmWidth, bm.bmHeight, sdc, 0, 0, SRCCOPY);
  SelectObject(sdc, os);
  SelectObject(ddc, od);
  DeleteDC(sdc);
  DeleteDC(ddc);
  return dst;
}

static bool CopyBitmapToClipboard(HWND hwnd, HBITMAP bmp, std::wstring* err) {
  if (!bmp) {
    if (err) *err = L"无效位图";
    return false;
  }
  HBITMAP copy = CreateBitmapCopy(bmp);
  if (!copy) {
    if (err) *err = L"复制位图失败";
    return false;
  }

  if (!OpenClipboard(hwnd)) {
    DeleteObject(copy);
    if (err) *err = L"无法打开剪贴板";
    return false;
  }
  EmptyClipboard();
  if (!SetClipboardData(CF_BITMAP, copy)) {
    CloseClipboard();
    DeleteObject(copy);
    if (err) *err = L"写入剪贴板失败";
    return false;
  }
  CloseClipboard();
  // Clipboard now owns `copy`.
  return true;
}

static void DrawHintText(HDC dc, const RECT& rc, const wchar_t* text) {
  SetBkMode(dc, TRANSPARENT);
  SetTextColor(dc, RGB(20, 20, 20));

  RECT bg = rc;
  bg.right = bg.left + 520;
  bg.bottom = bg.top + 46;
  HBRUSH br = CreateSolidBrush(RGB(245, 246, 248));
  HPEN pen = CreatePen(PS_SOLID, 1, RGB(210, 214, 220));
  HGDIOBJ ob = SelectObject(dc, br);
  HGDIOBJ op = SelectObject(dc, pen);
  RoundRect(dc, bg.left, bg.top, bg.right, bg.bottom, 10, 10);
  SelectObject(dc, op);
  SelectObject(dc, ob);
  DeleteObject(pen);
  DeleteObject(br);

  RECT t = bg;
  t.left += 10;
  t.top += 8;
  DrawTextW(dc, text, -1, &t, DT_LEFT | DT_TOP | DT_NOPREFIX);
}

struct ScreenshotState {
  HWND hwnd{};
  HWND owner{};

  int vx{};
  int vy{};
  int vw{};
  int vh{};

  HBITMAP full_bmp{};

  enum class Mode { Selecting, Editing } mode{Mode::Selecting};

  bool dragging{false};
  POINT down{};
  POINT cur{};
  RECT sel{};  // client coords
  bool has_sel{false};

  HBITMAP crop_base{};
  HBITMAP crop_work{};
  std::vector<Op> ops;

  Op::Type tool{Op::Type::Pen};
  bool drawing{false};
  Op current_op{};
};

static void FreeBitmap(HBITMAP* b) {
  if (b && *b) {
    DeleteObject(*b);
    *b = nullptr;
  }
}

static void ResetCrop(ScreenshotState* s) {
  FreeBitmap(&s->crop_base);
  FreeBitmap(&s->crop_work);
  s->ops.clear();
  s->drawing = false;
}

static void RebuildCropWork(ScreenshotState* s) {
  if (!s || !s->crop_base || !s->crop_work) return;

  BITMAP bm{};
  if (GetObjectW(s->crop_base, sizeof(bm), &bm) == 0) return;

  HDC sdc = CreateCompatibleDC(nullptr);
  HDC ddc = CreateCompatibleDC(nullptr);
  if (!sdc || !ddc) {
    if (sdc) DeleteDC(sdc);
    if (ddc) DeleteDC(ddc);
    return;
  }
  HGDIOBJ ob = SelectObject(sdc, s->crop_base);
  HGDIOBJ ow = SelectObject(ddc, s->crop_work);

  BitBlt(ddc, 0, 0, bm.bmWidth, bm.bmHeight, sdc, 0, 0, SRCCOPY);

  for (const auto& op : s->ops) {
    HPEN pen = CreatePen(PS_SOLID, op.width, op.color);
    HGDIOBJ opn = SelectObject(ddc, pen);
    HGDIOBJ obr = SelectObject(ddc, GetStockObject(HOLLOW_BRUSH));
    SetBkMode(ddc, TRANSPARENT);

    if (op.type == Op::Type::Pen) {
      if (op.pts.size() >= 2) {
        Polyline(ddc, op.pts.data(), (int)op.pts.size());
      } else if (op.pts.size() == 1) {
        MoveToEx(ddc, op.pts[0].x, op.pts[0].y, nullptr);
        LineTo(ddc, op.pts[0].x + 1, op.pts[0].y + 1);
      }
    } else if (op.type == Op::Type::Rect) {
      if (op.pts.size() >= 2) {
        RECT r = NormalizeRect(op.pts[0], op.pts[1]);
        Rectangle(ddc, r.left, r.top, r.right, r.bottom);
      }
    }

    SelectObject(ddc, obr);
    SelectObject(ddc, opn);
    DeleteObject(pen);
  }

  SelectObject(ddc, ow);
  SelectObject(sdc, ob);
  DeleteDC(ddc);
  DeleteDC(sdc);
}

static bool CaptureFullVirtualScreen(ScreenshotState* s, std::wstring* err) {
  if (!s) return false;

  s->vx = GetSystemMetrics(SM_XVIRTUALSCREEN);
  s->vy = GetSystemMetrics(SM_YVIRTUALSCREEN);
  s->vw = GetSystemMetrics(SM_CXVIRTUALSCREEN);
  s->vh = GetSystemMetrics(SM_CYVIRTUALSCREEN);
  if (s->vw <= 0 || s->vh <= 0) {
    if (err) *err = L"无法获取屏幕尺寸";
    return false;
  }

  HDC screen = GetDC(nullptr);
  if (!screen) {
    if (err) *err = L"GetDC(NULL) 失败";
    return false;
  }
  HDC mem = CreateCompatibleDC(screen);
  if (!mem) {
    ReleaseDC(nullptr, screen);
    if (err) *err = L"CreateCompatibleDC 失败";
    return false;
  }
  HBITMAP bmp = CreateCompatibleBitmap(screen, s->vw, s->vh);
  if (!bmp) {
    DeleteDC(mem);
    ReleaseDC(nullptr, screen);
    if (err) *err = L"CreateCompatibleBitmap 失败";
    return false;
  }

  HGDIOBJ ob = SelectObject(mem, bmp);
  BOOL ok = BitBlt(mem, 0, 0, s->vw, s->vh, screen, s->vx, s->vy, SRCCOPY | CAPTUREBLT);
  SelectObject(mem, ob);
  DeleteDC(mem);
  ReleaseDC(nullptr, screen);

  if (!ok) {
    DeleteObject(bmp);
    if (err) *err = L"BitBlt 截图失败";
    return false;
  }

  s->full_bmp = bmp;
  return true;
}

static bool BuildCropBitmaps(ScreenshotState* s, std::wstring* err) {
  if (!s || !s->full_bmp || !s->has_sel) return false;
  int w = RectW(s->sel);
  int h = RectH(s->sel);
  if (w <= 0 || h <= 0) return false;

  ResetCrop(s);

  HDC srcdc = CreateCompatibleDC(nullptr);
  HDC dstdc = CreateCompatibleDC(nullptr);
  if (!srcdc || !dstdc) {
    if (srcdc) DeleteDC(srcdc);
    if (dstdc) DeleteDC(dstdc);
    if (err) *err = L"创建DC失败";
    return false;
  }

  // Create base/work as compatible bitmaps.
  HDC screen = GetDC(nullptr);
  HBITMAP base = CreateCompatibleBitmap(screen, w, h);
  HBITMAP work = CreateCompatibleBitmap(screen, w, h);
  ReleaseDC(nullptr, screen);
  if (!base || !work) {
    if (base) DeleteObject(base);
    if (work) DeleteObject(work);
    DeleteDC(srcdc);
    DeleteDC(dstdc);
    if (err) *err = L"创建位图失败";
    return false;
  }

  HGDIOBJ os = SelectObject(srcdc, s->full_bmp);
  HGDIOBJ ob = SelectObject(dstdc, base);
  BitBlt(dstdc, 0, 0, w, h, srcdc, s->sel.left, s->sel.top, SRCCOPY);

  SelectObject(dstdc, work);
  BitBlt(dstdc, 0, 0, w, h, srcdc, s->sel.left, s->sel.top, SRCCOPY);

  SelectObject(srcdc, os);
  SelectObject(dstdc, ob);
  DeleteDC(srcdc);
  DeleteDC(dstdc);

  s->crop_base = base;
  s->crop_work = work;
  return true;
}

static void FinishAndCloseCopy(HWND hwnd, ScreenshotState* s) {
  if (!s || !s->crop_work) {
    DestroyWindow(hwnd);
    return;
  }
  std::wstring err;
  if (!CopyBitmapToClipboard(hwnd, s->crop_work, &err)) {
    MessageBoxW(hwnd, err.c_str(), L"复制截图失败", MB_OK | MB_ICONERROR);
    return;
  }
  DestroyWindow(hwnd);
}

static void SaveAsBmpDialog(HWND hwnd, ScreenshotState* s) {
  if (!s || !s->crop_work) return;

  std::wstring data = GetDataRootDir();
  std::wstring dir = JoinPath(data, L"screenshots");
  EnsureDirExists(data);
  EnsureDirExists(dir);

  std::wstring def = JoinPath(dir, L"screenshot_" + GetNowStamp() + L".bmp");

  wchar_t filebuf[MAX_PATH]{};
  wcsncpy_s(filebuf, def.c_str(), _TRUNCATE);

  OPENFILENAMEW ofn{};
  ofn.lStructSize = sizeof(ofn);
  ofn.hwndOwner = hwnd;
  ofn.lpstrFilter = L"BMP 图片 (*.bmp)\0*.bmp\0所有文件 (*.*)\0*.*\0";
  ofn.lpstrFile = filebuf;
  ofn.nMaxFile = MAX_PATH;
  ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;
  ofn.lpstrDefExt = L"bmp";
  ofn.lpstrTitle = L"保存截图为 BMP";
  if (!GetSaveFileNameW(&ofn)) return;

  std::wstring err;
  if (!WriteBmpFile(filebuf, s->crop_work, &err)) {
    MessageBoxW(hwnd, err.c_str(), L"保存失败", MB_OK | MB_ICONERROR);
    return;
  }
}

static void ShowContextMenu(HWND hwnd, ScreenshotState* s, int x, int y) {
  HMENU m = CreatePopupMenu();
  if (!m) return;
  AppendMenuW(m, MF_STRING, 1, L"复制到剪贴板 (C/Enter)");
  AppendMenuW(m, MF_STRING, 2, L"保存为 BMP... (S)");
  AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(m, MF_STRING, 3, L"工具: 画笔 (P)");
  AppendMenuW(m, MF_STRING, 4, L"工具: 矩形 (R)");
  AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(m, MF_STRING, 5, L"撤销 (Ctrl+Z)");
  AppendMenuW(m, MF_STRING, 6, L"清空标记");
  AppendMenuW(m, MF_SEPARATOR, 0, nullptr);
  AppendMenuW(m, MF_STRING, 7, L"取消 (Esc)");

  UINT cmd = TrackPopupMenu(m, TPM_RETURNCMD | TPM_RIGHTBUTTON, x, y, 0, hwnd, nullptr);
  DestroyMenu(m);
  if (!cmd) return;

  switch (cmd) {
    case 1:
      FinishAndCloseCopy(hwnd, s);
      break;
    case 2:
      SaveAsBmpDialog(hwnd, s);
      break;
    case 3:
      s->tool = Op::Type::Pen;
      InvalidateRect(hwnd, nullptr, FALSE);
      break;
    case 4:
      s->tool = Op::Type::Rect;
      InvalidateRect(hwnd, nullptr, FALSE);
      break;
    case 5:
      if (!s->ops.empty()) {
        s->ops.pop_back();
        RebuildCropWork(s);
        InvalidateRect(hwnd, nullptr, FALSE);
      }
      break;
    case 6:
      s->ops.clear();
      RebuildCropWork(s);
      InvalidateRect(hwnd, nullptr, FALSE);
      break;
    case 7:
      DestroyWindow(hwnd);
      break;
    default:
      break;
  }
}

static LRESULT CALLBACK ScreenshotWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
  auto* s = reinterpret_cast<ScreenshotState*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

  switch (msg) {
    case WM_CREATE: {
      auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
      s = reinterpret_cast<ScreenshotState*>(cs->lpCreateParams);
      s->hwnd = hwnd;
      SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(s));
      return 0;
    }

    case WM_ERASEBKGND:
      return 1;

    case WM_SETCURSOR:
      SetCursor(LoadCursorW(nullptr, IDC_CROSS));
      return TRUE;

    case WM_KEYDOWN: {
      if (!s) break;
      bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
      if (wParam == VK_ESCAPE) {
        DestroyWindow(hwnd);
        return 0;
      }
      if (wParam == VK_RETURN) {
        if (s->mode == ScreenshotState::Mode::Editing) FinishAndCloseCopy(hwnd, s);
        return 0;
      }
      if (wParam == 'C') {
        if (s->mode == ScreenshotState::Mode::Editing) FinishAndCloseCopy(hwnd, s);
        return 0;
      }
      if (wParam == 'S') {
        if (s->mode == ScreenshotState::Mode::Editing) SaveAsBmpDialog(hwnd, s);
        return 0;
      }
      if (wParam == 'P') {
        s->tool = Op::Type::Pen;
        InvalidateRect(hwnd, nullptr, FALSE);
        return 0;
      }
      if (wParam == 'R') {
        s->tool = Op::Type::Rect;
        InvalidateRect(hwnd, nullptr, FALSE);
        return 0;
      }
      if (ctrl && (wParam == 'Z')) {
        if (!s->ops.empty()) {
          s->ops.pop_back();
          RebuildCropWork(s);
          InvalidateRect(hwnd, nullptr, FALSE);
        }
        return 0;
      }
      break;
    }

    case WM_LBUTTONDOWN: {
      if (!s) break;
      SetCapture(hwnd);
      s->dragging = true;
      s->down = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
      s->cur = s->down;

      if (s->mode == ScreenshotState::Mode::Selecting) {
        s->has_sel = false;
      } else {
        // Begin drawing.
        s->drawing = true;
        s->current_op = Op{};
        s->current_op.type = s->tool;
        s->current_op.pts.clear();
        POINT p{GET_X_LPARAM(lParam) - s->sel.left, GET_Y_LPARAM(lParam) - s->sel.top};
        p.x = (LONG)std::clamp<int>((int)p.x, 0, RectW(s->sel));
        p.y = (LONG)std::clamp<int>((int)p.y, 0, RectH(s->sel));
        s->current_op.pts.push_back(p);
      }

      InvalidateRect(hwnd, nullptr, FALSE);
      return 0;
    }

    case WM_MOUSEMOVE: {
      if (!s) break;
      if (!s->dragging && !s->drawing) break;
      POINT p{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
      s->cur = p;

      if (s->mode == ScreenshotState::Mode::Editing && s->drawing) {
        POINT lp{p.x - s->sel.left, p.y - s->sel.top};
        lp.x = (LONG)std::clamp<int>((int)lp.x, 0, RectW(s->sel));
        lp.y = (LONG)std::clamp<int>((int)lp.y, 0, RectH(s->sel));

        if (s->current_op.type == Op::Type::Pen) {
          // Avoid huge point spam if mouse is stuck.
          if (s->current_op.pts.empty() ||
              (std::abs(s->current_op.pts.back().x - lp.x) + std::abs(s->current_op.pts.back().y - lp.y) >= 1)) {
            s->current_op.pts.push_back(lp);
          }
        } else if (s->current_op.type == Op::Type::Rect) {
          if (s->current_op.pts.size() == 1) {
            s->current_op.pts.push_back(lp);
          } else if (s->current_op.pts.size() >= 2) {
            s->current_op.pts[1] = lp;
          }
        }
      }

      InvalidateRect(hwnd, nullptr, FALSE);
      return 0;
    }

    case WM_LBUTTONUP: {
      if (!s) break;
      ReleaseCapture();
      s->dragging = false;
      POINT up{GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
      s->cur = up;

      if (s->mode == ScreenshotState::Mode::Selecting) {
        RECT r = NormalizeRect(s->down, up);
        // Clamp to client.
        r.left = (LONG)std::clamp<int>((int)r.left, 0, s->vw);
        r.right = (LONG)std::clamp<int>((int)r.right, 0, s->vw);
        r.top = (LONG)std::clamp<int>((int)r.top, 0, s->vh);
        r.bottom = (LONG)std::clamp<int>((int)r.bottom, 0, s->vh);

        if (RectW(r) < 10 || RectH(r) < 10) {
          s->has_sel = false;
          InvalidateRect(hwnd, nullptr, FALSE);
          return 0;
        }

        s->sel = r;
        s->has_sel = true;
        s->mode = ScreenshotState::Mode::Editing;
        std::wstring err;
        if (!BuildCropBitmaps(s, &err)) {
          MessageBoxW(hwnd, err.c_str(), L"截图失败", MB_OK | MB_ICONERROR);
          DestroyWindow(hwnd);
          return 0;
        }
      } else if (s->mode == ScreenshotState::Mode::Editing) {
        if (s->drawing) {
          s->drawing = false;
          // Commit op if meaningful.
          bool ok = false;
          if (s->current_op.type == Op::Type::Pen) ok = (s->current_op.pts.size() >= 1);
          if (s->current_op.type == Op::Type::Rect) ok = (s->current_op.pts.size() >= 2);
          if (ok) {
            s->ops.push_back(std::move(s->current_op));
            RebuildCropWork(s);
          }
        }
      }

      InvalidateRect(hwnd, nullptr, FALSE);
      return 0;
    }

    case WM_RBUTTONUP: {
      if (!s) break;
      POINT pt{};
      GetCursorPos(&pt);
      ShowContextMenu(hwnd, s, pt.x, pt.y);
      return 0;
    }

    case WM_PAINT: {
      if (!s) break;
      PAINTSTRUCT ps{};
      HDC dc = BeginPaint(hwnd, &ps);

      // Draw captured screen background.
      HDC mem = CreateCompatibleDC(dc);
      HGDIOBJ ob = SelectObject(mem, s->full_bmp);
      BitBlt(dc, 0, 0, s->vw, s->vh, mem, 0, 0, SRCCOPY);
      SelectObject(mem, ob);
      DeleteDC(mem);

      // Draw selection / crop preview.
      if (s->mode == ScreenshotState::Mode::Selecting) {
        if (s->dragging) {
          RECT r = NormalizeRect(s->down, s->cur);
          HPEN pen = CreatePen(PS_SOLID, 2, RGB(0, 120, 215));
          HGDIOBJ op = SelectObject(dc, pen);
          HGDIOBJ obr = SelectObject(dc, GetStockObject(HOLLOW_BRUSH));
          Rectangle(dc, r.left, r.top, r.right, r.bottom);
          SelectObject(dc, obr);
          SelectObject(dc, op);
          DeleteObject(pen);
        }
        RECT hint{};
        hint.left = 12;
        hint.top = 12;
        hint.right = s->vw - 12;
        hint.bottom = s->vh - 12;
        DrawHintText(dc, hint, L"框选截图区域: 鼠标拖拽。Esc 取消");
      } else if (s->mode == ScreenshotState::Mode::Editing && s->has_sel) {
        // Draw crop work over the original region.
        if (s->crop_work) {
          HDC cdc = CreateCompatibleDC(dc);
          HGDIOBJ oc = SelectObject(cdc, s->crop_work);
          BitBlt(dc, s->sel.left, s->sel.top, RectW(s->sel), RectH(s->sel), cdc, 0, 0, SRCCOPY);
          SelectObject(cdc, oc);
          DeleteDC(cdc);
        }

        // Outline selection.
        HPEN pen = CreatePen(PS_SOLID, 2, RGB(0, 120, 215));
        HGDIOBJ op = SelectObject(dc, pen);
        HGDIOBJ obr = SelectObject(dc, GetStockObject(HOLLOW_BRUSH));
        Rectangle(dc, s->sel.left, s->sel.top, s->sel.right, s->sel.bottom);
        SelectObject(dc, obr);
        SelectObject(dc, op);
        DeleteObject(pen);

        RECT hint{};
        hint.left = s->sel.left + 10;
        hint.top = s->sel.top + 10;
        hint.right = s->vw - 12;
        hint.bottom = s->vh - 12;

        const wchar_t* tool = (s->tool == Op::Type::Pen) ? L"画笔" : L"矩形";
        wchar_t text[256]{};
        wsprintfW(text, L"标记模式(%s): 左键绘制, 右键菜单。P画笔 R矩形 Ctrl+Z撤销 C/Enter复制 S保存 Esc取消", tool);
        DrawHintText(dc, hint, text);
      }

      EndPaint(hwnd, &ps);
      return 0;
    }

    case WM_DESTROY: {
      if (s) {
        FreeBitmap(&s->full_bmp);
        ResetCrop(s);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
      }
      return 0;
    }
  }

  return DefWindowProcW(hwnd, msg, wParam, lParam);
}

static ATOM EnsureScreenshotClass() {
  static ATOM a = 0;
  if (a) return a;

  WNDCLASSW wc{};
  wc.lpfnWndProc = ScreenshotWndProc;
  wc.hInstance = GetModuleHandleW(nullptr);
  wc.lpszClassName = L"WorkLogLite_ScreenshotOverlay";
  wc.hCursor = LoadCursorW(nullptr, IDC_CROSS);
  wc.hbrBackground = (HBRUSH)GetStockObject(BLACK_BRUSH);
  a = RegisterClassW(&wc);
  return a;
}

}  // namespace

void StartScreenshotTool(HWND owner) {
  ScreenshotState st{};
  st.owner = owner;

  std::wstring err;
  if (!CaptureFullVirtualScreen(&st, &err)) {
    MessageBoxW(owner, err.c_str(), L"截图失败", MB_OK | MB_ICONERROR);
    return;
  }

  if (!EnsureScreenshotClass()) {
    MessageBoxW(owner, L"RegisterClass 失败", L"截图失败", MB_OK | MB_ICONERROR);
    DeleteObject(st.full_bmp);
    return;
  }

  HWND hwnd = CreateWindowExW(WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
                              L"WorkLogLite_ScreenshotOverlay",
                              L"",
                              WS_POPUP,
                              st.vx,
                              st.vy,
                              st.vw,
                              st.vh,
                              owner,
                              nullptr,
                              GetModuleHandleW(nullptr),
                              &st);
  if (!hwnd) {
    MessageBoxW(owner, L"无法创建截图窗口", L"截图失败", MB_OK | MB_ICONERROR);
    DeleteObject(st.full_bmp);
    return;
  }

  ShowWindow(hwnd, SW_SHOW);
  SetForegroundWindow(hwnd);
  SetFocus(hwnd);

  // Modal loop so this tool is self-contained and doesn't interfere with the main app state.
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

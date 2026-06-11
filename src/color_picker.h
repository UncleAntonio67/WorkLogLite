#pragma once

#include <windows.h>
#include <string>

bool PickScreenColor(HWND owner, COLORREF* out_color, std::wstring* err = nullptr);

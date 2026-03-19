#pragma once

#include <windows.h>

// Self-contained screenshot tool:
// - capture full virtual screen into memory (offline)
// - user drags to select a rectangle
// - user can mark (pen / rectangle)
// - copy result to clipboard, optionally save as BMP
//
// This intentionally avoids calling system screen snip tools so it keeps working in locked-down environments.
void StartScreenshotTool(HWND owner);


#pragma once

#include <windows.h>

#include <string>

// Appends a set of realistic "work day" demo data into current data directory.
// - Does NOT delete existing user data.
// - Adds demo categories/tasks/recurring meetings if missing.
// - Adds demo day entries in a small time window around `anchor` (usually "today").
bool GenerateDemoData(const SYSTEMTIME& anchor, std::wstring* err);


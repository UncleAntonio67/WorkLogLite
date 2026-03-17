#pragma once

#include "types.h"

#include <windows.h>

#include <string>
#include <vector>

struct DayData {
  SYSTEMTIME date{};
  std::vector<Entry> entries;
};

// Root: <exe_dir>\data
std::wstring GetDataRootDir();
std::wstring GetDayFilePath(const SYSTEMTIME& date);

bool LoadDayFile(const SYSTEMTIME& date, DayData* out, std::wstring* err);
bool SaveDayFile(const DayData& day, std::wstring* err);


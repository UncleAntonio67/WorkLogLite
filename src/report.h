#pragma once

#include "types.h"

#include <windows.h>

#include <string>

struct ReportRange {
  SYSTEMTIME start{};  // inclusive
  SYSTEMTIME end{};    // inclusive
};

ReportRange QuarterRangeForDate(const SYSTEMTIME& selected);
ReportRange YearRangeForDate(const SYSTEMTIME& selected);

// Generate a Markdown summary for range by loading existing day files.
// Includes item headers and bodies.
bool GenerateReportMarkdown(const ReportRange& range, std::wstring* out_md, std::wstring* err);

// Category-focused report for quarter/year summaries.
bool GenerateReportMarkdownByCategory(const ReportRange& range, std::wstring* out_md, std::wstring* err);

// Flat CSV export for further processing (Excel/Python/etc).
// One row per entry, includes task metadata when applicable.
bool GenerateReportCsvFlat(const ReportRange& range, std::wstring* out_csv, std::wstring* err);

#pragma once

#include "types.h"

#include <windows.h>

#include <string>
#include <vector>

struct Task {
  std::wstring id;
  std::wstring category;
  std::wstring title;

  // Optional: user-selected materials folder for this task (local path).
  // If empty, the app can fall back to the default auto materials folder.
  std::wstring materials_dir;

  // Optional metadata for "task has basis and description".
  std::wstring basis;        // e.g. meeting/date/doc reference
  std::wstring desc_plain;   // fallback text
  std::string desc_rtf_b64;  // rich text (RTF bytes, base64)

  SYSTEMTIME start{};  // inclusive
  SYSTEMTIME end{};    // inclusive
  EntryStatus status{EntryStatus::Todo};
};

std::wstring GetTasksFilePath();

bool LoadTasks(std::vector<Task>* out, std::wstring* err);
bool SaveTasks(const std::vector<Task>& tasks, std::wstring* err);

bool TaskActiveOn(const Task& t, const SYSTEMTIME& date);

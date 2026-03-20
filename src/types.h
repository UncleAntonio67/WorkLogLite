#pragma once

#include <string>

enum class EntryType {
  Note = 0,
  TaskProgress = 1,
  Meeting = 2,
};

enum class EntryStatus {
  None = 0,
  Todo = 1,
  Doing = 2,
  Blocked = 3,
  Done = 4,
};

struct Entry {
  std::wstring id;  // unique within a day file
  bool placeholder{false};  // injected UI-only entry; won't be persisted unless saved

  EntryType type{EntryType::Note};
  std::wstring task_id;  // only for TaskProgress

  // Optional: user-selected materials folder for this entry (local path).
  // If empty, the app can fall back to the default auto materials folder.
  std::wstring materials_dir;

  std::wstring category;    // user-defined
  std::wstring start_time;  // "HH:MM" or empty
  std::wstring end_time;    // "HH:MM" or empty
  std::wstring title;
  EntryStatus status{EntryStatus::None};

  // body_plain is used for reports/search and as a fallback if body_rtf_b64 is empty.
  std::wstring body_plain;
  std::string body_rtf_b64;  // base64-encoded RTF bytes (RichEdit)
};

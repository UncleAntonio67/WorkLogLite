#pragma once

#include <windows.h>

#include <string>
#include <vector>

struct RecurringMeeting {
  std::wstring id;
  std::wstring category;
  std::wstring title;

  // Weekly recurrence.
  // weekday: 0=Sun ... 6=Sat
  int weekday{1};
  int interval_weeks{1};  // 1=weekly, 2=bi-weekly, etc.

  SYSTEMTIME start{};  // inclusive
  SYSTEMTIME end{};    // inclusive

  std::wstring start_time;  // optional "HH:MM"
  std::wstring end_time;    // optional "HH:MM"
  std::wstring template_plain;
};

std::wstring GetRecurringMeetingsFilePath();

bool LoadRecurringMeetings(std::vector<RecurringMeeting>* out, std::wstring* err);
bool SaveRecurringMeetings(const std::vector<RecurringMeeting>& items, std::wstring* err);

bool RecurringMeetingOccursOn(const RecurringMeeting& m, const SYSTEMTIME& date);


#include "categories.h"

#include "storage.h"
#include "win32_util.h"

#include <cstdio>
#include <cwctype>
#include <sstream>
#include <utility>

static std::wstring Trim(const std::wstring& s) {
  size_t i = 0;
  while (i < s.size() && iswspace(s[i])) i++;
  size_t j = s.size();
  while (j > i && iswspace(s[j - 1])) j--;
  return s.substr(i, j - i);
}

std::wstring GetCategoriesFilePath() {
  // Store alongside data root so portable mode stays portable.
  return JoinPath(GetDataRootDir(), L"categories.txt");
}

bool CategoriesContains(const std::vector<std::wstring>& categories, const std::wstring& cat) {
  for (const auto& c : categories) {
    if (c == cat) return true;
  }
  return false;
}

void NormalizeCategories(std::vector<std::wstring>* categories) {
  if (!categories) return;
  std::vector<std::wstring> out;
  out.reserve(categories->size());
  for (const auto& raw : *categories) {
    std::wstring t = Trim(raw);
    if (t.empty()) continue;
    if (CategoriesContains(out, t)) continue;
    out.push_back(std::move(t));
  }
  *categories = std::move(out);
}

static bool ReadUtf8TextFile(const std::wstring& path, std::wstring* out, std::wstring* err) {
  if (!out) return false;
  FILE* f = nullptr;
  _wfopen_s(&f, path.c_str(), L"rb");
  if (!f) {
    if (err) *err = L"无法打开文件: " + path;
    return false;
  }
  fseek(f, 0, SEEK_END);
  long size = ftell(f);
  fseek(f, 0, SEEK_SET);
  if (size < 0) size = 0;
  std::string bytes;
  bytes.resize((size_t)size);
  if (size > 0) {
    size_t n = fread(bytes.data(), 1, (size_t)size, f);
    bytes.resize(n);
  }
  fclose(f);
  *out = Utf8ToWide(bytes);
  return true;
}

static bool WriteUtf8TextFileAtomic(const std::wstring& path, const std::wstring& content, std::wstring* err) {
  std::string bytes = WideToUtf8(content);
  std::wstring tmp = path + L".tmp";
  FILE* f = nullptr;
  _wfopen_s(&f, tmp.c_str(), L"wb");
  if (!f) {
    if (err) *err = L"无法写入文件: " + tmp;
    return false;
  }
  size_t n = fwrite(bytes.data(), 1, bytes.size(), f);
  fflush(f);
  fclose(f);
  if (n != bytes.size()) {
    if (err) *err = L"写入失败: " + tmp;
    return false;
  }
  if (!MoveFileExW(tmp.c_str(), path.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
    if (err) *err = L"写入失败: " + path;
    DeleteFileW(tmp.c_str());
    return false;
  }
  return true;
}

bool LoadCategories(std::vector<std::wstring>* out, std::wstring* err) {
  if (!out) return false;
  out->clear();

  // Ensure root exists.
  EnsureDirExists(GetDataRootDir());

  std::wstring path = GetCategoriesFilePath();
  if (!FileExists(path)) {
    // Defaults.
    out->push_back(L"工作");
    out->push_back(L"会议");
    out->push_back(L"目标");
    NormalizeCategories(out);
    // Best-effort save; don't fail load if write fails.
    std::wstring save_err;
    SaveCategories(*out, &save_err);
    return true;
  }

  std::wstring content;
  if (!ReadUtf8TextFile(path, &content, err)) return false;

  std::wstring line;
  std::wstringstream ss(content);
  while (std::getline(ss, line)) {
    if (!line.empty() && line.back() == L'\r') line.pop_back();
    out->push_back(line);
  }
  NormalizeCategories(out);

  // If file existed but was empty, restore defaults.
  if (out->empty()) {
    out->push_back(L"工作");
    out->push_back(L"会议");
    out->push_back(L"目标");
    NormalizeCategories(out);
    std::wstring save_err;
    SaveCategories(*out, &save_err);
  }
  return true;
}

bool SaveCategories(const std::vector<std::wstring>& categories, std::wstring* err) {
  std::vector<std::wstring> cats = categories;
  NormalizeCategories(&cats);

  std::wstring content;
  for (const auto& c : cats) {
    content += c;
    content += L"\n";
  }
  return WriteUtf8TextFileAtomic(GetCategoriesFilePath(), content, err);
}

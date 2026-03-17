#pragma once

#include <string>
#include <vector>

// Categories are stored as UTF-8 text: one category per line.
std::wstring GetCategoriesFilePath();

bool LoadCategories(std::vector<std::wstring>* out, std::wstring* err);
bool SaveCategories(const std::vector<std::wstring>& categories, std::wstring* err);

// Trim, remove empties, de-duplicate (case-sensitive), keep original order.
void NormalizeCategories(std::vector<std::wstring>* categories);

bool CategoriesContains(const std::vector<std::wstring>& categories, const std::wstring& cat);


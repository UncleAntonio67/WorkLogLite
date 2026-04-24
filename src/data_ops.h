#pragma once

#include <string>

bool ResolveWorkLogLiteImportRoot(const std::wstring& picked_folder, std::wstring* out_root);
bool ExportWorkLogLiteData(const std::wstring& src_root,
                           const std::wstring& picked_folder,
                           std::wstring* out_export_dir,
                           std::wstring* err);
bool ImportWorkLogLiteData(const std::wstring& picked_folder,
                           const std::wstring& dst_root,
                           std::wstring* out_backup_dir,
                           std::wstring* err);
bool ClearWorkLogLiteData(const std::wstring& root,
                          std::wstring* out_backup_dir,
                          std::wstring* err);

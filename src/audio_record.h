#pragma once

#include <windows.h>

#include <string>

enum class AudioRecordSource {
  SystemLoopback,
  MicrophoneCapture,
};

struct AudioRecorder;

std::wstring GetAudioRecordingsDir(AudioRecordSource source);

bool AudioRecorderIsRunning(const AudioRecorder* recorder);

bool StartAudioRecorder(HWND owner,
                        AudioRecordSource source,
                        AudioRecorder** io_recorder,
                        std::wstring* out_path,
                        std::wstring* err);

bool StopAudioRecorder(AudioRecorder** io_recorder,
                       std::wstring* out_path,
                       std::wstring* err);

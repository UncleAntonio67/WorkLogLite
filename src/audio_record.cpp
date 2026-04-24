#include "audio_record.h"

#include "storage.h"
#include "win32_util.h"

#include <initguid.h>
#include <audioclient.h>
#include <mmdeviceapi.h>

#include <cmath>
#include <cstdio>
#include <string>
#include <vector>

struct AudioRecorder {
  HANDLE thread{};
  HANDLE stop_event{};
  HANDLE ready_event{};
  bool running{false};
  bool init_ok{false};
  AudioRecordSource source{AudioRecordSource::SystemLoopback};
  ULONGLONG start_tick{};
  ULONGLONG written_frames{};
  std::wstring path;
  std::wstring error;
};

namespace {

const GUID kSubtypePcm =
    {0x00000001, 0x0000, 0x0010, {0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71}};
const GUID kSubtypeIeeeFloat =
    {0x00000003, 0x0000, 0x0010, {0x80, 0x00, 0x00, 0xaa, 0x00, 0x38, 0x9b, 0x71}};

template <typename T>
void SafeRelease(T** p) {
  if (p && *p) {
    (*p)->Release();
    *p = nullptr;
  }
}

static const wchar_t* SourceStem(AudioRecordSource source) {
  return source == AudioRecordSource::MicrophoneCapture ? L"microphone-audio-" : L"system-audio-";
}

static const wchar_t* SourceFolder(AudioRecordSource source) {
  return source == AudioRecordSource::MicrophoneCapture ? L"microphone_audio" : L"system_audio";
}

static const wchar_t* SourceDisplayName(AudioRecordSource source) {
  return source == AudioRecordSource::MicrophoneCapture ? L"microphone" : L"system audio";
}

static EDataFlow SourceDataFlow(AudioRecordSource source) {
  return source == AudioRecordSource::MicrophoneCapture ? eCapture : eRender;
}

static DWORD SourceStreamFlags(AudioRecordSource source) {
  return source == AudioRecordSource::MicrophoneCapture ? 0 : AUDCLNT_STREAMFLAGS_LOOPBACK;
}

static std::wstring GetNowStamp() {
  SYSTEMTIME st{};
  GetLocalTime(&st);
  wchar_t buf[64]{};
  wsprintfW(buf, L"%04d%02d%02d_%02d%02d%02d",
            (int)st.wYear, (int)st.wMonth, (int)st.wDay,
            (int)st.wHour, (int)st.wMinute, (int)st.wSecond);
  return buf;
}

static bool WriteAll(FILE* f, const void* data, size_t size) {
  return f && size == fwrite(data, 1, size, f);
}

static bool RewriteCanonicalPcmHeader(FILE* f, const WAVEFORMATEX* fmt, DWORD data_size) {
  if (!f || !fmt) return false;
  if (fmt->wFormatTag != WAVE_FORMAT_PCM) return false;
  const DWORD riff_size = 36u + data_size;
  const DWORD fmt_size = 16u;
  if (!WriteAll(f, "RIFF", 4)) return false;
  if (!WriteAll(f, &riff_size, sizeof(riff_size))) return false;
  if (!WriteAll(f, "WAVE", 4)) return false;
  if (!WriteAll(f, "fmt ", 4)) return false;
  if (!WriteAll(f, &fmt_size, sizeof(fmt_size))) return false;
  if (!WriteAll(f, &fmt->wFormatTag, sizeof(fmt->wFormatTag))) return false;
  if (!WriteAll(f, &fmt->nChannels, sizeof(fmt->nChannels))) return false;
  if (!WriteAll(f, &fmt->nSamplesPerSec, sizeof(fmt->nSamplesPerSec))) return false;
  if (!WriteAll(f, &fmt->nAvgBytesPerSec, sizeof(fmt->nAvgBytesPerSec))) return false;
  if (!WriteAll(f, &fmt->nBlockAlign, sizeof(fmt->nBlockAlign))) return false;
  if (!WriteAll(f, &fmt->wBitsPerSample, sizeof(fmt->wBitsPerSample))) return false;
  if (!WriteAll(f, "data", 4)) return false;
  if (!WriteAll(f, &data_size, sizeof(data_size))) return false;
  return true;
}

static bool WriteWavHeader(FILE* f, const WAVEFORMATEX* fmt) {
  if (!f || !fmt) return false;
  if (fseek(f, 0, SEEK_SET) != 0) return false;
  return RewriteCanonicalPcmHeader(f, fmt, 0);
}

static bool FinalizeWavFile(FILE* f, const WAVEFORMATEX* fmt) {
  if (!f || !fmt) return false;
  if (fflush(f) != 0) return false;
  if (fseek(f, 0, SEEK_END) != 0) return false;
  long end_pos = ftell(f);
  if (end_pos < 44) return false;
  DWORD data_size = (DWORD)(end_pos - 44);
  if (fseek(f, 0, SEEK_SET) != 0) return false;
  return RewriteCanonicalPcmHeader(f, fmt, data_size);
}

static bool AppendSilenceFrames(FILE* f, const WAVEFORMATEX* fmt, ULONGLONG frames) {
  if (!f || !fmt || fmt->nBlockAlign == 0) return false;
  if (frames == 0) return true;

  const size_t chunk_frames = 4096;
  std::vector<BYTE> zeros(chunk_frames * fmt->nBlockAlign, 0);
  while (frames > 0) {
    const size_t batch_frames = (frames > chunk_frames) ? chunk_frames : (size_t)frames;
    const size_t batch_bytes = batch_frames * fmt->nBlockAlign;
    if (!WriteAll(f, zeros.data(), batch_bytes)) return false;
    frames -= batch_frames;
  }
  return true;
}

static WAVEFORMATEX MakePcm16Format(const WAVEFORMATEX* src) {
  WAVEFORMATEX fmt{};
  fmt.wFormatTag = WAVE_FORMAT_PCM;
  fmt.nChannels = src ? src->nChannels : 2;
  if (fmt.nChannels == 0) fmt.nChannels = 2;
  if (fmt.nChannels > 2) fmt.nChannels = 2;
  fmt.nSamplesPerSec = src ? src->nSamplesPerSec : 44100;
  if (fmt.nSamplesPerSec == 0) fmt.nSamplesPerSec = 44100;
  fmt.wBitsPerSample = 16;
  fmt.nBlockAlign = (WORD)(fmt.nChannels * sizeof(short));
  fmt.nAvgBytesPerSec = fmt.nSamplesPerSec * fmt.nBlockAlign;
  fmt.cbSize = 0;
  return fmt;
}

static bool GetSourceFormatInfo(const WAVEFORMATEX* mix, WORD* out_bits, bool* out_is_float) {
  if (!mix || !out_bits || !out_is_float) return false;
  *out_bits = mix->wBitsPerSample;
  *out_is_float = false;

  if (mix->wFormatTag == WAVE_FORMAT_PCM) return true;
  if (mix->wFormatTag == WAVE_FORMAT_IEEE_FLOAT) {
    *out_is_float = true;
    return true;
  }
  if (mix->wFormatTag != WAVE_FORMAT_EXTENSIBLE) return false;

  auto* ext = reinterpret_cast<const WAVEFORMATEXTENSIBLE*>(mix);
  if (IsEqualGUID(ext->SubFormat, kSubtypePcm)) {
    *out_bits = ext->Samples.wValidBitsPerSample ? ext->Samples.wValidBitsPerSample : mix->wBitsPerSample;
    return true;
  }
  if (IsEqualGUID(ext->SubFormat, kSubtypeIeeeFloat)) {
    *out_is_float = true;
    *out_bits = ext->Samples.wValidBitsPerSample ? ext->Samples.wValidBitsPerSample : mix->wBitsPerSample;
    return true;
  }
  return false;
}

static float ReadNormalizedSample(const BYTE* src, WORD bits, bool is_float) {
  if (!src) return 0.0f;
  if (is_float && bits == 32) {
    float v = *reinterpret_cast<const float*>(src);
    if (v > 1.0f) v = 1.0f;
    if (v < -1.0f) v = -1.0f;
    return v;
  }
  if (bits == 8) {
    int v = (int)src[0] - 128;
    return (float)v / 128.0f;
  }
  if (bits == 16) {
    short v = *reinterpret_cast<const short*>(src);
    return (float)v / 32768.0f;
  }
  if (bits == 24) {
    int v = (int)src[0] | ((int)src[1] << 8) | ((int)src[2] << 16);
    if ((v & 0x00800000) != 0) v |= ~0x00FFFFFF;
    return (float)v / 8388608.0f;
  }
  if (bits == 32) {
    int v = *reinterpret_cast<const int*>(src);
    return (float)((double)v / 2147483648.0);
  }
  return 0.0f;
}

static short FloatToPcm16(float sample) {
  if (sample > 1.0f) sample = 1.0f;
  if (sample < -1.0f) sample = -1.0f;
  long v = lroundf(sample * 32767.0f);
  if (v > 32767) v = 32767;
  if (v < -32768) v = -32768;
  return (short)v;
}

static bool ConvertFramesToPcm16(const BYTE* src,
                                 UINT32 frames,
                                 const WAVEFORMATEX* mix,
                                 const WAVEFORMATEX* out_fmt,
                                 std::vector<BYTE>* out) {
  if (!src || !mix || !out_fmt || !out || mix->nChannels == 0 || out_fmt->nChannels == 0) return false;

  WORD bits = 0;
  bool is_float = false;
  if (!GetSourceFormatInfo(mix, &bits, &is_float)) return false;

  const UINT32 src_bytes_per_sample = (bits + 7) / 8;
  if (src_bytes_per_sample == 0) return false;
  const UINT32 src_frame_bytes = src_bytes_per_sample * mix->nChannels;
  const UINT32 dst_channels = out_fmt->nChannels;
  const UINT32 dst_frame_bytes = sizeof(short) * dst_channels;
  out->resize((size_t)frames * dst_frame_bytes);

  BYTE* dst = out->data();
  for (UINT32 frame = 0; frame < frames; ++frame) {
    const BYTE* src_frame = src + (size_t)frame * src_frame_bytes;
    short* dst_frame = reinterpret_cast<short*>(dst + (size_t)frame * dst_frame_bytes);
    if (dst_channels == 1) {
      float acc = 0.0f;
      for (UINT32 ch = 0; ch < mix->nChannels; ++ch) {
        acc += ReadNormalizedSample(src_frame + ch * src_bytes_per_sample, bits, is_float);
      }
      acc /= (float)mix->nChannels;
      dst_frame[0] = FloatToPcm16(acc);
      continue;
    }

    if (mix->nChannels == 1) {
      float sample = ReadNormalizedSample(src_frame, bits, is_float);
      dst_frame[0] = FloatToPcm16(sample);
      dst_frame[1] = FloatToPcm16(sample);
      continue;
    }

    float left = ReadNormalizedSample(src_frame, bits, is_float);
    float right = ReadNormalizedSample(src_frame + src_bytes_per_sample, bits, is_float);
    dst_frame[0] = FloatToPcm16(left);
    dst_frame[1] = FloatToPcm16(right);
    for (UINT32 ch = 2; ch < dst_channels; ++ch) {
      dst_frame[ch] = 0;
    }
  }
  return true;
}

struct CaptureThreadArgs {
  AudioRecorder* recorder{};
};

DWORD WINAPI CaptureThreadMain(void* param) {
  auto* args = reinterpret_cast<CaptureThreadArgs*>(param);
  AudioRecorder* rec = args ? args->recorder : nullptr;
  delete args;
  if (!rec) return 1;

  HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
  const bool coinit_ok = SUCCEEDED(hr) || hr == RPC_E_CHANGED_MODE;

  IMMDeviceEnumerator* enumerator = nullptr;
  IMMDevice* device = nullptr;
  IAudioClient* client = nullptr;
  IAudioCaptureClient* capture = nullptr;
  WAVEFORMATEX* mix = nullptr;
  WAVEFORMATEX out_fmt{};
  FILE* f = nullptr;

  auto fail = [&](const wchar_t* msg) {
    rec->error = msg ? msg : L"Audio recording failed.";
    rec->init_ok = false;
    rec->running = false;
    SetEvent(rec->ready_event);
  };

  if (!coinit_ok) {
    fail(L"Failed to initialize COM for audio capture.");
    goto cleanup;
  }

  hr = CoCreateInstance(CLSID_MMDeviceEnumerator, nullptr, CLSCTX_ALL,
                        IID_IMMDeviceEnumerator, reinterpret_cast<void**>(&enumerator));
  if (FAILED(hr)) {
    fail(L"Failed to create the audio device enumerator.");
    goto cleanup;
  }

  hr = enumerator->GetDefaultAudioEndpoint(SourceDataFlow(rec->source), eConsole, &device);
  if (FAILED(hr)) {
    if (rec->source == AudioRecordSource::MicrophoneCapture) {
      fail(L"No default microphone device is available.");
    } else {
      fail(L"No default playback device is available.");
    }
    goto cleanup;
  }

  hr = device->Activate(IID_IAudioClient, CLSCTX_ALL, nullptr, reinterpret_cast<void**>(&client));
  if (FAILED(hr)) {
    if (rec->source == AudioRecordSource::MicrophoneCapture) {
      fail(L"Failed to open the default microphone device.");
    } else {
      fail(L"Failed to open the default playback device.");
    }
    goto cleanup;
  }

  hr = client->GetMixFormat(&mix);
  if (FAILED(hr) || !mix) {
    fail(L"Failed to query the audio mix format.");
    goto cleanup;
  }
  out_fmt = MakePcm16Format(mix);

  hr = client->Initialize(AUDCLNT_SHAREMODE_SHARED, SourceStreamFlags(rec->source), 0, 0, mix, nullptr);
  if (FAILED(hr)) {
    if (rec->source == AudioRecordSource::MicrophoneCapture) {
      fail(L"Microphone recording is not supported on this device.");
    } else {
      fail(L"Loopback recording is not supported on this device.");
    }
    goto cleanup;
  }

  hr = client->GetService(IID_IAudioCaptureClient, reinterpret_cast<void**>(&capture));
  if (FAILED(hr)) {
    fail(L"Failed to get the audio capture service.");
    goto cleanup;
  }

  _wfopen_s(&f, rec->path.c_str(), L"wb");
  if (!f) {
    fail(L"Failed to create the WAV file.");
    goto cleanup;
  }
  if (!WriteWavHeader(f, &out_fmt)) {
    fail(L"Failed to write the WAV header.");
    goto cleanup;
  }

  hr = client->Start();
  if (FAILED(hr)) {
    if (rec->source == AudioRecordSource::MicrophoneCapture) {
      fail(L"Failed to start microphone recording.");
    } else {
      fail(L"Failed to start loopback recording.");
    }
    goto cleanup;
  }

  rec->start_tick = GetTickCount64();
  rec->written_frames = 0;
  rec->init_ok = true;
  rec->running = true;
  SetEvent(rec->ready_event);

  for (;;) {
    const bool stopping = (WaitForSingleObject(rec->stop_event, 5) == WAIT_OBJECT_0);
    UINT32 packet_frames = 0;
    hr = capture->GetNextPacketSize(&packet_frames);
    if (FAILED(hr)) {
      rec->error = L"Failed to read captured audio packets.";
      break;
    }

    while (packet_frames > 0) {
      BYTE* data = nullptr;
      UINT32 frames = 0;
      DWORD flags = 0;
      hr = capture->GetBuffer(&data, &frames, &flags, nullptr, nullptr);
      if (FAILED(hr)) {
        rec->error = L"Failed to read the capture buffer.";
        rec->running = false;
        goto cleanup;
      }

      if (frames > 0) {
        std::vector<BYTE> pcm16;
        if ((flags & AUDCLNT_BUFFERFLAGS_SILENT) != 0) {
          pcm16.resize((size_t)frames * out_fmt.nBlockAlign, 0);
        } else if (!ConvertFramesToPcm16(data, frames, mix, &out_fmt, &pcm16)) {
          capture->ReleaseBuffer(frames);
          rec->error = L"Failed to convert audio data to PCM.";
          rec->running = false;
          goto cleanup;
        }

        if (!WriteAll(f, pcm16.data(), pcm16.size())) {
          capture->ReleaseBuffer(frames);
          rec->error = L"Failed to write audio data to disk.";
          rec->running = false;
          goto cleanup;
        }
        rec->written_frames += frames;
      }

      capture->ReleaseBuffer(frames);
      hr = capture->GetNextPacketSize(&packet_frames);
      if (FAILED(hr)) {
        rec->error = L"Failed to continue reading captured audio.";
        rec->running = false;
        goto cleanup;
      }
    }
    if (stopping) break;
  }

cleanup:
  if (client) client->Stop();
  rec->running = false;
  if (rec->ready_event) SetEvent(rec->ready_event);

  if (f) {
    if (rec->init_ok && rec->error.empty()) {
      const ULONGLONG elapsed_ms = (rec->start_tick != 0) ? (GetTickCount64() - rec->start_tick) : 0;
      const ULONGLONG expected_frames =
          (elapsed_ms * (ULONGLONG)out_fmt.nSamplesPerSec + 999ULL) / 1000ULL;
      if (expected_frames > rec->written_frames) {
        const ULONGLONG missing_frames = expected_frames - rec->written_frames;
        if (!AppendSilenceFrames(f, &out_fmt, missing_frames)) {
          rec->error = L"Failed to finalize trailing silence in the WAV file.";
        } else {
          rec->written_frames = expected_frames;
        }
      }
    }
    if (rec->init_ok && rec->error.empty()) (void)FinalizeWavFile(f, &out_fmt);
    fclose(f);
    f = nullptr;
  }
  if (!rec->error.empty()) DeleteFileW(rec->path.c_str());

  if (mix) CoTaskMemFree(mix);
  SafeRelease(&capture);
  SafeRelease(&client);
  SafeRelease(&device);
  SafeRelease(&enumerator);
  if (coinit_ok) CoUninitialize();
  return 0;
}

}  // namespace

std::wstring GetAudioRecordingsDir(AudioRecordSource source) {
  SYSTEMTIME st{};
  GetLocalTime(&st);
  std::wstring root = JoinPath(GetDataRootDir(), SourceFolder(source));
  EnsureDirExists(root);
  std::wstring day = JoinPath(root, FormatDateYYYYMMDD(st));
  EnsureDirExists(day);
  return day;
}

bool AudioRecorderIsRunning(const AudioRecorder* recorder) {
  return recorder && recorder->running;
}

bool StartAudioRecorder(HWND, AudioRecordSource source, AudioRecorder** io_recorder, std::wstring* out_path, std::wstring* err) {
  if (out_path) out_path->clear();
  if (err) err->clear();
  if (!io_recorder) {
    if (err) *err = L"Recorder state is invalid.";
    return false;
  }
  if (*io_recorder && (*io_recorder)->running) {
    if (err) {
      std::wstring msg = L"A ";
      msg += SourceDisplayName((*io_recorder)->source);
      msg += L" recording is already running.";
      *err = msg;
    }
    return false;
  }

  auto* rec = new AudioRecorder();
  rec->source = source;
  rec->stop_event = CreateEventW(nullptr, TRUE, FALSE, nullptr);
  rec->ready_event = CreateEventW(nullptr, TRUE, FALSE, nullptr);
  if (!rec->stop_event || !rec->ready_event) {
    if (err) *err = L"Failed to create recorder control events.";
    if (rec->stop_event) CloseHandle(rec->stop_event);
    if (rec->ready_event) CloseHandle(rec->ready_event);
    delete rec;
    return false;
  }

  rec->path = JoinPath(GetAudioRecordingsDir(source), std::wstring(SourceStem(source)) + GetNowStamp() + L".wav");
  auto* args = new CaptureThreadArgs();
  args->recorder = rec;
  rec->thread = CreateThread(nullptr, 0, &CaptureThreadMain, args, 0, nullptr);
  if (!rec->thread) {
    if (err) *err = L"Failed to start the audio capture thread.";
    delete args;
    CloseHandle(rec->stop_event);
    CloseHandle(rec->ready_event);
    delete rec;
    return false;
  }

  WaitForSingleObject(rec->ready_event, INFINITE);
  if (!rec->init_ok) {
    WaitForSingleObject(rec->thread, INFINITE);
    if (err) {
      if (rec->error.empty()) {
        std::wstring msg = L"Failed to start ";
        msg += SourceDisplayName(source);
        msg += L" recording.";
        *err = msg;
      } else {
        *err = rec->error;
      }
    }
    CloseHandle(rec->thread);
    CloseHandle(rec->stop_event);
    CloseHandle(rec->ready_event);
    delete rec;
    return false;
  }

  *io_recorder = rec;
  if (out_path) *out_path = rec->path;
  return true;
}

bool StopAudioRecorder(AudioRecorder** io_recorder, std::wstring* out_path, std::wstring* err) {
  if (out_path) out_path->clear();
  if (err) err->clear();
  if (!io_recorder || !*io_recorder) {
    if (err) *err = L"No audio recording is currently running.";
    return false;
  }

  AudioRecorder* rec = *io_recorder;
  SetEvent(rec->stop_event);
  if (rec->thread) WaitForSingleObject(rec->thread, INFINITE);

  const bool ok = rec->error.empty();
  if (out_path) *out_path = rec->path;
  if (!ok && err) *err = rec->error;

  if (rec->thread) CloseHandle(rec->thread);
  if (rec->stop_event) CloseHandle(rec->stop_event);
  if (rec->ready_event) CloseHandle(rec->ready_event);
  delete rec;
  *io_recorder = nullptr;
  return ok;
}

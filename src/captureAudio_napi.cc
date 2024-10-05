#include <iostream>
#include <windows.h>
#include <mmdeviceapi.h>
#include <audiopolicy.h>
#include <vector>
#include <TlHelp32.h>
#include <shlobj.h>
#include <wchar.h>
#include <audioclientactivationparams.h>
#include <AudioClient.h>
#include <initguid.h>
#include <guiddef.h>
#include <mfapi.h>

#include <wrl\client.h>
#include <wrl\implements.h>
#include <wil\com.h>
#include <wil\result.h>

#include "Common.h"

#define NAPI_CPP_EXCEPTIONS 1
#include <node_api.h>
#include <random>
#include <napi.h>
#include <sstream>
#include <string>
#include <deque>
#include <queue>

using namespace Microsoft::WRL;
BYTE *m_AudioData;
UINT32 m_AudioDataSize;
bool isActivateCompleted = false;
bool isOnSampleReady = false;

class CLoopbackCapture : public RuntimeClass<RuntimeClassFlags<ClassicCom>, FtmBase, IActivateAudioInterfaceCompletionHandler>
{
public:
    wil::com_ptr_nothrow<IAudioClient> m_AudioClient;
    WAVEFORMATEX m_CaptureFormat{};
    UINT32 m_BufferFrames = 0;
    wil::com_ptr_nothrow<IAudioCaptureClient> m_AudioCaptureClient;
    wil::com_ptr_nothrow<IMFAsyncResult> m_SampleReadyAsyncResult;
    wil::unique_event_nothrow m_SampleReadyEvent;
    MFWORKITEM_KEY m_SampleReadyKey = 0;

    METHODASYNCCALLBACK(CLoopbackCapture, SampleReady, OnSampleReady);

    STDMETHODIMP ActivateCompleted(IActivateAudioInterfaceAsyncOperation *operation)
    {
        isActivateCompleted = !isActivateCompleted;
        HRESULT hrActivateResult = E_UNEXPECTED;
        wil::com_ptr_nothrow<IUnknown> punkAudioInterface;
        HRESULT hr = (operation->GetActivateResult(&hrActivateResult, &punkAudioInterface));
        RETURN_IF_FAILED(hrActivateResult);

        RETURN_IF_FAILED(punkAudioInterface.copy_to(&m_AudioClient));

        m_CaptureFormat.wFormatTag = WAVE_FORMAT_PCM;
        m_CaptureFormat.nChannels = 2;
        m_CaptureFormat.nSamplesPerSec = 44100;
        m_CaptureFormat.wBitsPerSample = 16;
        m_CaptureFormat.nBlockAlign = m_CaptureFormat.nChannels * m_CaptureFormat.wBitsPerSample / 8;
        m_CaptureFormat.nAvgBytesPerSec = m_CaptureFormat.nSamplesPerSec * m_CaptureFormat.nBlockAlign;

        RETURN_IF_FAILED(m_AudioClient->Initialize(AUDCLNT_SHAREMODE_SHARED,
                                                   AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                                                   200000,
                                                   AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
                                                   &m_CaptureFormat,
                                                   nullptr));

        RETURN_IF_FAILED(m_AudioClient->GetBufferSize(&m_BufferFrames));

        RETURN_IF_FAILED(m_AudioClient->GetService(IID_PPV_ARGS(&m_AudioCaptureClient)));

        RETURN_IF_FAILED(MFCreateAsyncResult(nullptr, &m_xSampleReady, nullptr, &m_SampleReadyAsyncResult));

        RETURN_IF_FAILED(m_AudioClient->SetEventHandle(m_SampleReadyEvent.get()));

        return S_OK;
    }

    HRESULT OnSampleReady(IMFAsyncResult *pResult)
    {
        isOnSampleReady = !isOnSampleReady;
        BYTE *data;
        UINT32 numFramesAvailable;
        DWORD flags;

        HRESULT hr = m_AudioCaptureClient->GetBuffer(&data, &numFramesAvailable, &flags, nullptr, nullptr);
        if (FAILED(hr))
        {
            return hr;
        }

        // 这里我们假设你有一个函数 returnAudioData(data, numFramesAvailable) 可以通过 V8 API 将音频数据返回给 Node.js
        // returnAudioData(data, numFramesAvailable);
        m_AudioData = new BYTE[numFramesAvailable * m_CaptureFormat.nBlockAlign];
        memcpy(m_AudioData, data, numFramesAvailable * m_CaptureFormat.nBlockAlign);
        m_AudioDataSize = numFramesAvailable * m_CaptureFormat.nBlockAlign;

        hr = m_AudioCaptureClient->ReleaseBuffer(numFramesAvailable);
        if (FAILED(hr))
        {
            return hr;
        }

        // 重新注册异步回调
        return MFPutWaitingWorkItem(m_SampleReadyEvent.get(), 0, m_SampleReadyAsyncResult.get(), &m_SampleReadyKey);
    }
};

Napi::Value StartCaputure(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();

    if (info.Length() < 1 || !info[0].IsNumber())
    {
        Napi::TypeError::New(env, "Expected a number as the first argument").ThrowAsJavaScriptException();
        return env.Null();
    }

    uint32_t arg0 = info[0].As<Napi::Number>().Uint32Value();
    DWORD processId = static_cast<DWORD>(arg0);

    AUDIOCLIENT_ACTIVATION_PARAMS audioclientActivationParams = {};
    audioclientActivationParams.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;
    audioclientActivationParams.ProcessLoopbackParams.ProcessLoopbackMode = PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE;
    audioclientActivationParams.ProcessLoopbackParams.TargetProcessId = processId;

    PROPVARIANT activateParams = {};
    activateParams.vt = VT_BLOB;
    activateParams.blob.cbSize = sizeof(audioclientActivationParams);
    activateParams.blob.pBlobData = (BYTE *)&audioclientActivationParams;

    wil::com_ptr_nothrow<IActivateAudioInterfaceAsyncOperation> asyncOp;
    Microsoft::WRL::ComPtr<CLoopbackCapture> completionHandler = Make<CLoopbackCapture>();
    HRESULT hr = ActivateAudioInterfaceAsync(
        VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK,
        __uuidof(IAudioClient),
        &activateParams,
        completionHandler.Get(),
        &asyncOp);
    return Napi::String::New(env, std::to_string(hr));
}

Napi::Value GetAudioData(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();

    Napi::Buffer<char> audioDataBuffer = Napi::Buffer<char>::Copy(env, (char *)m_AudioData, m_AudioDataSize);
    // return audioDataBuffer;

    std::stringstream ss;
    ss << "isActivateCompleted: " << isActivateCompleted << "\n"
       << "isOnSampleReady: " << isOnSampleReady << "\n";

    return Napi::String::New(env, ss.str().c_str());
}

WAVEFORMATEX *pWaveFormat = nullptr;
UINT32 nNumFramesAvailable = 0;
BYTE *pcaptured = nullptr;
UINT32 nNumFramesCaptured = 0;
DWORD dwFlags = 0;
std::queue<std::vector<BYTE>> audioQueue;
void capture(Napi::Env env)
{
    CoInitialize(NULL);

    IMMDeviceEnumerator *pEnumerator = nullptr;
    CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void **)&pEnumerator);

    IMMDevice *pDevice = nullptr;
    pEnumerator->GetDefaultAudioEndpoint(eCapture, eConsole, &pDevice);

    IAudioClient *pAudioClient = nullptr;
    pDevice->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, (void **)&pAudioClient);

    // WAVEFORMATEX *pWaveFormat = nullptr;
    pAudioClient->GetMixFormat(&pWaveFormat);

    pAudioClient->Initialize(AUDCLNT_SHAREMODE_SHARED, 0, 10000000, 0, pWaveFormat, nullptr);

    IAudioCaptureClient *pCaptureClient = nullptr;
    pAudioClient->GetService(__uuidof(IAudioCaptureClient), (void **)&pCaptureClient);

    pAudioClient->Start();

    UINT32 nFrameSize = 0;
    pAudioClient->GetBufferSize(&nFrameSize);

    BYTE *pData = new BYTE[nFrameSize * pWaveFormat->nBlockAlign];

    while (true)
    {
        std::this_thread::sleep_for(std::chrono::milliseconds(500));
        // UINT32 nNumFramesAvailable = 0;
        pCaptureClient->GetNextPacketSize(&nNumFramesAvailable);

        if (nNumFramesAvailable > 0)
        {
            // BYTE *pcaptured = nullptr;
            // UINT32 nNumFramesCaptured = 0;
            // DWORD dwFlags = 0;

            pCaptureClient->GetBuffer(&pcaptured, &nNumFramesCaptured, &dwFlags, nullptr, nullptr);

            // 在这里处理捕获到的音频数据
            /*for (UINT32 i = 0; i < nNumFramesCaptured * pWaveFormat->nBlockAlign; i++)
            {
                std::cout << static_cast<int>(pcaptured[i]) << " ";
            }*/

            std::vector<BYTE> audioData(pcaptured, pcaptured + nNumFramesCaptured * sizeof(BYTE));
            audioQueue.push(audioData);

            pCaptureClient->ReleaseBuffer(nNumFramesCaptured);
        }
    }

    pAudioClient->Stop();
    pAudioClient->Release();
    pDevice->Release();
    pEnumerator->Release();
    CoUninitialize();
}
Napi::Value DefaultCapture(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    std::thread(capture, env).detach();
    return env.Null();
}
Napi::Value getpcapture(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    // Napi::Buffer<BYTE> audioBuffer = Napi::Buffer<BYTE>::Copy(env, pcaptured, nNumFramesCaptured * sizeof(BYTE));
    // return audioBuffer;
    if (!audioQueue.empty())
    {
        std::vector<BYTE> audioData = audioQueue.front();
        audioQueue.pop();

        Napi::Buffer<BYTE> audioBuffer = Napi::Buffer<BYTE>::Copy(env, audioData.data(), audioData.size());
        return audioBuffer;
    }

    return env.Null();
}
Napi::Value getnNumFramesCaptured(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    double num = static_cast<double>(nNumFramesCaptured);
    return Napi::Number::New(env, num);
}
Napi::Value getdwFlags(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    std::stringstream ss;
    bool silence;
    if (dwFlags & AUDCLNT_BUFFERFLAGS_SILENT)
    {
        silence = true;
    }
    ss << "flags:" << dwFlags << "ABS:" << AUDCLNT_BUFFERFLAGS_SILENT << "silence:" << silence;
    std::string str = ss.str();
    return Napi::String::New(env, str);
}
Napi::Value getnNumFramesAvailable(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    double num = static_cast<double>(nNumFramesAvailable);
    return Napi::Number::New(env, num);
}
Napi::Value getAudioFormat(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    Napi::Object res = Napi::Object::New(env);
    res.Set("channels", Napi::Number::New(env, pWaveFormat->nChannels));
    res.Set("sampleRate", Napi::Number::New(env, pWaveFormat->nSamplesPerSec));
    res.Set("bitsPerSample", Napi::Number::New(env, pWaveFormat->wBitsPerSample));
    res.Set("avgBytesPerSec", Napi::Number::New(env, pWaveFormat->nAvgBytesPerSec));
    return res;
}

Napi::Object Init(Napi::Env env, Napi::Object exports)
{
    exports.Set(Napi::String::New(env, "startCapture"), Napi::Function::New(env, StartCaputure));
    exports.Set(Napi::String::New(env, "getAudioData"), Napi::Function::New(env, GetAudioData));

    exports.Set(Napi::String::New(env, "defaultCaputure"), Napi::Function::New(env, DefaultCapture));
    exports.Set(Napi::String::New(env, "getpcapture"), Napi::Function::New(env, getpcapture));
    exports.Set(Napi::String::New(env, "getnNumFramesCaptured"), Napi::Function::New(env, getnNumFramesCaptured));
    exports.Set(Napi::String::New(env, "getdwFlags"), Napi::Function::New(env, getdwFlags));
    exports.Set(Napi::String::New(env, "getnNumFramesAvailable"), Napi::Function::New(env, getnNumFramesAvailable));
    exports.Set(Napi::String::New(env, "getAudioFormat"), Napi::Function::New(env, getAudioFormat));

    return exports;
}

NODE_API_MODULE(myncpp1, Init)
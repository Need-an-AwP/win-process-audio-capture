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
#include <propsys.h>
#include <propkey.h>
#include <codecvt>
#include <locale>
#include <fstream>
#include <iomanip>
#include <cstring>
#include <cstdint>

using namespace Microsoft::WRL;

#define REFTIMES_PER_SEC 10000000
#define REFTIMES_PER_MILLISEC 10000

#pragma pack(push, 1)
struct WAVHeader
{
    char riff[4] = {'R', 'I', 'F', 'F'};
    uint32_t chunkSize;
    char wave[4] = {'W', 'A', 'V', 'E'};
    char fmt[4] = {'f', 'm', 't', ' '};
    uint32_t fmtChunkSize = 16;
    uint16_t audioFormat = 1; // PCM = 1
    uint16_t numChannels;
    uint32_t sampleRate;
    uint32_t byteRate;
    uint16_t blockAlign;
    uint16_t bitsPerSample;
    char data[4] = {'d', 'a', 't', 'a'};
    uint32_t dataChunkSize;
};
#pragma pack(pop)

REFERENCE_TIME hnsRequestedDuration = REFTIMES_PER_SEC;
REFERENCE_TIME hnsActualDuration;
UINT32 bufferFrameCount;
WAVEFORMATEX *pWaveFormat = nullptr;
UINT32 packetLength = 0;
BYTE *pcaptured = nullptr;
UINT32 nNumFramesCaptured = 0;
DWORD dwFlags = 0;
IAudioCaptureClient *pCaptureClient = nullptr;
wil::com_ptr_nothrow<IAudioCaptureClient> mAudioCaptureClient = nullptr;
WAVEFORMATEX *mCaptureFormat = new WAVEFORMATEX;
UINT32 mBufferFrames;
wil::com_ptr_nothrow<IActivateAudioInterfaceAsyncOperation> asyncOp;

wil::com_ptr_nothrow<IAudioClient> g_AudioClient;

Napi::Value getAudioProcessInfo(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    HRESULT hr;
    IMMDeviceEnumerator *pEnumerator = NULL;
    IMMDevice *pDevice = NULL;
    IAudioSessionManager2 *pSessionManager = NULL;
    IAudioSessionEnumerator *pSessionEnumerator = NULL;
    CoInitialize(NULL);
    CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void **)&pEnumerator);
    pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);
    pDevice->Activate(__uuidof(IAudioSessionManager2), CLSCTX_ALL, NULL, (void **)&pSessionManager);
    pSessionManager->GetSessionEnumerator(&pSessionEnumerator);
    int sessionCount = 0;
    pSessionEnumerator->GetCount(&sessionCount);
    HANDLE hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);

    Napi::Array resultArray = Napi::Array::New(env);
    int index = 0;
    for (int i = 0; i < sessionCount; i++)
    {
        IAudioSessionControl *pSessionControl = NULL;
        IAudioSessionControl2 *pSessionControl2 = NULL;
        hr = pSessionEnumerator->GetSession(i, &pSessionControl);
        if (SUCCEEDED(hr))
        {
            // Query for IAudioSessionControl2 interface
            hr = pSessionControl->QueryInterface(__uuidof(IAudioSessionControl2), (void **)&pSessionControl2);
            if (SUCCEEDED(hr))
            {
                // Get the process ID of the audio session
                DWORD processId;
                hr = pSessionControl2->GetProcessId(&processId);
                if (SUCCEEDED(hr))
                {
                    // Find the process name
                    PROCESSENTRY32 pe32;
                    pe32.dwSize = sizeof(PROCESSENTRY32);
                    if (Process32First(hSnapshot, &pe32))
                    {
                        do
                        {
                            if (pe32.th32ProcessID == processId)
                            {
                                Napi::Object sessionInfo = Napi::Object::New(env);
                                sessionInfo.Set("processId", Napi::Number::From(env, processId));
                                sessionInfo.Set("processName", Napi::String::From(env, pe32.szExeFile));
                                resultArray.Set(index++, sessionInfo);
                                break;
                            }
                        } while (Process32Next(hSnapshot, &pe32));
                    }
                }
                pSessionControl2->Release();
            }
            pSessionControl->Release();
        }
    }

    CloseHandle(hSnapshot);
    pSessionEnumerator->Release();
    pSessionManager->Release();
    pDevice->Release();
    pEnumerator->Release();
    CoUninitialize();

    return resultArray;
}

Napi::Value initializeCapture(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    CoInitialize(NULL);

    IMMDeviceEnumerator *pEnumerator = nullptr;
    CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void **)&pEnumerator);

    IMMDevice *pDefaultDevice = nullptr;
    // pEnumerator->GetDefaultAudioEndpoint(eCapture, eConsole, &pDefaultDevice);
    pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDefaultDevice);

    IMMDevice *pSelectedDevice;
    IMMDeviceCollection *pDevices;
    pEnumerator->EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE, &pDevices);
    UINT deviceCount;
    pDevices->GetCount(&deviceCount);
    bool forLoopBreak;
    forLoopBreak = false;
    for (UINT i = 0; i < deviceCount; i++)
    {
        IMMDevice *pDevice = NULL;
        pDevices->Item(i, &pDevice);

        IMMEndpoint *pEndpoint;
        pDevice->QueryInterface(__uuidof(IMMEndpoint), (void **)&pEndpoint);

        IPropertyStore *pProps;
        pDevice->OpenPropertyStore(STGM_READ, &pProps);
        DWORD propertyCount;
        pProps->GetCount(&propertyCount);

        for (DWORD j = 0; j < propertyCount; j++)
        {
            PROPERTYKEY key;
            pProps->GetAt(j, &key);
            PROPVARIANT varName;
            PropVariantInit(&varName);
            pProps->GetValue(key, &varName);
            switch (varName.vt)
            {
            case VT_LPWSTR:
                if (varName.pwszVal)
                {
                    std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
                    std::string utf8Str = converter.to_bytes(varName.pwszVal);
                    // std::cout << utf8Str << "\n";
                    if (utf8Str == "Realtek(R) Audio")
                    {
                        pSelectedDevice = pDevice;
                        forLoopBreak = true;
                    }
                }
                break;
            default:
                break;
            }
            PropVariantClear(&varName);
            if (forLoopBreak)
            {
                break;
            }
        }
        if (forLoopBreak)
        {
            break;
        }
    }

    IAudioClient *pAudioClient = nullptr;
    // pDefaultDevice->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, (void **)&pAudioClient);
    pSelectedDevice->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, (void **)&pAudioClient);

    pAudioClient->GetMixFormat(&pWaveFormat);

    pAudioClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        AUDCLNT_STREAMFLAGS_NOPERSIST, // AUDCLNT_STREAMFLAGS_LOOPBACK, //
        hnsRequestedDuration,
        0,
        pWaveFormat,
        nullptr);

    pAudioClient->GetService(__uuidof(IAudioCaptureClient), (void **)&pCaptureClient);

    pAudioClient->Start();

    pAudioClient->GetBufferSize(&bufferFrameCount);

    hnsActualDuration = (double)REFTIMES_PER_SEC * bufferFrameCount / pWaveFormat->nSamplesPerSec;

    // return env.Null();
    return Napi::Number::From(env, bufferFrameCount);
}

Napi::Value getAudioFormat(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    Napi::Object res = Napi::Object::New(env);
    res.Set("cbSize", Napi::Number::New(env, pWaveFormat->cbSize));
    res.Set("nAvgBytesPerSec", Napi::Number::New(env, pWaveFormat->nAvgBytesPerSec));
    res.Set("nBlockAlign", Napi::Number::New(env, pWaveFormat->nBlockAlign));
    res.Set("nChannels", Napi::Number::New(env, pWaveFormat->nChannels));
    res.Set("wBitsPerSample", Napi::Number::New(env, pWaveFormat->wBitsPerSample));
    res.Set("nSamplesPerSec", Napi::Number::New(env, pWaveFormat->nSamplesPerSec));
    res.Set("wFormatTag", Napi::Number::New(env, pWaveFormat->wFormatTag));

    return res;
}

Napi::Object ConstructAudioBufferData(const Napi::Env &env, const WAVEFORMATEX *wfx, const std::vector<uint8_t> &pcmData)
{
    const uint32_t numChannels = wfx->nChannels;
    const uint32_t length = pcmData.size() / (wfx->wBitsPerSample / 8) / numChannels;

    std::vector<std::vector<float>> channelData(numChannels, std::vector<float>(length));

    const uint32_t bytesPerSample = wfx->wBitsPerSample / 8;
    const int32_t maxValue = (1 << (wfx->wBitsPerSample - 1)) - 1;

    for (uint32_t i = 0; i < length; ++i)
    {
        for (uint32_t channel = 0; channel < numChannels; ++channel)
        {
            int32_t sample = 0;
            uint8_t value = 0;
            for (uint32_t byte = 0; byte < bytesPerSample; ++byte)
            {
                value = pcmData[i * numChannels * bytesPerSample + channel * bytesPerSample + byte];
                sample |= static_cast<int32_t>(value) << (byte * 8);
            }

            float normalizedSample = static_cast<float>(sample) / maxValue;
            // float normalizedSample = (static_cast<float>(sample) - 128.0) / 128.0;
            // channelData[channel][i] = normalizedSample;
            channelData[channel][i] = value;
        }
    }

    Napi::Object audioBuffer = Napi::Object::New(env);

    Napi::Object subObject = Napi::Object::New(env);
    audioBuffer.Set("constructAudioBufferData", subObject);

    subObject.Set("nSampleRate", static_cast<double>(wfx->nSamplesPerSec));
    subObject.Set("length", static_cast<double>(length));
    subObject.Set("nChannels", static_cast<double>(numChannels));
    subObject.Set("wBitsPerSample", static_cast<double>(wfx->wBitsPerSample));

    for (uint32_t channel = 0; channel < numChannels; ++channel)
    {
        Napi::Float32Array channelBuffer = Napi::Float32Array::New(env, length);
        for (size_t i = 0; i < length; i++)
        {
            channelBuffer[i] = channelData[channel][i];
        }
        subObject.Set(std::to_string(channel), channelBuffer);
    }

    return audioBuffer;
}

Napi::Value getBuffer(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    pCaptureClient->GetNextPacketSize(&packetLength);
    if (packetLength != 0)
    {
        pCaptureClient->GetBuffer(&pcaptured, &nNumFramesCaptured, &dwFlags, nullptr, nullptr);
        std::vector<BYTE> audioData(pcaptured, pcaptured + nNumFramesCaptured * sizeof(BYTE));
        Napi::Buffer<BYTE> audioBuffer = Napi::Buffer<BYTE>::Copy(env, audioData.data(), audioData.size());

        Napi::Object res = ConstructAudioBufferData(env, pWaveFormat, audioData);

        Napi::Buffer<BYTE> omd = Napi::Buffer<BYTE>::Copy(env, pcaptured, nNumFramesCaptured * sizeof(BYTE));
        res.Set("originMemoryData", omd);

        double numC = static_cast<double>(nNumFramesCaptured);
        res.Set("nNumFramesCaptured", numC);

        std::stringstream ss;
        bool silence;
        if (dwFlags & AUDCLNT_BUFFERFLAGS_SILENT)
        {
            silence = true;
        }
        ss << "flags:" << dwFlags << "\nABS:" << AUDCLNT_BUFFERFLAGS_SILENT << "\nsilence:" << silence;
        std::string str = ss.str();
        res.Set("flags", str);

        double numA = static_cast<double>(packetLength);
        res.Set("nNumFramesAvailable", numA);

        double sleepDuration = hnsActualDuration / REFTIMES_PER_MILLISEC / 2;
        res.Set("sleepDuration", sleepDuration);

        res.Set("GetBufferSize", Napi::Number::From(env, bufferFrameCount));

        pCaptureClient->ReleaseBuffer(nNumFramesCaptured);
        return res;
    }
    else
    {
        return env.Null();
    }
}

Napi::Value getHalfSecWAV(const Napi::CallbackInfo &info)
{
    std::vector<uint8_t> pcmData;
    UINT32 totalFramesCaptured = 0;
    Napi::Env env = info.Env();
    pCaptureClient->GetNextPacketSize(&packetLength);
    UINT8 n = 0;
    while (packetLength != 0)
    {
        pCaptureClient->GetBuffer(
            &pcaptured,
            &nNumFramesCaptured,
            &dwFlags,
            nullptr,
            nullptr);
        if (dwFlags & AUDCLNT_BUFFERFLAGS_SILENT)
        {
            pcaptured = NULL;
        }
        totalFramesCaptured += nNumFramesCaptured;
        pcmData.insert(pcmData.end(), pcaptured, pcaptured + nNumFramesCaptured * pWaveFormat->nBlockAlign);

        pCaptureClient->ReleaseBuffer(nNumFramesCaptured);
        pCaptureClient->GetNextPacketSize(&packetLength);
        n++;
    }

    Napi::Object res = ConstructAudioBufferData(env, pWaveFormat, pcmData);

    res.Set("totalFramesCaptured", static_cast<int>(totalFramesCaptured));
    res.Set("n", Napi::Number::From(env, n));
    Napi::Uint8Array pcmArray = Napi::Uint8Array::New(env, pcmData.size());
    for (size_t i = 0; i < pcmData.size(); i++)
    {
        pcmArray[i] = pcmData[i];
    }
    res.Set("pcmData", pcmArray);
    return res;
}

Napi::Value getWAVfromfile(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    std::ifstream file("C:/Users/17904/Desktop/voiceChat/test/src/loopback-capture.wav", std::ios::binary);
    file.seekg(0, std::ios::end);
    size_t fileSize = file.tellg();
    file.seekg(0, std::ios::beg);
    std::vector<uint8_t> wavData(fileSize);
    file.read(reinterpret_cast<char *>(wavData.data()), fileSize);
    file.close();

    WAVEFORMATEX wfx = {0};
    wfx.wFormatTag = *(reinterpret_cast<uint16_t *>(wavData.data() + 20));
    wfx.nChannels = *(reinterpret_cast<uint16_t *>(wavData.data() + 22));
    wfx.nSamplesPerSec = *(reinterpret_cast<uint32_t *>(wavData.data() + 24));
    wfx.wBitsPerSample = *(reinterpret_cast<uint16_t *>(wavData.data() + 34));
    wfx.nBlockAlign = *(reinterpret_cast<uint16_t *>(wavData.data() + 32)); // 2*32/8=8
    wfx.nAvgBytesPerSec = wfx.nSamplesPerSec * wfx.nBlockAlign;

    size_t dataMarkerPos = 0;
    bool found = false;
    for (size_t i = 0; i < wavData.size() - 3; i++)
    {
        if (wavData[i] == 0x64 && wavData[i + 1] == 0x61 && wavData[i + 2] == 0x74 && wavData[i + 3] == 0x61)
        {
            dataMarkerPos = i;
            found = true;
            break;
        }
    }
    uint32_t audioDataLength = *(reinterpret_cast<uint32_t *>(wavData.data() + dataMarkerPos + 4));
    size_t audioDataOffset = dataMarkerPos + 8;

    std::vector<uint8_t> WAVpcmData(wavData.begin() + audioDataOffset,
                                    wavData.begin() + audioDataOffset + audioDataLength);
    // Napi::Object res = ConstructAudioBufferData(env, &wfx, WAVpcmData);
    Napi::Object res = Napi::Object::New(env);

    Napi::Uint8Array wavArray = Napi::Uint8Array::New(env, wavData.size());
    for (size_t i = 0; i < wavData.size(); i++)
    {
        wavArray[i] = wavData[i];
    }
    res.Set("wavData", wavArray);

    Napi::Uint8Array pcmArray = Napi::Uint8Array::New(env, WAVpcmData.size());
    for (size_t i = 0; i < WAVpcmData.size(); i++)
    {
        pcmArray[i] = WAVpcmData[i];
    }
    res.Set("pcmData", pcmArray);
    res.Set("WAVfileSize", Napi::Number::New(env, fileSize));
    res.Set("wFormatTag", Napi::Number::From(env, wfx.wFormatTag));
    res.Set("wBitsPerSample", Napi::Number::From(env, wfx.wBitsPerSample));
    res.Set("nBlockAlign", Napi::Number::From(env, wfx.nBlockAlign));

    return res;
}

class AudioInterfaceCompletionHandler : public RuntimeClass<RuntimeClassFlags<ClassicCom>, FtmBase, IActivateAudioInterfaceCompletionHandler>
{
public:
    wil::com_ptr_nothrow<IAudioClient> m_AudioClient;
    HRESULT m_ActivateCompletedResult = S_OK;
    HRESULT audioClientInitializeResult;
    HRESULT audioClientStartResult;
    HRESULT captureClientResult;
    WAVEFORMATEX m_CaptureFormat{};
    UINT32 m_BufferFrames = 666;

    STDMETHOD(ActivateCompleted)
    (IActivateAudioInterfaceAsyncOperation *operation)
    {
        HRESULT hrActivateResult = E_UNEXPECTED;
        wil::com_ptr_nothrow<IUnknown> punkAudioInterface;
        operation->GetActivateResult(&hrActivateResult, &punkAudioInterface);
        if (SUCCEEDED(hrActivateResult))
        {
            m_ActivateCompletedResult = S_OK;
            punkAudioInterface.copy_to(&m_AudioClient);

            m_CaptureFormat.wFormatTag = WAVE_FORMAT_PCM;
            m_CaptureFormat.nChannels = 2;
            m_CaptureFormat.nSamplesPerSec = 44100; // 48000; //
            m_CaptureFormat.wBitsPerSample = 16;    // 32;    //
            m_CaptureFormat.nBlockAlign = m_CaptureFormat.nChannels * m_CaptureFormat.wBitsPerSample / 8;
            m_CaptureFormat.nAvgBytesPerSec = m_CaptureFormat.nSamplesPerSec * m_CaptureFormat.nBlockAlign;

            audioClientInitializeResult = m_AudioClient->Initialize(AUDCLNT_SHAREMODE_SHARED,
                                                                    AUDCLNT_STREAMFLAGS_LOOPBACK,       //| AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
                                                                    hnsRequestedDuration,               // 200000,
                                                                    AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM, // 0,
                                                                    &m_CaptureFormat,
                                                                    NULL);
            if (SUCCEEDED(audioClientInitializeResult))
            {
                std::memcpy(mCaptureFormat, &m_CaptureFormat, sizeof(WAVEFORMATEX));
                captureClientResult = m_AudioClient->GetService(IID_PPV_ARGS(&mAudioCaptureClient));
                audioClientStartResult = m_AudioClient->Start();
                m_AudioClient->GetBufferSize(&mBufferFrames);
                // m_AudioClient->GetBufferSize(&m_BufferFrames);
                g_AudioClient = m_AudioClient;
            }
        }
        else
        {
            m_ActivateCompletedResult = hrActivateResult;
        }
        return S_OK;
    }
    STDMETHOD(GetActivateCompletedResult)
    (HRESULT *pResult)
    {
        *pResult = m_ActivateCompletedResult;
        return S_OK;
    }
    STDMETHOD(GetAudioClientInitializeResult)
    (HRESULT *pResult)
    {
        *pResult = audioClientInitializeResult;
        return S_OK;
    }
    STDMETHOD(GetCaptureClientResult)
    (HRESULT *pResult)
    {
        *pResult = captureClientResult;
        return S_OK;
    }
    STDMETHOD(GetStartResult)
    (HRESULT *pResult)
    {
        *pResult = audioClientStartResult;
        return S_OK;
    }
    STDMETHOD(GetBufferSize)
    (UINT32 *pResult)
    {
        *pResult = m_BufferFrames;
        return S_OK;
    }

    ~AudioInterfaceCompletionHandler()
    {
    }
};

Napi::Value initializeCLoopbackCapture(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    CoInitialize(NULL);
    HRESULT hr;
    if (info.Length() < 1 || !info[0].IsNumber())
    {
        Napi::TypeError::New(env, "Expected a number as the first argument").ThrowAsJavaScriptException();
        return env.Null();
    }
    // CoInitialize(NULL);
    uint32_t arg0 = info[0].As<Napi::Number>().Uint32Value();
    DWORD processId = static_cast<DWORD>(arg0);
    bool includeProcessTree = true;
    AUDIOCLIENT_ACTIVATION_PARAMS audioclientActivationParams = {};
    audioclientActivationParams.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;
    audioclientActivationParams.ProcessLoopbackParams.ProcessLoopbackMode = includeProcessTree ? PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE : PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE;
    audioclientActivationParams.ProcessLoopbackParams.TargetProcessId = processId;

    PROPVARIANT activateParams = {};
    activateParams.vt = VT_BLOB;
    activateParams.blob.cbSize = sizeof(audioclientActivationParams);
    activateParams.blob.pBlobData = (BYTE *)&audioclientActivationParams;

    Microsoft::WRL::ComPtr<AudioInterfaceCompletionHandler> completionHandler = Make<AudioInterfaceCompletionHandler>();
    hr = ActivateAudioInterfaceAsync(VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK,
                                     __uuidof(IAudioClient),
                                     &activateParams,
                                     completionHandler.Get(),
                                     &asyncOp);

    Napi::Object res = Napi::Object::New(env);
    res.Set("ActivateAudioInterfaceAsyncResult", Napi::Number::New(env, static_cast<double>(hr)));
    HRESULT activateCompletedResult;
    completionHandler->GetActivateCompletedResult(&activateCompletedResult);
    res.Set("ActivateCompletedResult", Napi::Number::From(env, static_cast<double>(activateCompletedResult)));
    HRESULT audioClientInitializeResult;
    completionHandler->GetAudioClientInitializeResult(&audioClientInitializeResult);
    res.Set("audioClientInitializeResult", Napi::Number::From(env, static_cast<double>(audioClientInitializeResult)));
    HRESULT captureClientResult;
    completionHandler->GetAudioClientInitializeResult(&captureClientResult);
    res.Set("captureClientResult", Napi::Number::From(env, static_cast<double>(captureClientResult)));
    HRESULT audioClientStartResult;
    completionHandler->GetStartResult(&audioClientStartResult);
    res.Set("audioClientStartResult", Napi::Number::From(env, static_cast<double>(audioClientStartResult)));
    res.Set("processId", Napi::Number::From(env, processId));
    UINT32 m_BufferFrames;
    completionHandler->GetBufferSize(&m_BufferFrames);
    res.Set("m_BufferFrames", Napi::Number::From(env, m_BufferFrames));
    return res;
}

Napi::Value getActivateStatus(const Napi::CallbackInfo &info)
{
    HRESULT activateResult;
    Napi::Object result = Napi::Object::New(info.Env());
    if (asyncOp)
    {
        wil::com_ptr_nothrow<IUnknown> activatedInterface;
        HRESULT hr = asyncOp->GetActivateResult(&activateResult, &activatedInterface);
        if (SUCCEEDED(hr))
        {
            result.Set("interfaceActivateResult", Napi::Number::New(info.Env(), static_cast<double>(activateResult)));
            return result;
        }
        else
        {
            result.Set("interfaceActivateResult", Napi::Number::New(info.Env(), static_cast<double>(hr)));
            return result;
        }
    }
    else
    {
        result.Set("error", Napi::Number::New(info.Env(), static_cast<double>(E_UNEXPECTED)));
        return result;
    }
}

Napi::Value getProcessCaptureFormat(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    if (mCaptureFormat == nullptr)
    {
        return Napi::String::New(env, "mCaptureFormat is not initialized");
    }
    else
    {
        Napi::Object res = Napi::Object::New(env);
        res.Set("cbSize", Napi::Number::New(env, mCaptureFormat->cbSize));
        res.Set("nAvgBytesPerSec", Napi::Number::New(env, mCaptureFormat->nAvgBytesPerSec));
        res.Set("nBlockAlign", Napi::Number::New(env, mCaptureFormat->nBlockAlign));
        res.Set("nChannels", Napi::Number::New(env, mCaptureFormat->nChannels));
        res.Set("wBitsPerSample", Napi::Number::New(env, mCaptureFormat->wBitsPerSample));
        res.Set("nSamplesPerSec", Napi::Number::New(env, mCaptureFormat->nSamplesPerSec));
        res.Set("wFormatTag", Napi::Number::New(env, mCaptureFormat->wFormatTag));

        res.Set("GetBufferSize", Napi::Number::From(env, mBufferFrames));
        double sleepDuration = ((double)REFTIMES_PER_SEC * mBufferFrames / mCaptureFormat->nSamplesPerSec) / REFTIMES_PER_MILLISEC / 2;
        res.Set("sleepDuration", sleepDuration);
        return res;
    }
}

Napi::Value captureProcessAudio(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    HRESULT hr;
    std::vector<uint8_t> pcmData;

    BYTE *Data = nullptr;
    UINT32 nNumFramesCaptured = 0;
    UINT32 packetLength = 0;
    DWORD dwCaptureFlags;
    UINT64 u64DevicePosition = 0;
    UINT64 u64QPCPosition = 0;
    DWORD cbBytesToCapture = 0;
    Napi::Array pcmDataArray = Napi::Array::New(env);
    mAudioCaptureClient->GetNextPacketSize(&packetLength);

    if (packetLength > 0)
    {
        cbBytesToCapture = nNumFramesCaptured * mCaptureFormat->nBlockAlign;
        hr = mAudioCaptureClient->GetBuffer(&Data,
                                            &nNumFramesCaptured,
                                            &dwCaptureFlags,
                                            NULL,  //&u64DevicePosition,
                                            NULL); //&u64QPCPosition);
        if (FAILED(hr))
        {
            return Napi::String::New(env, std::to_string(hr));
        }
        // pcmData.insert(pcmData.end(), Data, Data + nNumFramesCaptured * mCaptureFormat->nBlockAlign);

        mAudioCaptureClient->GetNextPacketSize(&nNumFramesCaptured);
        Napi::Object res = Napi::Object::New(env);
        Napi::Buffer<BYTE> omd = Napi::Buffer<BYTE>::Copy(env, Data, nNumFramesCaptured * sizeof(BYTE));
        res.Set("originMemoryData", omd);
        res.Set("dwFlags", Napi::Number::From(env, dwCaptureFlags));
        res.Set("GetNextPacketSize", Napi::Number::From(env, packetLength));
        res.Set("GetBuffer_nNumFramesCaptured", Napi::Number::From(env, nNumFramesCaptured));
        /*Napi::Object res = ConstructAudioBufferData(env, mCaptureFormat, pcmData);
        Napi::Uint8Array pcmArray = Napi::Uint8Array::New(env, pcmData.size());
        for (size_t i = 0; i < pcmData.size(); i++)
        {
            pcmArray[i] = pcmData[i];
        }
        res.Set("pcmData", pcmArray);*/

        mAudioCaptureClient->ReleaseBuffer(nNumFramesCaptured);

        return res;
    }
    else
    {
        Napi::Object res = Napi::Object::New(env);
        res.Set("GetNextPacketSize", Napi::Number::From(env, packetLength));
        return res;
    }
}

Napi::Value whileCaptureProcessAudio(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    UINT32 packetLength = 0;
    BYTE *Data = nullptr;
    HRESULT hr;
    DWORD dwCaptureFlags;
    UINT32 nNumFramesCaptured = 0;
    UINT64 u64DevicePosition = 0;
    UINT64 u64QPCPosition = 0;
    UINT32 totalFramesCaptured = 0;
    std::vector<uint8_t> pcmData;
    std::vector<UINT32> capturedFramesList;
    std::vector<UINT32> packetSizeList;
    mAudioCaptureClient->GetNextPacketSize(&packetLength);
    if (packetLength > 0)
    {
        packetSizeList.push_back(packetLength);
        UINT8 n = 0;
        std::vector<Napi::Buffer<BYTE>> rawDataList;
        while (packetLength > 0)
        {
            hr = mAudioCaptureClient->GetBuffer(&Data,
                                                &nNumFramesCaptured,
                                                &dwCaptureFlags,
                                                NULL,  //&u64DevicePosition,
                                                NULL); //&u64QPCPosition);
            capturedFramesList.push_back(nNumFramesCaptured);
            totalFramesCaptured += nNumFramesCaptured;
            pcmData.insert(pcmData.end(), Data, Data + nNumFramesCaptured * mCaptureFormat->nBlockAlign);
            Napi::Buffer<BYTE> rawData = Napi::Buffer<BYTE>::Copy(env, Data, nNumFramesCaptured * mCaptureFormat->nBlockAlign);
            rawDataList.push_back(rawData);
            mAudioCaptureClient->ReleaseBuffer(nNumFramesCaptured);
            mAudioCaptureClient->GetNextPacketSize(&packetLength);
            n++;
            packetSizeList.push_back(packetLength);
        }

        Napi::Object res = Napi::Object::New(env);

        UINT32 numPaddingFrames = 0;
        if (g_AudioClient)
        {
            hr = g_AudioClient->GetCurrentPadding(&numPaddingFrames);
            if (SUCCEEDED(hr))
            {
                res.Set("currentPaddingFrames", Napi::Number::New(env, static_cast<double>(numPaddingFrames)));
            }
            else
            {
                res.Set("currentPaddingFrames", Napi::Number::New(env, static_cast<double>(hr)));
            }
        }
        else
        {
            res.Set("currentPaddingFrames", Napi::Number::New(env, static_cast<double>(E_UNEXPECTED)));
        }

        Napi::Array rawDataArray = Napi::Array::New(env, rawDataList.size());
        for (size_t i = 0; i < rawDataList.size(); i++)
        {
            rawDataArray[i] = rawDataList[i];
        }
        res.Set("rawData", rawDataArray);
        res.Set("totalFramesCaptured", Napi::Number::From(env, totalFramesCaptured));
        res.Set("n", Napi::Number::From(env, n));

        Napi::Uint8Array pcmArray = Napi::Uint8Array::New(env, pcmData.size());
        for (size_t i = 0; i < pcmData.size(); i++)
        {
            pcmArray[i] = pcmData[i];
        }
        res.Set("pcmData", pcmArray);

        Napi::Uint32Array capturedFramesArray = Napi::Uint32Array::New(env, capturedFramesList.size());
        for (size_t i = 0; i < capturedFramesList.size(); i++)
        {
            capturedFramesArray[i] = capturedFramesList[i];
        }
        res.Set("numberOfcapturedFramesArray", capturedFramesArray);

        Napi::Uint32Array packetSizeArray = Napi::Uint32Array::New(env, packetSizeList.size());
        for (size_t i = 0; i < packetSizeList.size(); i++)
        {
            packetSizeArray[i] = packetSizeList[i];
        }
        res.Set("packetSizeArray", packetSizeArray);

        WAVHeader wavHeader{};
        wavHeader.numChannels = mCaptureFormat->nChannels;
        wavHeader.sampleRate = mCaptureFormat->nSamplesPerSec;
        wavHeader.bitsPerSample = mCaptureFormat->wBitsPerSample;
        wavHeader.blockAlign = mCaptureFormat->nBlockAlign;
        wavHeader.byteRate = mCaptureFormat->nAvgBytesPerSec;
        wavHeader.dataChunkSize = static_cast<uint32_t>(pcmData.size());
        wavHeader.chunkSize = wavHeader.dataChunkSize + sizeof(WAVHeader) - 8; // 减去 'RIFF' 和 'WAVE' 的长度
        std::vector<uint8_t> wavData(sizeof(WAVHeader) + pcmData.size());
        std::memcpy(wavData.data(), &wavHeader, sizeof(WAVHeader));
        std::memcpy(wavData.data() + sizeof(WAVHeader), pcmData.data(), pcmData.size());
        Napi::Uint8Array wavArray = Napi::Uint8Array::New(env, wavData.size());
        for (size_t i = 0; i < wavData.size(); i++)
        {
            wavArray[i] = wavData[i];
        }
        res.Set("wavData", wavArray);

        return res;
    }
    else
    {
        return env.Null();
    }
}

class CaptureWorker : public Napi::AsyncWorker
{
public:
    CaptureWorker(Napi::Function &callback, int intervalMs)
        : Napi::AsyncWorker(callback), hasData(false), intervalMs(intervalMs) {}

    void Execute() override
    {
        UINT32 packetLength = 0;
        BYTE *Data = nullptr;
        HRESULT hr;
        DWORD dwCaptureFlags;
        UINT32 nNumFramesCaptured = 0;

        mAudioCaptureClient->GetNextPacketSize(&packetLength);
        if (packetLength > 0)
        {
            result.packetSizeList.push_back(packetLength);
            // UINT32 desiredFrameCount = 22050; // 约 500ms 的数据 (假设 44.1kHz 采样率)
            UINT32 desiredFrameCount = static_cast<UINT32>((intervalMs / 1000.0) * mCaptureFormat->nSamplesPerSec);
            result.totalFramesCaptured = 0;
            result.n = 0;

            while (result.totalFramesCaptured < desiredFrameCount)
            {
                mAudioCaptureClient->GetNextPacketSize(&packetLength);

                if (packetLength == 0)
                {
                    continue;
                }

                hr = mAudioCaptureClient->GetBuffer(&Data,
                                                    &nNumFramesCaptured,
                                                    &dwCaptureFlags,
                                                    NULL,
                                                    NULL);
                result.capturedFramesList.push_back(nNumFramesCaptured);
                result.totalFramesCaptured += nNumFramesCaptured;
                result.pcmData.insert(result.pcmData.end(), Data, Data + nNumFramesCaptured * mCaptureFormat->nBlockAlign);
                mAudioCaptureClient->ReleaseBuffer(nNumFramesCaptured);
                result.n++;
                result.packetSizeList.push_back(packetLength);

                if (result.totalFramesCaptured >= desiredFrameCount)
                {
                    break;
                }
            }

            if (g_AudioClient)
            {
                hr = g_AudioClient->GetCurrentPadding(&result.numPaddingFrames);
                result.currentPaddingFramesSuccess = SUCCEEDED(hr);
            }

            // 设置 WAV 头
            result.wavHeader.numChannels = mCaptureFormat->nChannels;
            result.wavHeader.sampleRate = mCaptureFormat->nSamplesPerSec;
            result.wavHeader.bitsPerSample = mCaptureFormat->wBitsPerSample;
            result.wavHeader.blockAlign = mCaptureFormat->nBlockAlign;
            result.wavHeader.byteRate = mCaptureFormat->nAvgBytesPerSec;
            result.wavHeader.dataChunkSize = static_cast<uint32_t>(result.pcmData.size());
            result.wavHeader.chunkSize = result.wavHeader.dataChunkSize + sizeof(WAVHeader) - 8;

            hasData = true;
        }
    }

    void OnOK() override
    {
        Napi::HandleScope scope(Env());

        if (hasData)
        {
            Napi::Object res = Napi::Object::New(Env());

            if (result.currentPaddingFramesSuccess)
            {
                res.Set("currentPaddingFrames", Napi::Number::New(Env(), static_cast<double>(result.numPaddingFrames)));
            }
            else
            {
                res.Set("currentPaddingFrames", Napi::Number::New(Env(), static_cast<double>(E_UNEXPECTED)));
            }

            res.Set("totalFramesCaptured", Napi::Number::From(Env(), result.totalFramesCaptured));
            res.Set("n", Napi::Number::From(Env(), result.n));

            Napi::Uint8Array pcmArray = Napi::Uint8Array::New(Env(), result.pcmData.size());
            std::copy(result.pcmData.begin(), result.pcmData.end(), pcmArray.Data());
            res.Set("pcmData", pcmArray);

            Napi::Uint32Array capturedFramesArray = Napi::Uint32Array::New(Env(), result.capturedFramesList.size());
            std::copy(result.capturedFramesList.begin(), result.capturedFramesList.end(), capturedFramesArray.Data());
            res.Set("numberOfcapturedFramesArray", capturedFramesArray);

            Napi::Uint32Array packetSizeArray = Napi::Uint32Array::New(Env(), result.packetSizeList.size());
            std::copy(result.packetSizeList.begin(), result.packetSizeList.end(), packetSizeArray.Data());
            res.Set("packetSizeArray", packetSizeArray);

            std::vector<uint8_t> wavData(sizeof(WAVHeader) + result.pcmData.size());
            std::memcpy(wavData.data(), &result.wavHeader, sizeof(WAVHeader));
            std::memcpy(wavData.data() + sizeof(WAVHeader), result.pcmData.data(), result.pcmData.size());
            Napi::Uint8Array wavArray = Napi::Uint8Array::New(Env(), wavData.size());
            std::copy(wavData.begin(), wavData.end(), wavArray.Data());
            res.Set("wavData", wavArray);

            Callback().Call({Env().Null(), res});
        }
        else
        {
            Callback().Call({Env().Null(), Env().Null()});
        }
    }

private:
    struct CaptureResult
    {
        std::vector<uint8_t> pcmData;
        std::vector<UINT32> capturedFramesList;
        std::vector<UINT32> packetSizeList;
        UINT32 totalFramesCaptured = 0;
        UINT8 n = 0;
        UINT32 numPaddingFrames = 0;
        bool currentPaddingFramesSuccess = false;
        WAVHeader wavHeader{};
    };

    CaptureResult result;
    bool hasData;
    int intervalMs;
};

Napi::Value capture_500_async(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();

    // default interval is 500ms
    int intervalMs = 500;

    if (info.Length() < 1 || info.Length() > 2)
    {
        Napi::TypeError::New(env, "Wrong number of arguments").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    if (info.Length() == 2)
    {
        if (!info[0].IsNumber() || !info[1].IsFunction())
        {
            Napi::TypeError::New(env, "Wrong arguments").ThrowAsJavaScriptException();
            return env.Undefined();
        }
        intervalMs = info[0].As<Napi::Number>().Int32Value();
        if (intervalMs <= 0)
        {
            Napi::RangeError::New(env, "Interval must be positive").ThrowAsJavaScriptException();
            return env.Undefined();
        }
    }
    else if (!info[0].IsFunction())
    {
        Napi::TypeError::New(env, "Function expected").ThrowAsJavaScriptException();
        return env.Undefined();
    }

    Napi::Function callback = info[info.Length() - 1].As<Napi::Function>();

    CaptureWorker *worker = new CaptureWorker(callback, intervalMs);
    worker->Queue();

    return env.Undefined();
}

HRESULT eventCallBack_SampleReady()
{

    return S_OK;
}

Napi::Object Init(Napi::Env env, Napi::Object exports)
{
    exports.Set(Napi::String::New(env, "getAudioProcessInfo"), Napi::Function::New(env, getAudioProcessInfo));
    exports.Set(Napi::String::New(env, "initializeCapture"), Napi::Function::New(env, initializeCapture));
    exports.Set(Napi::String::New(env, "getBuffer"), Napi::Function::New(env, getBuffer));
    exports.Set(Napi::String::New(env, "getAudioFormat"), Napi::Function::New(env, getAudioFormat));
    exports.Set(Napi::String::New(env, "getHalfSecWAV"), Napi::Function::New(env, getHalfSecWAV));
    exports.Set(Napi::String::New(env, "getWAVfromfile"), Napi::Function::New(env, getWAVfromfile));
    exports.Set(Napi::String::New(env, "initializeCLoopbackCapture"), Napi::Function::New(env, initializeCLoopbackCapture));
    exports.Set(Napi::String::New(env, "getActivateStatus"), Napi::Function::New(env, getActivateStatus));
    exports.Set(Napi::String::New(env, "getProcessCaptureFormat"), Napi::Function::New(env, getProcessCaptureFormat));
    exports.Set(Napi::String::New(env, "captureProcessAudio"), Napi::Function::New(env, captureProcessAudio));
    exports.Set(Napi::String::New(env, "whileCaptureProcessAudio"), Napi::Function::New(env, whileCaptureProcessAudio));
    exports.Set(Napi::String::New(env, "capture_500_async"), Napi::Function::New(env, capture_500_async));

    return exports;
}

NODE_API_MODULE(myncpp1, Init)
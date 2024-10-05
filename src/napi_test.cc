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

using namespace Microsoft::WRL;
BYTE *m_AudioData;
UINT32 m_AudioDataSize;
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
    Napi::Buffer<char> audioDataBuffer = Napi::Buffer<char>::Copy(env, (char*)m_AudioData, m_AudioDataSize);
    return audioDataBuffer;
}

/*
napi_value GenerateRandomNumber(napi_env env, napi_callback_info info)
{
    napi_status status;

    int randomNumber = rand() % 100;

    napi_value result;
    status = napi_create_int32(env, randomNumber, &result);

    return result;
}

napi_value init(napi_env env, napi_value exports)
{
    napi_property_descriptor des = {};
    des.utf8name = "generateRandomNumber";
    des.method = GenerateRandomNumber;
    des.attributes = napi_default_method;

    napi_define_properties(env, exports, 1, &des);

    return exports;
}
NAPI_MODULE(NODE_GYP_MODULE_NAME, init)

std::atomic<bool> a(true);

Napi::Value CallEmit(const Napi::CallbackInfo &info)
{
    char buff1[128];
    char buff2[128];
    int sensor1 = 0;
    int sensor2 = 0;
    Napi::Env env = info.Env();
    Napi::Function emit = info[0].As<Napi::Function>(); // 管理stream的函数其实是从js传来的

    emit.Call({Napi::String::New(env, "start")});

    int i = 0;
    while (a.load())
    {
        // 模拟从传感器收集数据的延迟
        std::this_thread::sleep_for(std::chrono::seconds(1));
        sprintf(buff1, "sensor1 data %d ... i=%d", ++sensor1, i);
        emit.Call({Napi::String::New(env, "sensor1"), Napi::String::New(env, buff1)});
        /*
        // 假设sensor2的数据报告频率是sensor1的一半
        if (i % 2)
        {
            sprintf(buff2, "sensor2 data %d ...", ++sensor2);
            emit.Call({Napi::String::New(env, "sensor2"), Napi::String::New(env, buff2)});
        }

        i++;
    }

    emit.Call({Napi::String::New(env, "end")});
    return Napi::String::New(env, "OK");
}

Napi::Value StopCollection(const Napi::CallbackInfo &info)
{
    a.store(false);
    return Napi::String::New(info.Env(), "OK");
}
*/

const size_t ArrayLength = 100;
int intArray[ArrayLength];

std::atomic<bool> stopThread(false);
void UpdateIntArrayAsync(Napi::Env env)
{
    while (!stopThread)
    {
        // 生成随机数填充数组
        for (int i = 0; i < ArrayLength; i++)
        {
            intArray[i] = rand() % 100; // 生成 0 到 99 之间的随机整数
        }
        // 等待 1 秒
        std::this_thread::sleep_for(std::chrono::seconds(1));
    }
}
Napi::Value StartUpdatingArray(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();

    // 在单独的线程中启动定时刷新任务
    std::thread(UpdateIntArrayAsync, env).detach();

    return env.Null();
}

Napi::Value GenerateArray(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    for (int i = 0; i < ArrayLength; i++)
    {
        intArray[i] = rand() % 100; // 生成 0 到 99 之间的随机整数
    }
    return env.Null();
}
Napi::Value GetArrayBuffer(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();

    Napi::Buffer<int> arrayBuffer = Napi::Buffer<int>::Copy(env, intArray, ArrayLength); // sizeof(intArray) / sizeof(intArray[0])
    return arrayBuffer;
}
Napi::Value GetOriginArray(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();

    Napi::Array resultArray = Napi::Array::New(env, ArrayLength);
    for (size_t i = 0; i < ArrayLength; i++)
    {
        resultArray[i] = Napi::Number::New(env, intArray[i]);
    }
    return resultArray;
}

int gv;
Napi::Value SetGV(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    int value = info[0].As<Napi::Number>().Int32Value();
    gv = value;
    return Napi::String::New(env, "global value created");
}
Napi::Value GetGV(const Napi::CallbackInfo &info)
{
    Napi::Env env = info.Env();
    return Napi::Value::From(env, gv);
}

Napi::Object Init(Napi::Env env, Napi::Object exports)
{
    exports.Set(Napi::String::New(env, "startCapture"), Napi::Function::New(env, StartCaputure));
    // exports.Set(Napi::String::New(env, "callEmit"), Napi::Function::New(env, CallEmit));
    // exports.Set(Napi::String::New(env, "stopCollection"), Napi::Function::New(env, StopCollection));

    exports.Set(Napi::String::New(env, "generateArray"), Napi::Function::New(env, GenerateArray));
    exports.Set(Napi::String::New(env, "startUpdatingArray"), Napi::Function::New(env, StartUpdatingArray));
    exports.Set(Napi::String::New(env, "getArrayBuffer"), Napi::Function::New(env, GetArrayBuffer));
    exports.Set(Napi::String::New(env, "getOriginArray"), Napi::Function::New(env, GetOriginArray));

    exports.Set(Napi::String::New(env, "setGV"), Napi::Function::New(env, SetGV));
    exports.Set(Napi::String::New(env, "getGV"), Napi::Function::New(env, GetGV));

    return exports;
}

NODE_API_MODULE(myncpp1, Init)

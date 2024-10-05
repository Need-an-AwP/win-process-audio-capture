#include <iostream>
#include <windows.h>
#include <mmdeviceapi.h>
#include <audiopolicy.h>
#include <vector>
#include <TlHelp32.h>
#include <shlobj.h>
#include <wchar.h>
#include <audioclientactivationparams.h>

#include <wrl\implements.h>
#include <wil\com.h>
#include <wil\result.h>

#include <node.h>
#include <node_object_wrap.h>

namespace demo
{
    void getAudioProcess(const v8::FunctionCallbackInfo<v8::Value> &args)
    {
        v8::Isolate *isolate = args.GetIsolate();

        HRESULT hr;
        IMMDeviceEnumerator *pEnumerator = NULL;
        IMMDevice *pDevice = NULL;
        IAudioSessionManager2 *pSessionManager = NULL;
        IAudioSessionEnumerator *pSessionEnumerator = NULL;

        // Initialize COM
        hr = CoInitialize(NULL);
        if (FAILED(hr))
        {
            v8::Local<v8::Value> exception = v8::Exception::Error(
                v8::String::NewFromUtf8(isolate, "COM initialization failed").ToLocalChecked());
            isolate->ThrowException(exception);
        }

        // Create a multimedia device enumerator
        hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void **)&pEnumerator);
        if (FAILED(hr))
        {
            v8::Local<v8::Value> exception = v8::Exception::Error(
                v8::String::NewFromUtf8(isolate, "Failed to create device enumerator").ToLocalChecked());
            isolate->ThrowException(exception);
            CoUninitialize();
            return;
        }

        // Get the default audio endpoint
        hr = pEnumerator->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice);
        if (FAILED(hr))
        {
            v8::Local<v8::Value> exception = v8::Exception::Error(
                v8::String::NewFromUtf8(isolate, "Failed to get default audio endpoint").ToLocalChecked());
            isolate->ThrowException(exception);
            pEnumerator->Release();
            CoUninitialize();
            return;
        }

        // Get the audio session manager
        hr = pDevice->Activate(__uuidof(IAudioSessionManager2), CLSCTX_ALL, NULL, (void **)&pSessionManager);
        if (FAILED(hr))
        {
            v8::Local<v8::Value> exception = v8::Exception::Error(
                v8::String::NewFromUtf8(isolate, "Failed to get audio session manager").ToLocalChecked());
            isolate->ThrowException(exception);
            pDevice->Release();
            pEnumerator->Release();
            CoUninitialize();
            return;
        }

        // Get the audio session enumerator
        hr = pSessionManager->GetSessionEnumerator(&pSessionEnumerator);
        if (FAILED(hr))
        {
            v8::Local<v8::Value> exception = v8::Exception::Error(
                v8::String::NewFromUtf8(isolate, "Failed to get session enumerator").ToLocalChecked());
            isolate->ThrowException(exception);
            pSessionManager->Release();
            pDevice->Release();
            pEnumerator->Release();
            CoUninitialize();
            return;
        }

        // Get the count of audio sessions
        int sessionCount = 0;
        hr = pSessionEnumerator->GetCount(&sessionCount);
        if (FAILED(hr))
        {
            v8::Local<v8::Value> exception = v8::Exception::Error(
                v8::String::NewFromUtf8(isolate, "Failed to get session count").ToLocalChecked());
            isolate->ThrowException(exception);
            pSessionEnumerator->Release();
            pSessionManager->Release();
            pDevice->Release();
            pEnumerator->Release();
            CoUninitialize();
            return;
        }

        // Create a process snapshot
        HANDLE hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
        if (hSnapshot == INVALID_HANDLE_VALUE)
        {
            v8::Local<v8::Value> exception = v8::Exception::Error(
                v8::String::NewFromUtf8(isolate, "Failed to create process snapshot").ToLocalChecked());
            isolate->ThrowException(exception);
            pSessionEnumerator->Release();
            pSessionManager->Release();
            pDevice->Release();
            pEnumerator->Release();
            CoUninitialize();
            return;
        }

        v8::Local<v8::Array> resultArray = v8::Array::New(isolate);

        // Iterate through each audio session
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
                                    v8::Local<v8::Context> context = isolate->GetCurrentContext();
                                    v8::Local<v8::Object> sessionInfo = v8::Object::New(isolate);
                                    sessionInfo->Set(
                                        context,
                                        v8::String::NewFromUtf8(isolate, "processId").ToLocalChecked(),
                                        v8::Integer::New(isolate, processId));
                                    sessionInfo->Set(
                                        context,
                                        v8::String::NewFromUtf8(isolate, "processName").ToLocalChecked(),
                                        v8::String::NewFromUtf8(isolate, pe32.szExeFile).ToLocalChecked());
                                    resultArray->Set(context, index++, sessionInfo);
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

        // Close process snapshot handle
        CloseHandle(hSnapshot);

        // Release all COM objects
        pSessionEnumerator->Release();
        pSessionManager->Release();
        pDevice->Release();
        pEnumerator->Release();
        CoUninitialize();

        args.GetReturnValue().Set(resultArray);
    }

    void capture(const v8::FunctionCallbackInfo<v8::Value> &args)
    {
        v8::Isolate *isolate = args.GetIsolate();
        DWORD processId = 14620;
        AUDIOCLIENT_ACTIVATION_PARAMS audioclientActivationParams = {};
        audioclientActivationParams.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;
        audioclientActivationParams.ProcessLoopbackParams.ProcessLoopbackMode = PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE;
        audioclientActivationParams.ProcessLoopbackParams.TargetProcessId = processId;

        PROPVARIANT activateParams = {};
        activateParams.vt = VT_BLOB;
        activateParams.blob.cbSize = sizeof(audioclientActivationParams);
        activateParams.blob.pBlobData = (BYTE *)&audioclientActivationParams;

        wil::com_ptr_nothrow<IActivateAudioInterfaceAsyncOperation> asyncOp;
        HRESULT hr = ActivateAudioInterfaceAsync(VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK, __uuidof(IAudioClient), &activateParams, this, &asyncOp);

        v8::Local<v8::String> str = v8::String::NewFromUtf8(isolate, hr).ToLocalChecked();
        args.GetReturnValue().Set(str);
    }

    void Initialize(v8::Local<v8::Object> exports)
    {
        NODE_SET_METHOD(exports, "capture", capture);
        NODE_SET_METHOD(exports, "getAudioProcess", getAudioProcess);
    }

#undef NODE_MODULE_VERSION
#define NODE_MODULE_VERSION 121
    NODE_MODULE(NODE_GYP_MODULE_NAME, Initialize)
}
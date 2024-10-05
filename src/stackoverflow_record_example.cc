#include <Windows.h>
#include <mmsystem.h>
#include <mmdeviceapi.h>
#include <audioclient.h>
#include <time.h>
#include <iostream>

#include <propsys.h>
#include <propkey.h>
#include <codecvt>
#include <locale>
#include <comdef.h>
#include <string>
/*
#include <mfapi.h>
#include <audioclientactivationparams.h>
#include <wrl\client.h>
#include <wrl\implements.h>
#include <wil\com.h>
#include <wil\result.h>
#include "Common.h"
#include <audiopolicy.h>
#include <TlHelp32.h>
#include <shlobj.h>
#include <wchar.h>
#include <initguid.h>
#include <guiddef.h>
*/
using namespace std;
// using namespace Microsoft::WRL;

#pragma comment(lib, "Winmm.lib")

WCHAR fileName[] = L"loopback-capture.wav";
BOOL bDone = FALSE;
HMMIO hFile = NULL;
DWORD processId;
bool includeProcessTree;
// REFERENCE_TIME time units per second and per millisecond
#define REFTIMES_PER_SEC 10000000
#define REFTIMES_PER_MILLISEC 10000

#define EXIT_ON_ERROR(hres)                                \
    if (FAILED(hres))                                      \
    {                                                      \
        _com_error err(hres);                              \
        LPCTSTR errMsg;                                    \
        errMsg = err.ErrorMessage();                       \
        wcout << "exit with hr value: " << errMsg << endl; \
        goto Exit;                                         \
    }
#define SAFE_RELEASE(punk) \
    if ((punk) != NULL)    \
    {                      \
        (punk)->Release(); \
        (punk) = NULL;     \
    }

const CLSID CLSID_MMDeviceEnumerator = __uuidof(MMDeviceEnumerator);
const IID IID_IMMDeviceEnumerator = __uuidof(IMMDeviceEnumerator);
const IID IID_IAudioClient = __uuidof(IAudioClient);
const IID IID_IAudioCaptureClient = __uuidof(IAudioCaptureClient);

class MyAudioSink
{
public:
    HRESULT CopyData(BYTE *pData, UINT32 NumFrames, BOOL *pDone, WAVEFORMATEX *pwfx, HMMIO hFile);
};

HRESULT WriteWaveHeader(HMMIO hFile, LPCWAVEFORMATEX pwfx, MMCKINFO *pckRIFF, MMCKINFO *pckData);
HRESULT FinishWaveFile(HMMIO hFile, MMCKINFO *pckRIFF, MMCKINFO *pckData);
HRESULT RecordAudioStream(MyAudioSink *pMySink);

////////
/*
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
                                                   AUDCLNT_STREAMFLAGS_LOOPBACK,
                                                   200000,
                                                   AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM,
                                                   &m_CaptureFormat,
                                                   nullptr));
        RETURN_IF_FAILED(m_AudioClient->GetBufferSize(&m_BufferFrames));
        RETURN_IF_FAILED(m_AudioClient->GetService(IID_PPV_ARGS(&m_AudioCaptureClient)));
        return S_OK;
    }
};
*/
std::wstring HResultToString(HRESULT hr)
{
    wchar_t *pszMessage = nullptr;
    std::wstring result;

    DWORD dwFlags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
    DWORD dwLanguageId = MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT);
    if (FormatMessageW(dwFlags, nullptr, hr, dwLanguageId, (LPWSTR)&pszMessage, 0, nullptr) > 0)
    {
        result = pszMessage;
        LocalFree(pszMessage);
    }
    else
    {
        wchar_t buffer[32];
        swprintf_s(buffer, L"0x%08X", hr);
        result = buffer;
    }

    return result;
}
////////

int main()
{
    ////////
    /*includeProcessTree = true;
    cout << "enter a process id:" << "\n"
         << endl;
    std::string userInput;
    cin >> userInput;
    processId = std::strtoul(userInput.c_str(), nullptr, 10);
    // cout<<"processId(DWORD): "<<std::to_string(processId)<<endl;
    AUDIOCLIENT_ACTIVATION_PARAMS audioclientActivationParams = {};
    audioclientActivationParams.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;
    audioclientActivationParams.ProcessLoopbackParams.ProcessLoopbackMode = includeProcessTree ? PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE : PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE;
    audioclientActivationParams.ProcessLoopbackParams.TargetProcessId = processId;

    PROPVARIANT activateParams = {};
    activateParams.vt = VT_BLOB;
    activateParams.blob.cbSize = sizeof(audioclientActivationParams);
    activateParams.blob.pBlobData = (BYTE *)&audioclientActivationParams;

    wil::com_ptr_nothrow<IActivateAudioInterfaceAsyncOperation> asyncOp;
    Microsoft::WRL::ComPtr<CLoopbackCapture> completionHandler = Make<CLoopbackCapture>();
    HRESULT hr = ActivateAudioInterfaceAsync(VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK,
                                             __uuidof(IAudioClient),
                                             &activateParams,
                                             completionHandler.Get(),
                                             &asyncOp);
    RETURN_IF_FAILED(hr);*/
    ////////
    clock();

    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);

    // Create file
    MMIOINFO mi = {0};
    hFile = mmioOpen(
        // some flags cause mmioOpen write to this buffer
        // but not any that we're using
        (LPWSTR)fileName,
        &mi,
        MMIO_WRITE | MMIO_CREATE);

    if (NULL == hFile)
    {
        wprintf(L"mmioOpen(\"%ls\", ...) failed. wErrorRet == %u", fileName, GetLastError());
        return E_FAIL;
    }

    MyAudioSink AudioSink;
    RecordAudioStream(&AudioSink);

    mmioClose(hFile, 0);

    CoUninitialize();
    return 0;
}

HRESULT MyAudioSink::CopyData(BYTE *pData, UINT32 NumFrames, BOOL *pDone, WAVEFORMATEX *pwfx, HMMIO hFile)
{
    HRESULT hr = S_OK;

    if (0 == NumFrames)
    {
        wprintf(L"IAudioCaptureClient::GetBuffer said to read 0 frames\n");
        return E_UNEXPECTED;
    }

    LONG lBytesToWrite = NumFrames * pwfx->nBlockAlign;
#pragma prefast(suppress : __WARNING_INCORRECT_ANNOTATION, "IAudioCaptureClient::GetBuffer SAL annotation implies a 1-byte buffer")
    LONG lBytesWritten = mmioWrite(hFile, reinterpret_cast<PCHAR>(pData), lBytesToWrite);
    if (lBytesToWrite != lBytesWritten)
    {
        wprintf(L"mmioWrite wrote %u bytes : expected %u bytes", lBytesWritten, lBytesToWrite);
        return E_UNEXPECTED;
    }

    static int CallCount = 0;
    cout << "CallCount = " << CallCount++ << "NumFrames: " << NumFrames << endl;

    if (clock() > 5 * CLOCKS_PER_SEC) // Record 10 seconds. From the first time call clock() at the beginning of the main().
        *pDone = true;

    return S_OK;
}

HRESULT RecordAudioStream(MyAudioSink *pMySink)
{
    HRESULT hr;
    REFERENCE_TIME hnsRequestedDuration = REFTIMES_PER_SEC;
    REFERENCE_TIME hnsActualDuration;
    UINT32 bufferFrameCount;
    UINT32 numFramesAvailable;
    IMMDeviceEnumerator *pEnumerator = NULL;
    IMMDevice *pDefaultDevice = NULL;
    IAudioClient *pAudioClient = NULL;
    IAudioCaptureClient *pCaptureClient = NULL;
    WAVEFORMATEX *pwfx = new WAVEFORMATEX;
    WAVEFORMATEX *pwfx1 = new WAVEFORMATEX;
    WAVEFORMATEX *correctFormat = new WAVEFORMATEX;
    // WAVEFORMATEXTENSIBLE *pwfx1 = new WAVEFORMATEXTENSIBLE;
    std::wstring result;
    UINT32 packetLength = 0;

    BYTE *pData;
    DWORD flags;

    MMCKINFO ckRIFF = {0};
    MMCKINFO ckData = {0};

    hr = CoCreateInstance(
        CLSID_MMDeviceEnumerator, NULL,
        CLSCTX_ALL, IID_IMMDeviceEnumerator,
        (void **)&pEnumerator);
    EXIT_ON_ERROR(hr)

    hr = pEnumerator->GetDefaultAudioEndpoint(eRender, //eCapture,
                                              eConsole,
                                              &pDefaultDevice);
    EXIT_ON_ERROR(hr)
    // 使用eCapture（输入设备）获取默认设备时不能使用AUDCLNT_STREAMFLAGS_LOOPBACK

    ///////
    IMMDevice *pSelectedDevice;
    IMMDeviceCollection *pDevices;
    hr = pEnumerator->EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE, &pDevices); // dataflow is set as ecapture
    EXIT_ON_ERROR(hr)
    UINT deviceCount;
    hr = pDevices->GetCount(&deviceCount);
    EXIT_ON_ERROR(hr)

    bool forLoopBreak;
    forLoopBreak = false;
    for (UINT i = 0; i < deviceCount; i++)
    {
        IMMDevice *pDevice = NULL;
        hr = pDevices->Item(i, &pDevice);
        EXIT_ON_ERROR(hr)

        IMMEndpoint *pEndpoint;
        hr = pDevice->QueryInterface(__uuidof(IMMEndpoint), (void **)&pEndpoint);
        EDataFlow dataFlow;
        hr = pEndpoint->GetDataFlow(&dataFlow);
        EXIT_ON_ERROR(hr);
        if (dataFlow == eRender)
        {
            // std::cout << "output device" << std::endl;
        }
        else if (dataFlow == eCapture)
        {
            // std::cout << "input device" << std::endl;
        }

        IPropertyStore *pProps;
        hr = pDevice->OpenPropertyStore(STGM_READ, &pProps);
        EXIT_ON_ERROR(hr)

        DWORD propertyCount;
        hr = pProps->GetCount(&propertyCount);
        EXIT_ON_ERROR(hr)

        for (DWORD j = 0; j < propertyCount; j++)
        {
            PROPERTYKEY key;
            hr = pProps->GetAt(j, &key);
            EXIT_ON_ERROR(hr)
            PROPVARIANT varName;
            PropVariantInit(&varName);
            hr = pProps->GetValue(key, &varName);
            EXIT_ON_ERROR(hr)
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
        }
        if (forLoopBreak)
        {
            break;
        }
    }
    IPropertyStore *pProps;
    hr = pSelectedDevice->OpenPropertyStore(STGM_READ, &pProps);
    DWORD propertyCount;
    hr = pProps->GetCount(&propertyCount);
    //cout << "pSelectedDevice info:\n";
    for (UINT i = 0; i < propertyCount; i++)
    {
        PROPERTYKEY key;
        hr = pProps->GetAt(i, &key);
        PROPVARIANT varName;
        PropVariantInit(&varName);
        hr = pProps->GetValue(key, &varName);
        if (varName.vt == VT_LPWSTR)
        {
            std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
            std::string utf8Str = converter.to_bytes(varName.pwszVal);
            cout << utf8Str << "\n";
        }
    }
    cout << "\n\n";
    hr = pDefaultDevice->OpenPropertyStore(STGM_READ, &pProps);
    hr = pProps->GetCount(&propertyCount);
    cout << "pDefaultDevice info:\n";
    for (UINT i = 0; i < propertyCount; i++)
    {
        PROPERTYKEY key;
        hr = pProps->GetAt(i, &key);
        PROPVARIANT varName;
        PropVariantInit(&varName);
        hr = pProps->GetValue(key, &varName);
        if (varName.vt == VT_LPWSTR)
        {
            std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
            std::string utf8Str = converter.to_bytes(varName.pwszVal);
            cout << utf8Str << "\n";
        }
    }
    cout << "\n\n";

    ///////

    // hr = pSelectedDevice->Activate( // use Realtek(R) Audio micphone as device
    hr = pDefaultDevice->Activate(
        IID_IAudioClient, CLSCTX_ALL,
        NULL, (void **)&pAudioClient);
    EXIT_ON_ERROR(hr)

    hr = pAudioClient->GetMixFormat(&pwfx);
    EXIT_ON_ERROR(hr)

    /*
    REFERENCE_TIME defaultPeriodInFrames;
    hr = pAudioClient->GetDevicePeriod(&hnsRequestedDuration, &defaultPeriodInFrames);
    EXIT_ON_ERROR(hr);
    */
    pwfx1->wFormatTag = WAVE_FORMAT_PCM;//WAVE_FORMAT_EXTENSIBLE
    pwfx1->nChannels = 2;
    pwfx1->nSamplesPerSec = 44100; // 48000; 
    pwfx1->wBitsPerSample = 16;    // 32;    
    pwfx1->nBlockAlign = pwfx1->nChannels * pwfx1->wBitsPerSample / 8;
    pwfx1->nAvgBytesPerSec = pwfx1->nSamplesPerSec * pwfx1->nBlockAlign;
    pwfx1->cbSize = 0; // 22; 

    hr = pAudioClient->IsFormatSupported(AUDCLNT_SHAREMODE_SHARED,
                                         pwfx1,
                                         &correctFormat);

    std::cout << "Hexadecimal value: " << std::hex << std::showbase << hr << std::endl;
    //EXIT_ON_ERROR(hr)

    hr = pAudioClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        AUDCLNT_STREAMFLAGS_LOOPBACK,//AUDCLNT_STREAMFLAGS_NOPERSIST,  //0,
        hnsRequestedDuration,
        0, // defaultPeriodInFrames
        pwfx,
        nullptr);
    std::cout << "Hexadecimal value: " << std::hex << std::showbase << hr << std::endl;
    EXIT_ON_ERROR(hr)

    // Get the size of the allocated buffer.
    hr = pAudioClient->GetBufferSize(&bufferFrameCount);
    EXIT_ON_ERROR(hr)

    hr = pAudioClient->GetService(
        IID_IAudioCaptureClient,
        (void **)&pCaptureClient);
    EXIT_ON_ERROR(hr)

    hr = WriteWaveHeader((HMMIO)hFile, pwfx, &ckRIFF, &ckData);
    if (FAILED(hr))
    {
        // WriteWaveHeader does its own logging
        return hr;
    }

    // Calculate the actual duration of the allocated buffer.
    hnsActualDuration = (double)REFTIMES_PER_SEC *
                        bufferFrameCount / pwfx->nSamplesPerSec;

    hr = pAudioClient->Start(); // Start recording.
    EXIT_ON_ERROR(hr)

    // Each loop fills about half of the shared buffer.
    while (bDone == FALSE)
    {
        // Sleep for half the buffer duration.
        Sleep(hnsActualDuration / REFTIMES_PER_MILLISEC / 2);
        cout << "sleep time:" << hnsActualDuration / REFTIMES_PER_MILLISEC / 2 << endl;

        hr = pCaptureClient->GetNextPacketSize(&packetLength);
        EXIT_ON_ERROR(hr)

        while (packetLength != 0)
        {
            // Get the available data in the shared buffer.
            hr = pCaptureClient->GetBuffer(
                &pData,
                &numFramesAvailable,
                &flags, NULL, NULL);
            EXIT_ON_ERROR(hr)

            if (flags & AUDCLNT_BUFFERFLAGS_SILENT)
            {
                pData = NULL; // Tell CopyData to write silence.
            }

            // Copy the available capture data to the audio sink.
            hr = pMySink->CopyData(
                pData, numFramesAvailable, &bDone, pwfx, (HMMIO)hFile);
            EXIT_ON_ERROR(hr)

            hr = pCaptureClient->ReleaseBuffer(numFramesAvailable);
            EXIT_ON_ERROR(hr)

            hr = pCaptureClient->GetNextPacketSize(&packetLength);
            EXIT_ON_ERROR(hr)
        }
    }

    hr = pAudioClient->Stop(); // Stop recording.
    EXIT_ON_ERROR(hr)

    hr = FinishWaveFile((HMMIO)hFile, &ckData, &ckRIFF);
    if (FAILED(hr))
    {
        // FinishWaveFile does it's own logging
        return hr;
    }

Exit:
    CoTaskMemFree(pwfx);
    CoTaskMemFree(pwfx1);
    SAFE_RELEASE(pEnumerator);
    SAFE_RELEASE(pDefaultDevice);
    SAFE_RELEASE(pAudioClient);
    SAFE_RELEASE(pCaptureClient);
    SAFE_RELEASE(pSelectedDevice);

    return hr;
}

HRESULT WriteWaveHeader(HMMIO hFile, LPCWAVEFORMATEX pwfx, MMCKINFO *pckRIFF, MMCKINFO *pckData)
{
    MMRESULT result;

    // make a RIFF/WAVE chunk
    pckRIFF->ckid = MAKEFOURCC('R', 'I', 'F', 'F');
    pckRIFF->fccType = MAKEFOURCC('W', 'A', 'V', 'E');

    result = mmioCreateChunk(hFile, pckRIFF, MMIO_CREATERIFF);
    if (MMSYSERR_NOERROR != result)
    {
        wprintf(L"mmioCreateChunk(\"RIFF/WAVE\") failed: MMRESULT = 0x%08x", result);
        return E_FAIL;
    }

    // make a 'fmt ' chunk (within the RIFF/WAVE chunk)
    MMCKINFO chunk;
    chunk.ckid = MAKEFOURCC('f', 'm', 't', ' ');
    result = mmioCreateChunk(hFile, &chunk, 0);
    if (MMSYSERR_NOERROR != result)
    {
        wprintf(L"mmioCreateChunk(\"fmt \") failed: MMRESULT = 0x%08x", result);
        return E_FAIL;
    }

    // write the WAVEFORMATEX data to it
    LONG lBytesInWfx = sizeof(WAVEFORMATEX) + pwfx->cbSize;
    LONG lBytesWritten =
        mmioWrite(
            hFile,
            reinterpret_cast<PCHAR>(const_cast<LPWAVEFORMATEX>(pwfx)),
            lBytesInWfx);
    if (lBytesWritten != lBytesInWfx)
    {
        wprintf(L"mmioWrite(fmt data) wrote %u bytes; expected %u bytes", lBytesWritten, lBytesInWfx);
        return E_FAIL;
    }

    // ascend from the 'fmt ' chunk
    result = mmioAscend(hFile, &chunk, 0);
    if (MMSYSERR_NOERROR != result)
    {
        wprintf(L"mmioAscend(\"fmt \" failed: MMRESULT = 0x%08x", result);
        return E_FAIL;
    }

    // make a 'fact' chunk whose data is (DWORD)0
    chunk.ckid = MAKEFOURCC('f', 'a', 'c', 't');
    result = mmioCreateChunk(hFile, &chunk, 0);
    if (MMSYSERR_NOERROR != result)
    {
        wprintf(L"mmioCreateChunk(\"fmt \") failed: MMRESULT = 0x%08x", result);
        return E_FAIL;
    }

    // write (DWORD)0 to it
    // this is cleaned up later
    DWORD frames = 0;
    lBytesWritten = mmioWrite(hFile, reinterpret_cast<PCHAR>(&frames), sizeof(frames));
    if (lBytesWritten != sizeof(frames))
    {
        wprintf(L"mmioWrite(fact data) wrote %u bytes; expected %u bytes", lBytesWritten, (UINT32)sizeof(frames));
        return E_FAIL;
    }

    // ascend from the 'fact' chunk
    result = mmioAscend(hFile, &chunk, 0);
    if (MMSYSERR_NOERROR != result)
    {
        wprintf(L"mmioAscend(\"fact\" failed: MMRESULT = 0x%08x", result);
        return E_FAIL;
    }

    // make a 'data' chunk and leave the data pointer there
    pckData->ckid = MAKEFOURCC('d', 'a', 't', 'a');
    result = mmioCreateChunk(hFile, pckData, 0);
    if (MMSYSERR_NOERROR != result)
    {
        wprintf(L"mmioCreateChunk(\"data\") failed: MMRESULT = 0x%08x", result);
        return E_FAIL;
    }

    return S_OK;
}

HRESULT FinishWaveFile(HMMIO hFile, MMCKINFO *pckRIFF, MMCKINFO *pckData)
{
    MMRESULT result;

    result = mmioAscend(hFile, pckData, 0);
    if (MMSYSERR_NOERROR != result)
    {
        wprintf(L"mmioAscend(\"data\" failed: MMRESULT = 0x%08x", result);
        return E_FAIL;
    }

    result = mmioAscend(hFile, pckRIFF, 0);
    if (MMSYSERR_NOERROR != result)
    {
        wprintf(L"mmioAscend(\"RIFF/WAVE\" failed: MMRESULT = 0x%08x", result);
        return E_FAIL;
    }

    return S_OK;
}
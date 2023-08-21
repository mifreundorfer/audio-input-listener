#include <windows.h>
#include <audioclient.h>
#include <AudioSessionTypes.h>
#include <audioenginebaseapo.h>
#include <wrl/client.h>
#include <Functiondiscoverykeys_devpkey.h>
#include <avrt.h>

#include <math.h>
#include <queue>

#pragma comment(lib, "Avrt.lib")

using namespace Microsoft::WRL;

ComPtr<IMMDeviceEnumerator> deviceEnumerator;
ComPtr<IMMDevice> inputDevice;
ComPtr<IMMDevice> outputDevice;
ComPtr<IAudioClient> inputClient;
ComPtr<IAudioClient> outputClient;
ComPtr<IAudioCaptureClient> captureClient;
ComPtr<IAudioRenderClient> renderClient;

HANDLE inputEventHandle;
HANDLE outputEventHandle;

UINT32 bufferFrameCount;

DWORD taskIndex = 0;
HANDLE taskHandle;

bool FindAudioDevice(ComPtr<IMMDeviceEnumerator> enumerator, EDataFlow dataFlow, LPCWSTR deviceName, IMMDevice** foundDevice)
{
    HRESULT hr;

    ComPtr<IMMDeviceCollection> collection;
    hr = deviceEnumerator->EnumAudioEndpoints(dataFlow, DEVICE_STATE_ACTIVE, collection.GetAddressOf());
    if (FAILED(hr))
    {
        return false;
    }

    UINT numDevices;
    hr = collection->GetCount(&numDevices);
    if (FAILED(hr))
    {
        return false;
    }

    for (UINT i = 0; i < numDevices; i++)
    {
        ComPtr<IMMDevice> device;
        hr = collection->Item(i, device.GetAddressOf());
        if (FAILED(hr))
        {
            continue;
        }

        ComPtr<IPropertyStore> properties;
        hr = device->OpenPropertyStore(STGM_READ, properties.GetAddressOf());
        if (FAILED(hr))
        {
            continue;
        }

        PROPVARIANT varName;
        PropVariantInit(&varName);

        hr = properties->GetValue(PKEY_Device_FriendlyName, &varName);
        if (FAILED(hr))
        {
            continue;
        }

        if (varName.vt == VT_LPWSTR)
        {
            int diff = lstrcmpW(varName.pwszVal, deviceName);
            if (diff == 0)
            {
                hr = device.CopyTo(foundDevice);
                return SUCCEEDED(hr);
            }
        }
    }

    return false;
}

HRESULT InitializeAudioClients() {
    HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    if (FAILED(hr)) {
        // Handle error
        return hr;
    }

    taskHandle = AvSetMmThreadCharacteristicsW(L"Pro Audio", &taskIndex);
    if (taskHandle == NULL)
    {
        return E_FAIL;
    }

    hr = CoCreateInstance(
        __uuidof(MMDeviceEnumerator),
        nullptr,
        CLSCTX_ALL,
        __uuidof(IMMDeviceEnumerator),
        reinterpret_cast<void**>(deviceEnumerator.GetAddressOf())
    );
    if (FAILED(hr)) {
        // Handle error
        CoUninitialize();
        return hr;
    }

    if (!FindAudioDevice(deviceEnumerator, EDataFlow::eCapture, L"Broadcast Stream Mix (TC-Helicon GoXLR Mini)", inputDevice.GetAddressOf()))
    {
        return -1;
    }

    if (!FindAudioDevice(deviceEnumerator, EDataFlow::eRender, L"TX-SR252 (NVIDIA High Definition Audio)", outputDevice.GetAddressOf()))
    {
        return -1;
    }

    hr = inputDevice->Activate(
        __uuidof(IAudioClient),
        CLSCTX_ALL,
        nullptr,
        reinterpret_cast<void**>(inputClient.GetAddressOf())
    );
    if (FAILED(hr)) {
        // Handle error
        inputDevice.Reset();
        deviceEnumerator.Reset();
        CoUninitialize();
        return hr;
    }

    hr = outputDevice->Activate(
        __uuidof(IAudioClient),
        CLSCTX_ALL,
        nullptr,
        reinterpret_cast<void**>(outputClient.GetAddressOf())
    );
    if (FAILED(hr)) {
        // Handle error
        inputClient.Reset();
        inputDevice.Reset();
        outputDevice.Reset();
        deviceEnumerator.Reset();
        CoUninitialize();
        return hr;
    }

    WAVEFORMATEX format = {};
    format.wFormatTag = WAVE_FORMAT_PCM;
    format.nChannels = 2;
    format.wBitsPerSample = 16;
    format.nSamplesPerSec = 48000;
    format.nBlockAlign = format.wBitsPerSample / 8 * format.nChannels;
    format.nAvgBytesPerSec = format.nBlockAlign * format.nSamplesPerSec;
    format.cbSize = 0;

    REFERENCE_TIME inputDuration;
    hr = inputClient->GetDevicePeriod(NULL, &inputDuration);
    printf("Minimum period for input device %d\n", inputDuration);

    REFERENCE_TIME outputDuration;
    hr = outputClient->GetDevicePeriod(NULL, &outputDuration);
    printf("Minimum period for output device %d\n", outputDuration);

    hr = inputClient->Initialize(
        AUDCLNT_SHAREMODE_EXCLUSIVE,
        AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
        inputDuration,
        inputDuration,
        &format,
        nullptr
    );
    if (FAILED(hr)) {
        // Handle error
        outputClient.Reset();
        inputClient.Reset();
        inputDevice.Reset();
        outputDevice.Reset();
        deviceEnumerator.Reset();
        CoUninitialize();
        return hr;
    }

    hr = inputClient->GetService(
        __uuidof(IAudioCaptureClient),
        reinterpret_cast<void**>(captureClient.GetAddressOf())
    );
    if (FAILED(hr)) {
        // Handle error
        outputClient.Reset();
        inputClient.Reset();
        inputDevice.Reset();
        outputDevice.Reset();
        deviceEnumerator.Reset();
        CoUninitialize();
        return hr;
    }

    hr = outputClient->Initialize(
        AUDCLNT_SHAREMODE_EXCLUSIVE,
        AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
        outputDuration,
        outputDuration,
        &format,  // Use the same format as input
        nullptr
    );

    if (hr == AUDCLNT_E_BUFFER_SIZE_NOT_ALIGNED)
    {
        // Align the buffer if needed, see IAudioClient::Initialize() documentation
        UINT32 nFrames = 0;
        hr = outputClient->GetBufferSize(&nFrames);

        outputDuration = (REFERENCE_TIME)((double)10000000 / format.nSamplesPerSec * nFrames + 0.5);
        printf("Aligning buffer size for output device %d\n", outputDuration);

        hr = outputClient->Initialize(
            AUDCLNT_SHAREMODE_EXCLUSIVE,
            AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
            outputDuration,
            outputDuration,
            &format,
            NULL);
    }

    if (FAILED(hr)) {
        // Handle error
        captureClient.Reset();
        outputClient.Reset();
        inputClient.Reset();
        inputDevice.Reset();
        outputDevice.Reset();
        deviceEnumerator.Reset();
        CoUninitialize();
        return hr;
    }

    hr = outputClient->GetBufferSize(&bufferFrameCount);
    if (FAILED(hr)) {
        // Handle error
        captureClient.Reset();
        outputClient.Reset();
        inputClient.Reset();
        inputDevice.Reset();
        outputDevice.Reset();
        deviceEnumerator.Reset();
        CoUninitialize();
        return hr;
    }

    hr = outputClient->GetService(
        __uuidof(IAudioRenderClient),
        reinterpret_cast<void**>(renderClient.GetAddressOf())
    );
    if (FAILED(hr)) {
        // Handle error
        captureClient.Reset();
        outputClient.Reset();
        inputClient.Reset();
        inputDevice.Reset();
        outputDevice.Reset();
        deviceEnumerator.Reset();
        CoUninitialize();
        return hr;
    }

    inputEventHandle = CreateEvent(NULL, FALSE, FALSE, NULL);
    if (inputEventHandle == NULL)
    {
        return S_FALSE;
    }

    outputEventHandle = CreateEvent(NULL, FALSE, FALSE, NULL);
    if (outputEventHandle == NULL)
    {
        return S_FALSE;
    }

    hr = inputClient->SetEventHandle(inputEventHandle);
    if (FAILED(hr))
    {
        return hr;
    }


    outputClient->SetEventHandle(outputEventHandle);
    if (FAILED(hr))
    {
        return hr;
    }

    return S_OK;
}

std::queue<float*> bufferQueue;

void StartAudioLoopback() {
    HRESULT hr = inputClient->Start();
    if (FAILED(hr)) {
        // Handle error
        renderClient.Reset();
        outputClient.Reset();
        captureClient.Reset();
        inputClient.Reset();
        inputDevice.Reset();
        outputDevice.Reset();
        deviceEnumerator.Reset();
        CoUninitialize();
        return;
    }

    hr = outputClient->Start();
    if (FAILED(hr)) {
        // Handle error
        inputClient->Stop();
        renderClient.Reset();
        outputClient.Reset();
        captureClient.Reset();
        inputClient.Reset();
        inputDevice.Reset();
        outputDevice.Reset();
        deviceEnumerator.Reset();
        CoUninitialize();
        return;
    }

    HANDLE handles[2] = { inputEventHandle, outputEventHandle };

    int bufferSampleCount = 480000;
    short* recordingBuffer = (short*)malloc(bufferSampleCount * sizeof(short) * 2);
    memset(recordingBuffer, 0, bufferSampleCount * sizeof(short) * 2);

    int readHead = 0;
    int writeHead = 400;

    BYTE* pData;
    while (true) {
        DWORD waitResult = WaitForMultipleObjects(2, handles, FALSE, INFINITE);
        if (waitResult == WAIT_OBJECT_0) {
            // Input event received
            UINT32 packetSize;
            HRESULT hr = captureClient->GetNextPacketSize(&packetSize);

            while (packetSize != 0)
            {
                UINT32 numFramesAvailable;
                DWORD flags;

                hr = captureClient->GetBuffer(&pData, &numFramesAvailable, &flags, NULL, NULL);
                if (SUCCEEDED(hr))
                {
                    short* data = (short*)pData;
                    for (int i = 0; i < numFramesAvailable; i++)
                    {
                        recordingBuffer[writeHead * 2 + 0] = data[i * 2 + 0];
                        recordingBuffer[writeHead * 2 + 1] = data[i * 2 + 1];
                        writeHead += 1;
                        if (writeHead >= bufferSampleCount)
                        {
                            writeHead = 0;
                        }
                    }

                    hr = captureClient->ReleaseBuffer(numFramesAvailable);
                }

                hr = captureClient->GetNextPacketSize(&packetSize);
            }
        }
        else if (waitResult == WAIT_OBJECT_0 + 1) {
            UINT numFramesAvailable = bufferFrameCount;
            // Output event received
            HRESULT hr = renderClient->GetBuffer(numFramesAvailable, &pData);
            if (SUCCEEDED(hr)) {
                short* data = (short*)pData;
                for (int i = 0; i < numFramesAvailable; i++)
                {
                    data[i * 2 + 0] = recordingBuffer[readHead * 2 + 0];
                    data[i * 2 + 1] = recordingBuffer[readHead * 2 + 1];
                    readHead += 1;
                    if (readHead >= bufferSampleCount)
                    {
                        readHead = 0;
                    }
                }

                hr = renderClient->ReleaseBuffer(numFramesAvailable, 0);
                if (FAILED(hr)) {
                    // Handle error
                    break;
                }
            }
        }
        else {
            // Other events can be processed here if needed
        }
    }

    inputClient->Stop();
    outputClient->Stop();
    renderClient.Reset();
    outputClient.Reset();
    captureClient.Reset();
    inputClient.Reset();
    inputDevice.Reset();
    outputDevice.Reset();
    deviceEnumerator.Reset();
    CoUninitialize();
}

int WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nShowCmd)
{
    HRESULT hr = InitializeAudioClients();
    if (FAILED(hr)) {
        // Handle error
        return -1;
    }

    StartAudioLoopback();

    return 0;
}

#include <iostream>
#include <comdef.h>
#include <Wbemidl.h>
#include <vector>
#include <string>
#include <iomanip>

#pragma comment(lib, "wbemuuid.lib")

// Helper function to trim whitespace
std::string trim(const std::string& str) {
    size_t first = str.find_first_not_of(" \t\n\r");
    if (first == std::string::npos) return "";
    size_t last = str.find_last_not_of(" \t\n\r");
    return str.substr(first, last - first + 1);
}

// Helper function to convert BSTR to std::string
std::string BSTRToString(BSTR bstr) {
    if (!bstr) return "";
    char* str = _com_util::ConvertBSTRToString(bstr);
    std::string result(str);
    delete[] str;
    return result;
}

// Helper function to convert VARIANT to string
std::string VariantToString(VARIANT var) {
    if (var.vt == VT_BSTR) {
        return trim(BSTRToString(var.bstrVal));
    } else if (var.vt == VT_I4) {
        return std::to_string(var.lVal);
    } else if (var.vt == VT_UI4) {
        return std::to_string(var.ulVal);
    } else if (var.vt == VT_I8) {
        return std::to_string(var.llVal);
    } else if (var.vt == VT_UI8) {
        return std::to_string(var.ullVal);
    } else if (var.vt == VT_R4) {
        return std::to_string(var.fltVal);
    } else if (var.vt == VT_BOOL) {
        return var.boolVal ? "True" : "False";
    }
    return "Unknown";
}

// Forward declaration for VariantToULL (defined later)
unsigned long long VariantToULL(VARIANT var);

// Function to initialize COM and WMI
HRESULT InitializeWMI(IWbemServices** pSvc) {
    HRESULT hres;

    // Initialize COM
    hres = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (FAILED(hres)) {
        std::cout << "Failed to initialize COM library. Error code = 0x" << std::hex << hres << std::endl;
        return hres;
    }

    // Initialize COM security
    hres = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_DEFAULT,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        EOAC_NONE,
        NULL
    );
    if (FAILED(hres)) {
        std::cout << "Failed to initialize security. Error code = 0x" << std::hex << hres << std::endl;
        CoUninitialize();
        return hres;
    }

    // Obtain the initial locator to WMI
    IWbemLocator* pLoc = NULL;
    hres = CoCreateInstance(
        CLSID_WbemLocator,
        0,
        CLSCTX_INPROC_SERVER,
        IID_IWbemLocator,
        (LPVOID*)&pLoc
    );
    if (FAILED(hres)) {
        std::cout << "Failed to create IWbemLocator object. Error code = 0x" << std::hex << hres << std::endl;
        CoUninitialize();
        return hres;
    }

    // Connect to WMI through the IWbemLocator::ConnectServer method
    hres = pLoc->ConnectServer(
        _bstr_t(L"ROOT\\CIMV2"),
        NULL,
        NULL,
        0,
        NULL,
        0,
        0,
        pSvc
    );
    if (FAILED(hres)) {
        std::cout << "Could not connect. Error code = 0x" << std::hex << hres << std::endl;
        pLoc->Release();
        CoUninitialize();
        return hres;
    }

    // Set security levels on the proxy
    hres = CoSetProxyBlanket(
        *pSvc,
        RPC_C_AUTHN_WINNT,
        RPC_C_AUTHZ_NONE,
        NULL,
        RPC_C_AUTHN_LEVEL_CALL,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        EOAC_NONE
    );
    if (FAILED(hres)) {
        std::cout << "Could not set proxy blanket. Error code = 0x" << std::hex << hres << std::endl;
        (*pSvc)->Release();
        pLoc->Release();
        CoUninitialize();
        return hres;
    }

    pLoc->Release();
    return WBEM_S_NO_ERROR;
}

// Function to execute WMI query and get results
HRESULT ExecuteQuery(IWbemServices* pSvc, const std::string& query, IEnumWbemClassObject** pEnumerator) {
    HRESULT hres = pSvc->ExecQuery(
        bstr_t("WQL"),
        bstr_t(query.c_str()),
        WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
        NULL,
        pEnumerator
    );
    if (FAILED(hres)) {
        std::cout << "Query for " << query << " failed. Error code = 0x" << std::hex << hres << std::endl;
        return hres;
    }
    return WBEM_S_NO_ERROR;
}

// Function to get CPU information
void GetCPUInfo(IWbemServices* pSvc) {
    std::cout << "=== CPU Information ===" << std::endl;
    IEnumWbemClassObject* pEnumerator = NULL;
    if (FAILED(ExecuteQuery(pSvc, "SELECT Name, Manufacturer, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed FROM Win32_Processor", &pEnumerator))) {
        return;
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;
    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) break;

        VARIANT vtProp;
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Name: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Manufacturer: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"NumberOfCores", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Cores: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"NumberOfLogicalProcessors", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Logical Processors: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"MaxClockSpeed", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Max Clock Speed: " << VariantToString(vtProp) << " MHz" << std::endl;
            VariantClear(&vtProp);
        }

        pclsObj->Release();
        std::cout << std::endl;
    }
    pEnumerator->Release();
}

// Function to get RAM information
void GetRAMInfo(IWbemServices* pSvc) {
    std::cout << "=== RAM Information ===" << std::endl;
    IEnumWbemClassObject* pEnumerator = NULL;
    if (FAILED(ExecuteQuery(pSvc, "SELECT Capacity, Speed, Manufacturer, MemoryType FROM Win32_PhysicalMemory", &pEnumerator))) {
        return;
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;
    unsigned long long totalCapacity = 0;
    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) break;

        VARIANT vtProp;
        hr = pclsObj->Get(L"Capacity", 0, &vtProp, 0, 0);
        unsigned long long capacity = 0;
        if (SUCCEEDED(hr)) {
            capacity = VariantToULL(vtProp);
            totalCapacity += capacity;
            std::cout << "Module Capacity: " << (capacity / (1024ULL * 1024ULL * 1024ULL)) << " GB" << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"Speed", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Speed: " << VariantToString(vtProp) << " MHz" << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Manufacturer: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"MemoryType", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            int type = vtProp.lVal;
            std::string typeStr;
            switch (type) {
                case 20: typeStr = "DDR"; break;
                case 21: typeStr = "DDR2"; break;
                case 24: typeStr = "DDR3"; break;
                case 26: typeStr = "DDR4"; break;
                default: typeStr = "Unknown"; break;
            }
            std::cout << "Type: " << typeStr << std::endl;
            VariantClear(&vtProp);
        }

        pclsObj->Release();
        std::cout << std::endl;
    }
    pEnumerator->Release();
    std::cout << "Total RAM: " << (totalCapacity / (1024 * 1024 * 1024)) << " GB" << std::endl << std::endl;
}

// Function to get GPU information
void GetGPUInfo(IWbemServices* pSvc) {
    std::cout << "=== GPU Information ===" << std::endl;
    IEnumWbemClassObject* pEnumerator = NULL;
    if (FAILED(ExecuteQuery(pSvc, "SELECT Name, AdapterRAM, VideoProcessor FROM Win32_VideoController", &pEnumerator))) {
        return;
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;
    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) break;

        VARIANT vtProp;
        hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Name: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"AdapterRAM", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            unsigned long long ram = VariantToULL(vtProp);
            std::cout << "Memory: " << (ram / (1024ULL * 1024ULL * 1024ULL)) << " GB" << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"VideoProcessor", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Video Processor: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        pclsObj->Release();
        std::cout << std::endl;
    }
    pEnumerator->Release();
}

// Function to get storage information
void GetStorageInfo(IWbemServices* pSvc) {
    std::cout << "=== Storage Information ===" << std::endl;
    IEnumWbemClassObject* pEnumerator = NULL;
    if (FAILED(ExecuteQuery(pSvc, "SELECT Model, Manufacturer, TotalSectors, BytesPerSector, SerialNumber, InterfaceType, MediaType FROM Win32_DiskDrive", &pEnumerator))) {
        return;
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;
    while (pEnumerator) {
        HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
        if (0 == uReturn) break;

        VARIANT vtProp;
        hr = pclsObj->Get(L"Model", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Model: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        std::string manufacturer = "Unknown";
        if (SUCCEEDED(hr)) {
            manufacturer = VariantToString(vtProp);
            VariantClear(&vtProp);
        }
        // Try to infer manufacturer from model if not available
        if (manufacturer == "(Standard disk drives)" || manufacturer.empty()) {
            hr = pclsObj->Get(L"Model", 0, &vtProp, 0, 0);
            if (SUCCEEDED(hr)) {
                std::string model = VariantToString(vtProp);
                if (model.substr(0, 2) == "ST") manufacturer = "Seagate";
                else if (model.substr(0, 2) == "WD") manufacturer = "Western Digital";
                else if (model.find("Samsung") != std::string::npos) manufacturer = "Samsung";
                else if (model.find("SK Hynix") != std::string::npos) manufacturer = "SK Hynix";
                else if (model.find("Kingston") != std::string::npos) manufacturer = "Kingston";
                else if (model.find("Crucial") != std::string::npos) manufacturer = "Crucial";
                VariantClear(&vtProp);
            }
        }
        std::cout << "Manufacturer: " << manufacturer << std::endl;

        hr = pclsObj->Get(L"TotalSectors", 0, &vtProp, 0, 0);
        unsigned long long totalSectors = 0;
        if (SUCCEEDED(hr)) {
            totalSectors = VariantToULL(vtProp);
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"BytesPerSector", 0, &vtProp, 0, 0);
        unsigned long bytesPerSector = 512; // default
        if (SUCCEEDED(hr)) {
            bytesPerSector = static_cast<unsigned long>(VariantToULL(vtProp));
            if (bytesPerSector == 0) bytesPerSector = 512;
            VariantClear(&vtProp);
        }

        unsigned long long sizeBytes = totalSectors * bytesPerSector;
        unsigned long long sizeGB = sizeBytes / (1024ULL * 1024 * 1024);
        std::cout << "Size: " << sizeGB << " GB" << std::endl;

        hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Serial Number: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"InterfaceType", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Interface Type: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"MediaType", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Media Type: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        pclsObj->Release();
        std::cout << std::endl;
    }
    pEnumerator->Release();
}

// Function to get motherboard information
void GetMotherboardInfo(IWbemServices* pSvc) {
    std::cout << "=== Motherboard Information ===" << std::endl;
    IEnumWbemClassObject* pEnumerator = NULL;
    if (FAILED(ExecuteQuery(pSvc, "SELECT Manufacturer, Product, SerialNumber, Version FROM Win32_BaseBoard", &pEnumerator))) {
        return;
    }

    IWbemClassObject* pclsObj = NULL;
    ULONG uReturn = 0;
    if (pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn) == WBEM_S_NO_ERROR && uReturn > 0) {
        VARIANT vtProp;
        HRESULT hr = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Manufacturer: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"Product", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Product: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"SerialNumber", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Serial Number: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        hr = pclsObj->Get(L"Version", 0, &vtProp, 0, 0);
        if (SUCCEEDED(hr)) {
            std::cout << "Version: " << VariantToString(vtProp) << std::endl;
            VariantClear(&vtProp);
        }

        pclsObj->Release();
    }
    pEnumerator->Release();
    std::cout << std::endl;
}

int main() {
    IWbemServices* pSvc = NULL;
    if (FAILED(InitializeWMI(&pSvc))) {
        return 1;
    }

    GetCPUInfo(pSvc);
    GetRAMInfo(pSvc);
    GetGPUInfo(pSvc);
    GetStorageInfo(pSvc);
    GetMotherboardInfo(pSvc);

    pSvc->Release();
    CoUninitialize();

    std::cout << "Press Enter to exit...";
    std::cin.get();
    return 0;
}

// Convert VARIANT to unsigned long long safely
unsigned long long VariantToULL(VARIANT var) {
    if (var.vt == VT_UI8) return var.ullVal;
    if (var.vt == VT_I8) return static_cast<unsigned long long>(var.llVal);
    if (var.vt == VT_UI4) return var.ulVal;
    if (var.vt == VT_I4) return static_cast<unsigned long long>(var.lVal);
    if (var.vt == VT_R8) return static_cast<unsigned long long>(var.dblVal);
    if (var.vt == VT_R4) return static_cast<unsigned long long>(var.fltVal);
    if (var.vt == VT_BSTR && var.bstrVal) {
        try {
            std::string s = trim(BSTRToString(var.bstrVal));
            if (s.empty()) return 0ULL;
            return std::stoull(s);
        } catch (...) {
            return 0ULL;
        }
    }
    return 0ULL;
}
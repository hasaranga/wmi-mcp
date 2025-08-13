#include <iostream>
#include <stdio.h>
#include <Windows.h>
#include <comdef.h>
#include <Wbemidl.h>
#include <propvarutil.h>
#include <string>
#include <sstream>
#include <vector>
#include "json.hpp"

#pragma comment(lib, "wbemuuid.lib")
#pragma comment(lib, "Propsys.lib")

using json = nlohmann::json;

class WMIQueryExecutor {
private:
    IWbemServices* pSvc;
    IWbemLocator* pLoc;
    bool initialized;

public:
    WMIQueryExecutor() : pSvc(nullptr), pLoc(nullptr), initialized(false) {}

    ~WMIQueryExecutor() {
        cleanup();
    }

    bool initialize() {
        HRESULT hres;

        // Initialize COM
        hres = CoInitializeEx(0, COINIT_MULTITHREADED);
        if (FAILED(hres)) return false;

        // Set security
        hres = CoInitializeSecurity(
            NULL, -1, NULL, NULL,
            RPC_C_AUTHN_LEVEL_DEFAULT,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            NULL, EOAC_NONE, NULL
        );
        if (FAILED(hres)) {
            CoUninitialize();
            return false;
        }

        // Connect to WMI
        hres = CoCreateInstance(
            CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER,
            IID_IWbemLocator, (LPVOID*)&pLoc
        );
        if (FAILED(hres)) {
            CoUninitialize();
            return false;
        }

        initialized = true;
        return true;
    }

    json executeQuery(const std::string& wmiNamespace, const std::string& query) {
        json result = {
            {"success", false},
            {"namespace", wmiNamespace},
            {"query", query},
            {"objects", json::array()},
            {"count", 0},
            {"error", ""}
        };

        if (!initialized) {
            result["error"] = "WMI not initialized";
            return result;
        }

        HRESULT hres;
        IWbemServices* pTempSvc = nullptr;

        // Convert namespace to wide string
        std::wstring wideNamespace(wmiNamespace.begin(), wmiNamespace.end());

        // Connect to the specified namespace
        hres = pLoc->ConnectServer(
            _bstr_t(wideNamespace.c_str()),
            NULL, NULL, 0, NULL, 0, 0, &pTempSvc
        );

        if (FAILED(hres)) {
            result["error"] = "Failed to connect to WMI namespace: " + wmiNamespace;
            return result;
        }

        // Set proxy security
        hres = CoSetProxyBlanket(
            pTempSvc, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, NULL,
            RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE,
            NULL, EOAC_NONE
        );

        if (FAILED(hres)) {
            pTempSvc->Release();
            result["error"] = "Failed to set proxy security";
            return result;
        }

        // Execute query
        IEnumWbemClassObject* pEnumerator = NULL;
        std::wstring wideQuery(query.begin(), query.end());

        hres = pTempSvc->ExecQuery(
            bstr_t("WQL"),
            bstr_t(wideQuery.c_str()),
            WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            NULL, &pEnumerator
        );

        if (FAILED(hres)) {
            pTempSvc->Release();
            result["error"] = "Failed to execute WMI query: " + query;
            return result;
        }

        // Process results
        IWbemClassObject* pclsObj = nullptr;
        ULONG uReturn = 0;
        int objectCount = 0;

        while (pEnumerator) {
            HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);
            if (!uReturn) break;

            objectCount++;
            json wmiObject = json::object();

            // Get all properties of the object
            SAFEARRAY* psaNames = NULL;
            HRESULT hrNames = pclsObj->GetNames(NULL, WBEM_FLAG_ALWAYS, NULL, &psaNames);

            if (SUCCEEDED(hrNames)) {
                LONG lLower, lUpper;
                SafeArrayGetLBound(psaNames, 1, &lLower);
                SafeArrayGetUBound(psaNames, 1, &lUpper);

                for (LONG i = lLower; i <= lUpper; i++) {
                    BSTR bstrName = NULL;
                    SafeArrayGetElement(psaNames, &i, &bstrName);

                    VARIANT vtProp;
                    VariantInit(&vtProp);
                    pclsObj->Get(bstrName, 0, &vtProp, 0, 0);

                    // Convert BSTR to string
                    std::string propName = _com_util::ConvertBSTRToString(bstrName);

                    // Handle different variant types and add to JSON
                    if (vtProp.vt == VT_BSTR && vtProp.bstrVal != NULL) {
                        std::string value = _com_util::ConvertBSTRToString(vtProp.bstrVal);
                        wmiObject[propName] = value;
                    }
                    else if (vtProp.vt == VT_I4) {
                        wmiObject[propName] = vtProp.intVal;
                    }
                    else if (vtProp.vt == VT_UI4) {
                        wmiObject[propName] = static_cast<uint32_t>(vtProp.uintVal);
                    }
                    else if (vtProp.vt == VT_I8) {
                        wmiObject[propName] = static_cast<int64_t>(vtProp.llVal);
                    }
                    else if (vtProp.vt == VT_UI8) {
                        wmiObject[propName] = static_cast<uint64_t>(vtProp.ullVal);
                    }
                    else if (vtProp.vt == VT_BOOL) {
                        wmiObject[propName] = vtProp.boolVal ? true : false;
                    }
                    else if (vtProp.vt == VT_NULL) {
                        wmiObject[propName] = "null";
                    }
                    else {
                        // For unsupported type!
                        WCHAR szBuffer[512];
                        if (SUCCEEDED(VariantToString(vtProp, szBuffer, ARRAYSIZE(szBuffer)))) {
                            int len = WideCharToMultiByte(CP_UTF8, 0, szBuffer, -1, NULL, 0, NULL, NULL);
                            std::string propValue(len - 1, '\0');
                            WideCharToMultiByte(CP_UTF8, 0, szBuffer, -1, &propValue[0], len, NULL, NULL);
                            wmiObject[propName] = propValue;
                        }
                        else {
                            wmiObject[propName] = "unknown";
                        }
                    }

                    VariantClear(&vtProp);
                    SysFreeString(bstrName);
                }

                SafeArrayDestroy(psaNames);
            }

            result["objects"].push_back(wmiObject);
            pclsObj->Release();
        }

        result["count"] = objectCount;
        result["success"] = true;

        // Cleanup
        pEnumerator->Release();
        pTempSvc->Release();

        return result;
    }

    void cleanup() {
        if (pSvc) {
            pSvc->Release();
            pSvc = nullptr;
        }
        if (pLoc) {
            pLoc->Release();
            pLoc = nullptr;
        }
        if (initialized) {
            CoUninitialize();
            initialized = false;
        }
    }
};

int main() {
    // Initialize WMI
    WMIQueryExecutor wmi;
    if (!wmi.initialize()) {
        std::cerr << "Failed to initialize WMI" << std::endl;
        return 1;
    }

    std::string line;

    while (std::getline(std::cin, line)) {
        if (line.empty()) continue;

        try {
            json request = json::parse(line);
            json response;

            // Extract ID and method
            auto id = request.contains("id") ? request["id"] : nullptr;
            std::string method = request.value("method", "");

            response["jsonrpc"] = "2.0";
            response["id"] = id;

            if (method == "initialize") {
                response["result"] = {
                    {"protocolVersion", "2024-11-05"},
                    {"capabilities", {{"tools", json::object()}}},
                    {"serverInfo", {
                        {"name", "wmi-query-mcp"},
                        {"version", "1.0.0"}
                    }}
                };
            }
            else if (method == "tools/list") {
                response["result"] = {
                    {"tools", json::array({
                        {
                            {"name", "executeWMIQuery"},
                            {"description", "Executes WMI (Windows Management Instrumentation) queries to retrieve system information from Windows computers. This tool can query hardware details, running processes, services, system configuration, network information, disk usage, and much more. Common namespaces include ROOT\\CIMV2 (most common system info), ROOT\\WMI (hardware sensors), and ROOT\\RSOP (policy info). Use WQL (WMI Query Language) syntax similar to SQL."},
                            {"inputSchema", {
                                {"type", "object"},
                                {"properties", {
                                    {"namespace", {
                                        {"type", "string"},
                                        {"description", "WMI namespace to query (e.g., ROOT\\CIMV2, ROOT\\WMI)"},
                                        {"default", "ROOT\\CIMV2"}
                                    }},
                                    {"query", {
                                        {"type", "string"},
                                        {"description", "WQL query to execute (e.g., SELECT ExecutablePath FROM Win32_Process, SELECT Name, Size FROM Win32_LogicalDisk)"}
                                    }}
                                }},
                                {"required", json::array({"query"})}
                            }}
                        }
                    })}
                };
            }
            else if (method == "tools/call" &&
                request.contains("params") &&
                request["params"].contains("name") &&
                request["params"]["name"] == "executeWMIQuery") {

                // Get parameters
                std::string wmiNamespace = "ROOT\\CIMV2"; // default
                std::string query = "";

                if (request["params"].contains("arguments")) {
                    auto args = request["params"]["arguments"];
                    if (args.contains("namespace")) {
                        wmiNamespace = args["namespace"];
                    }
                    if (args.contains("query")) {
                        query = args["query"];
                    }
                }

                if (query.empty()) {
                    response["error"] = {
                        {"code", -32602},
                        {"message", "Invalid params: query parameter is required"}
                    };
                }
                else {
                    // Execute WMI query
                    json queryResult = wmi.executeQuery(wmiNamespace, query);

                    if (queryResult["success"]) {
                        // Return structured JSON data
                        response["result"] = {
                            {"content", json::array({
                                {
                                    {"type", "text"},
                                    {"text", "WMI query executed successfully. Found " +
                                            std::to_string(queryResult["count"].get<int>()) + " objects."}
                                },
                                {
                                    {"type", "text"},
                                    {"text", "Raw JSON data:\n" + queryResult.dump(2)}
                                }
                            })}
                        };
                    }
                    else {
                        response["result"] = {
                            {"content", json::array({
                                {
                                    {"type", "text"},
                                    {"text", "WMI query failed: " + queryResult["error"].get<std::string>()}
                                }
                            })}
                        };
                    }
                }
            }
            else {
                // Unknown method
                response["error"] = {
                    {"code", -32601},
                    {"message", "Method not found"}
                };
            }

            std::cout << response.dump() << std::endl;
            std::cout.flush();

        }
        catch (const json::exception& e) {
            (void)e; // Suppress unused warning
            // Handle JSON parsing errors
            json error_response = {
                {"jsonrpc", "2.0"},
                {"id", nullptr},
                {"error", {
                    {"code", -32700},
                    {"message", "Parse error"}
                }}
            };
            std::cout << error_response.dump() << std::endl;
            std::cout.flush();
        }
    }

    return 0;
}
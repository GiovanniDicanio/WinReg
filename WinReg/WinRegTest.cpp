//////////////////////////////////////////////////////////////////////////
//
// WinRegTest.cpp -- by Giovanni Dicanio
//
// Test some of the WinReg code
//
// NOTE --- Test Preparation ---
// In the folder containing this source file, there should be also a file
// "GioTest.reg". This REG file contains some initial data to load into
// the registry for this test.
//
//////////////////////////////////////////////////////////////////////////


#include "WinReg.hpp"   // Module to test

#include <exception>
#include <iostream>
#include <string>
#include <vector>


using std::pair;
using std::vector;
using std::wcout;
using std::wstring;

using winreg::RegKey;
using winreg::RegException;
using winreg::RegResult;


// Test common RegKey methods
void Test()
{
    wcout << "\n *** Testing Common RegKey Methods *** \n\n";

    //
    // Test subkey and value enumeration
    //

    const wstring testSubKey = L"SOFTWARE\\GioTest";
    RegKey key(HKEY_CURRENT_USER, testSubKey);

    vector<wstring> subKeyNames = key.EnumSubKeys();
    wcout << L"Subkeys:\n";
    for (const auto& s : subKeyNames)
    {
        wcout << L"  [" << s << L"]\n";
    }
    wcout << L'\n';

    vector<pair<wstring, DWORD>> values = key.EnumValues();
    wcout << L"Values:\n";
    for (const auto& v : values)
    {
        wcout << L"  [" << v.first << L"](" << RegKey::RegTypeToString(v.second) << L")\n";
    }
    wcout << L'\n';

    key.Close();


    //
    // Test SetXxxValue and GetXxxValue methods
    //

    key.Open(HKEY_CURRENT_USER, testSubKey);

    const DWORD testDw = 0x1234ABCD;
    const ULONGLONG testQw = 0xAABBCCDD11223344;
    const wstring testSz = L"CiaoTestSz";
    const wstring testExpandSz = L"%PATH%";
    const vector<BYTE> testBinary{ 0xAA, 0xBB, 0xCC, 0x11, 0x22, 0x33 };
    const vector<wstring> testMultiSz{ L"Hi", L"Hello", L"Ciao" };

    key.SetDwordValue(L"TestValueDword", testDw);
    key.SetQwordValue(L"TestValueQword", testQw);
    key.SetStringValue(L"TestValueString", testSz);
    key.SetExpandStringValue(L"TestValueExpandString", testExpandSz);
    key.SetMultiStringValue(L"TestValueMultiString", testMultiSz);
    key.SetBinaryValue(L"TestValueBinary", testBinary);

    DWORD testDw1 = key.GetDwordValue(L"TestValueDword");
    if (testDw1 != testDw)
    {
        wcout << L"RegKey::GetDwordValue failed.\n";
    }

    DWORD typeId = key.QueryValueType(L"TestValueDword");
    if (typeId != REG_DWORD)
    {
        wcout << L"RegKey::QueryValueType failed for REG_DWORD.\n";
    }

    ULONGLONG testQw1 = key.GetQwordValue(L"TestValueQword");
    if (testQw1 != testQw)
    {
        wcout << L"RegKey::GetQwordValue failed.\n";
    }

    typeId = key.QueryValueType(L"TestValueQword");
    if (typeId != REG_QWORD)
    {
        wcout << L"RegKey::QueryValueType failed for REG_QWORD.\n";
    }

    wstring testSz1 = key.GetStringValue(L"TestValueString");
    if (testSz1 != testSz)
    {
        wcout << L"RegKey::GetStringValue failed.\n";
    }

    typeId = key.QueryValueType(L"TestValueString");
    if (typeId != REG_SZ)
    {
        wcout << L"RegKey::QueryValueType failed for REG_SZ.\n";
    }

    wstring testExpandSz1 = key.GetExpandStringValue(L"TestValueExpandString");
    if (testExpandSz1 != testExpandSz)
    {
        wcout << L"RegKey::GetExpandStringValue failed.\n";
    }

    typeId = key.QueryValueType(L"TestValueExpandString");
    if (typeId != REG_EXPAND_SZ)
    {
        wcout << L"RegKey::QueryValueType failed for REG_EXPAND_SZ.\n";
    }

    vector<wstring> testMultiSz1 = key.GetMultiStringValue(L"TestValueMultiString");
    if (testMultiSz1 != testMultiSz)
    {
        wcout << L"RegKey::GetMultiStringValue failed.\n";
    }

    typeId = key.QueryValueType(L"TestValueMultiString");
    if (typeId != REG_MULTI_SZ)
    {
        wcout << L"RegKey::QueryValueType failed for REG_MULTI_SZ.\n";
    }

    vector<BYTE> testBinary1 = key.GetBinaryValue(L"TestValueBinary");
    if (testBinary1 != testBinary)
    {
        wcout << L"RegKey::GetBinaryValue failed.\n";
    }

    typeId = key.QueryValueType(L"TestValueBinary");
    if (typeId != REG_BINARY)
    {
        wcout << L"RegKey::QueryValueType failed for REG_BINARY.\n";
    }


    //
    // Remove some test values
    //

    key.DeleteValue(L"TestValueDword");
    key.DeleteValue(L"TestValueQword");
    key.DeleteValue(L"TestValueString");
    key.DeleteValue(L"TestValueExpandString");
    key.DeleteValue(L"TestValueMultiString");
    key.DeleteValue(L"TestValueBinary");
}


// Throws on failure retCode; used by TestTryMethods()
static inline void Check(const RegResult& retCode, const char* message)
{
    if (retCode.Failed())
    {
        throw RegException(retCode.Code(), message);
    }
}


// Test RegKey methods that return error codes
void TestTryMethods()
{
    wcout << "\n *** Testing Error-code-returning RegKey::TryAction Methods *** \n\n";

    //
    // Test subkey and value enumeration
    //

    const wstring testSubKey = L"SOFTWARE\\GioTest";
    RegKey key;
    RegResult retCode = key.TryOpen(HKEY_CURRENT_USER, testSubKey);
    Check(retCode, "RegKey::TryOpen failed.");

    vector<wstring> subKeyNames;
    retCode = key.TryEnumSubKeys(subKeyNames);
    Check(retCode, "RegKey::TryEnumSubKeys failed.");

    wcout << L"Subkeys:\n";
    for (const auto& s : subKeyNames)
    {
        wcout << L"  [" << s << L"]\n";
    }
    wcout << L'\n';

    vector<pair<wstring, DWORD>> values;
    retCode = key.TryEnumValues(values);
    Check(retCode, "RegKey::TryEnumValues failed.");

    wcout << L"Values:\n";
    for (const auto& v : values)
    {
        wcout << L"  [" << v.first << L"](" << RegKey::RegTypeToString(v.second) << L")\n";
    }
    wcout << L'\n';

    key.Close();


    //
    // Test TrySetXxxValue and TryGetXxxValue methods
    //

    retCode = key.TryOpen(HKEY_CURRENT_USER, testSubKey);
    Check(retCode, "RegKey::TryOpen failed.");

    const DWORD testDw = 0x1234ABCD;
    const ULONGLONG testQw = 0xAABBCCDD11223344;
    const wstring testSz = L"CiaoTestSz";
    const wstring testExpandSz = L"%PATH%";
    const vector<BYTE> testBinary{ 0xAA, 0xBB, 0xCC, 0x11, 0x22, 0x33 };
    const vector<wstring> testMultiSz{ L"Hi", L"Hello", L"Ciao" };

    retCode = key.TrySetDwordValue(L"TestValueDword", testDw);
    Check(retCode, "RegKey::TrySetDwordValue failed.");

    retCode = key.TrySetQwordValue(L"TestValueQword", testQw);
    Check(retCode, "RegKey:TrySetQwordValue: failed.");

    retCode = key.TrySetStringValue(L"TestValueString", testSz);
    Check(retCode, "RegKey::TrySetStringValue failed.");

    retCode = key.TrySetExpandStringValue(L"TestValueExpandString", testExpandSz);
    Check(retCode, "RegKey::TrySetExpandStringValue string failed.");

    retCode = key.TrySetMultiStringValue(L"TestValueMultiString", testMultiSz);
    Check(retCode, "RegKey::TrySetMultiStringValue failed.");

    retCode = key.TrySetBinaryValue(L"TestValueBinary", testBinary);
    Check(retCode, "RegKey::TrySetBinaryValue failed.");


    DWORD testDw1{};
    retCode = key.TryGetDwordValue(L"TestValueDword", testDw1);
    Check(retCode, "RegKey::TryGetDwordValue failed.");
    if (testDw1 != testDw)
    {
        wcout << L"RegKey::TryGetDwordValue failed.\n";
    }

    DWORD typeId{};
    retCode = key.TryQueryValueType(L"TestValueDword", typeId);
    Check(retCode, "RegKey::TryQueryValueType failed.");
    if (typeId != REG_DWORD)
    {
        wcout << L"RegKey::TryQueryValueType failed for REG_DWORD.\n";
    }

    ULONGLONG testQw1{};
    retCode = key.TryGetQwordValue(L"TestValueQword", testQw1);
    Check(retCode, "RegKey::TryGetQwordValue failed.");
    if (testQw1 != testQw)
    {
        wcout << L"RegKey::TryGetQwordValue failed.\n";
    }

    retCode = key.TryQueryValueType(L"TestValueQword", typeId);
    Check(retCode, "RegKey::TryQueryValueType failed.");
    if (typeId != REG_QWORD)
    {
        wcout << L"RegKey::TryQueryValueType failed for REG_QWORD.\n";
    }

    wstring testSz1;
    retCode = key.TryGetStringValue(L"TestValueString", testSz1);
    Check(retCode, "RegKey::TryGetStringValue failed.");
    if (testSz1 != testSz)
    {
        wcout << L"RegKey::TryGetStringValue failed.\n";
    }

    retCode = key.TryQueryValueType(L"TestValueString", typeId);
    Check(retCode, "RegKey::TryQueryValueType failed.");
    if (typeId != REG_SZ)
    {
        wcout << L"RegKey::TryQueryValueType failed for REG_SZ.\n";
    }

    wstring testExpandSz1;
    retCode = key.TryGetExpandStringValue(L"TestValueExpandString", testExpandSz1);
    Check(retCode, "RegKey::TryGetExpandStringValue failed.");
    if (testExpandSz1 != testExpandSz)
    {
        wcout << L"RegKey::TryGetExpandStringValue failed.\n";
    }

    retCode = key.TryQueryValueType(L"TestValueExpandString", typeId);
    Check(retCode, "RegKey::TryQueryValueType failed.");
    if (typeId != REG_EXPAND_SZ)
    {
        wcout << L"RegKey::TryQueryValueType failed for REG_EXPAND_SZ.\n";
    }

    vector<wstring> testMultiSz1;
    retCode = key.TryGetMultiStringValue(L"TestValueMultiString", testMultiSz1);
    Check(retCode, "RegKey::TryGetMultiStringValue failed.");
    if (testMultiSz1 != testMultiSz)
    {
        wcout << L"RegKey::TryGetMultiStringValue failed.\n";
    }

    retCode = key.TryQueryValueType(L"TestValueMultiString", typeId);
    Check(retCode, "RegKey::TryQueryValueType failed.");
    if (typeId != REG_MULTI_SZ)
    {
        wcout << L"RegKey::TryQueryValueType failed for REG_MULTI_SZ.\n";
    }

    vector<BYTE> testBinary1;
    retCode = key.TryGetBinaryValue(L"TestValueBinary", testBinary1);
    Check(retCode, "RegKey::TryGetBinaryValue failed.");
    if (testBinary1 != testBinary)
    {
        wcout << L"RegKey::TryGetBinaryValue failed.\n";
    }

    retCode = key.TryQueryValueType(L"TestValueBinary", typeId);
    Check(retCode, "RegKey::TryQueryValueType failed.");
    if (typeId != REG_BINARY)
    {
        wcout << L"RegKey::TryQueryValueType failed for REG_BINARY.\n";
    }


    //
    // Remove some test values
    //

    Check(key.TryDeleteValue(L"TestValueDword"),        "RegKey::TryDeleteValue failed.");
    Check(key.TryDeleteValue(L"TestValueQword"),        "RegKey::TryDeleteValue failed.");
    Check(key.TryDeleteValue(L"TestValueString"),       "RegKey::TryDeleteValue failed.");
    Check(key.TryDeleteValue(L"TestValueExpandString"), "RegKey::TryDeleteValue failed.");
    Check(key.TryDeleteValue(L"TestValueMultiString"),  "RegKey::TryDeleteValue failed.");
    Check(key.TryDeleteValue(L"TestValueBinary"),       "RegKey::TryDeleteValue failed.");
}


int main()
{
    constexpr int kExitOk = 0;
    constexpr int kExitError = 1;

    try
    {
        wcout << L"=========================================\n";
        wcout << L"*** Testing Giovanni Dicanio's WinReg ***\n";
        wcout << L"=========================================\n\n";

        Test();
        TestTryMethods();

        wcout << L"All right!! :)\n\n";
    }
    catch (const RegException& e)
    {
        wcout << L"\n*** Registry Exception: " << e.what();
        wcout << L"\n*** [Windows API error code = " << e.code() << L"]\n\n";
        return kExitError;
    }
    catch (const std::exception& e)
    {
        wcout << L"\n*** ERROR: " << e.what() << L'\n';
        return kExitError;
    }

    return kExitOk;
}


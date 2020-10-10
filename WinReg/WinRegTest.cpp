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

#include <exception>
#include <iostream>
#include <string>
#include <vector>
#include "WinReg.hpp"


using std::vector;

using winreg::RegException;
using RegKey = winreg::RegKey<char>;

// Test common RegKey methods
void Test()
{
    std::cout << "\n *** Testing Common RegKey Methods *** \n\n";

    //
    // Test subkey and value enumeration
    //

    const std::string testSubKey = "SOFTWARE\\GioTest";
    RegKey key{ HKEY_CURRENT_USER, testSubKey };

    std::vector<std::string> subKeyNames = key.EnumSubKeys();
    std::cout << "Subkeys:\n";
    for(const auto& s : subKeyNames)
    {
        std::cout << "  [" << s << "]\n";
    }
    std::cout << L'\n';

    vector<std::pair<std::string, DWORD>> values = key.EnumValues();
    std::cout << "Values:\n";
    for(const auto& v : values)
    {
        std::cout << "  [" << v.first << "](" << RegKey::RegTypeToString(v.second) << ")\n";
    }
    std::cout << L'\n';

    key.Close();

    //
    // Test SetXxxValue, GetXxxValue and TryGetXxxValue methods
    //

    key.Open(HKEY_CURRENT_USER, testSubKey);

    const DWORD testDw = 0x1234ABCD;
    const ULONGLONG testQw = 0xAABBCCDD11223344;
    const std::string testSz = "CiaoTestSz";
    const std::string testExpandSz = "%PATH%";
    const vector<BYTE> testBinary = { 0xAA, 0xBB, 0xCC, 0x11, 0x22, 0x33 };
    const vector<std::string> testMultiSz = { "Hi", "Hello", "Ciao" };

    key.SetDwordValue("TestValueDword", testDw);
    key.SetQwordValue("TestValueQword", testQw);
    key.SetStringValue("TestValueString", testSz);
    key.SetExpandStringValue("TestValueExpandString", testExpandSz);
    key.SetMultiStringValue("TestValueMultiString", testMultiSz);
    key.SetBinaryValue("TestValueBinary", testBinary);

    DWORD testDw1 = key.GetDwordValue("TestValueDword");
    if(testDw1 != testDw)
    {
        std::cout << "RegKey::GetDwordValue failed.\n";
    }

    if(auto testDw2 = key.TryGetDwordValue("TestValueDword"))
    {
        if(testDw2 != testDw)
        {
            std::cout << "RegKey::TryGetDwordValue failed.\n";
        }
    }
    else
    {
        std::cout << "RegKey::TryGetDwordValue failed (std::std::optional has no value).\n";
    }

    DWORD typeId = key.QueryValueType("TestValueDword");
    if(typeId != REG_DWORD)
    {
        std::cout << "RegKey::QueryValueType failed for REG_DWORD.\n";
    }

    ULONGLONG testQw1 = key.GetQwordValue("TestValueQword");
    if(testQw1 != testQw)
    {
        std::cout << "RegKey::GetQwordValue failed.\n";
    }

    if(auto testQw2 = key.TryGetQwordValue("TestValueQword"))
    {
        if(testQw2 != testQw)
        {
            std::cout << "RegKey::TryGetQwordValue failed.\n";
        }
    }
    else
    {
        std::cout << "RegKey::TryGetQwordValue failed (std::std::optional has no value).\n";
    }

    typeId = key.QueryValueType("TestValueQword");
    if(typeId != REG_QWORD)
    {
        std::cout << "RegKey::QueryValueType failed for REG_QWORD.\n";
    }

    std::string testSz1 = key.GetStringValue("TestValueString");
    if(testSz1 != testSz)
    {
        std::cout << "RegKey::GetStringValue failed.\n";
    }

    if(auto testSz2 = key.TryGetStringValue("TestValueString"))
    {
        if(testSz2 != testSz)
        {
            std::cout << "RegKey::TryGetStringValue failed.\n";
        }
    }
    else
    {
        std::cout << "RegKey::TryGetStringValue failed (std::std::optional has no value).\n";
    }

    typeId = key.QueryValueType("TestValueString");
    if(typeId != REG_SZ)
    {
        std::cout << "RegKey::QueryValueType failed for REG_SZ.\n";
    }

    std::string testExpandSz1 = key.GetExpandStringValue("TestValueExpandString");
    if(testExpandSz1 != testExpandSz)
    {
        std::cout << "RegKey::GetExpandStringValue failed.\n";
    }

    if(auto testExpandSz2 = key.TryGetExpandStringValue("TestValueExpandString"))
    {
        if(testExpandSz2 != testExpandSz)
        {
            std::cout << "RegKey::TryGetExpandStringValue failed.\n";
        }
    }
    else
    {
        std::cout << "RegKey::TryGetExpandStringValue failed (std::std::optional has no value).\n";
    }

    typeId = key.QueryValueType("TestValueExpandString");
    if(typeId != REG_EXPAND_SZ)
    {
        std::cout << "RegKey::QueryValueType failed for REG_EXPAND_SZ.\n";
    }

    vector<std::string> testMultiSz1 = key.GetMultiStringValue("TestValueMultiString");
    if(testMultiSz1 != testMultiSz)
    {
        std::cout << "RegKey::GetMultiStringValue failed.\n";
    }

    if(auto testMultiSz2 = key.TryGetMultiStringValue("TestValueMultiString"))
    {
        if(testMultiSz2 != testMultiSz)
        {
            std::cout << "RegKey::TryGetMultiStringValue failed.\n";
        }
    }
    else
    {
        std::cout << "RegKey::TryGetMultiStringValue failed (std::std::optional has no value).\n";
    }

    typeId = key.QueryValueType("TestValueMultiString");
    if(typeId != REG_MULTI_SZ)
    {
        std::cout << "RegKey::QueryValueType failed for REG_MULTI_SZ.\n";
    }

    vector<BYTE> testBinary1 = key.GetBinaryValue("TestValueBinary");
    if(testBinary1 != testBinary)
    {
        std::cout << "RegKey::GetBinaryValue failed.\n";
    }

    if(auto testBinary2 = key.TryGetBinaryValue("TestValueBinary"))
    {
        if(testBinary2 != testBinary)
        {
            std::cout << "RegKey::TryGetBinaryValue failed.\n";
        }
    }
    else
    {
        std::cout << "RegKey::TryGetBinaryValue failed (std::std::optional has no value).\n";
    }

    typeId = key.QueryValueType("TestValueBinary");
    if(typeId != REG_BINARY)
    {
        std::cout << "RegKey::QueryValueType failed for REG_BINARY.\n";
    }

    //
    // Remove some test values
    //

    key.DeleteValue("TestValueDword");
    key.DeleteValue("TestValueQword");
    key.DeleteValue("TestValueString");
    key.DeleteValue("TestValueExpandString");
    key.DeleteValue("TestValueMultiString");
    key.DeleteValue("TestValueBinary");
}

int main()
{
    const int kExitOk = 0;
    const int kExitError = 1;

    try
    {
        std::cout << "=========================================\n";
        std::cout << "*** Testing Giovanni Dicanio's WinReg ***\n";
        std::cout << "=========================================\n\n";

        Test();

        std::cout << "All right!! :)\n\n";
    }
    catch(const RegException& e)
    {
        std::cout << "\n*** Registry Exception: " << e.what();
        std::cout << "\n*** [Windows API error code = " << e.code() << "]\n\n";
        return kExitError;
    }
    catch(const std::exception& e)
    {
        std::cout << "\n*** ERROR: " << e.what() << L'\n';
        return kExitError;
    }

    return kExitOk;
}

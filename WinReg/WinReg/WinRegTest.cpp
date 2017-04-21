//////////////////////////////////////////////////////////////////////////
//
// WinRegTest.cpp -- by Giovanni Dicanio
// 
// Test some of the code in WinReg.hpp
// 
////////////////////////////////////////////////////////////////////////// 

#include "WinReg.hpp"   // Module to test
#include <exception>
#include <iostream>
#include <vector>
using namespace std;

int main()
{
    constexpr int kExitOk = 0;
    constexpr int kExitError = 1;

    try 
    {
        wcout << L"\n*** Testing Giovanni Dicanio's WinReg.hpp ***\n\n";

        //
        // Test subkey and value enumeration
        // 

        const wstring testSubKey = L"SOFTWARE\\GioTest";
        winreg::RegKey key{ HKEY_CURRENT_USER, testSubKey };    
        
        vector<wstring> subKeyNames = winreg::EnumSubKeys(key.Get());
        wcout << L"\nSubkeys:\n";
        for (const auto& s : subKeyNames)
        {
            wcout << L"  [" << s << L"]\n";
        }
        wcout << L'\n';

        vector<pair<wstring, DWORD>> values = winreg::EnumValues(key.Get());
        wcout << L"\nValues:\n";
        for (const auto& v : values)
        {
            wcout << L"  [" << v.first << L"](" << winreg::RegTypeToString(v.second) << L")\n";
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
        const vector<BYTE> testBinary{0xAA, 0xBB, 0xCC, 0x11, 0x22, 0x33};
        const vector<wstring> testMultiSz{ L"Hi", L"Hello", L"Ciao" };

        winreg::SetDwordValue(key.Get(), L"TestValueDword", testDw);
        winreg::SetQwordValue(key.Get(), L"TestValueQword", testQw);
        winreg::SetStringValue(key.Get(), L"TestValueString", testSz);
        winreg::SetExpandStringValue(key.Get(), L"TestValueExpandString", testExpandSz);
        winreg::SetMultiStringValue(key.Get(), L"TestValueMultiString", testMultiSz);
        winreg::SetBinaryValue(key.Get(), L"TestValueBinary", testBinary);

        DWORD testDw1 = winreg::GetDwordValue(key.Get(), L"TestValueDword");
        if (testDw1 != testDw) 
        {
            wcout << "winreg::GetDwordValue failed.\n";
        }

        DWORD typeId = winreg::QueryValueType(key.Get(), L"TestValueDword");
        if (typeId != REG_DWORD)
        {
            wcout << "winreg::QueryValueType failed for REG_DWORD.\n";
        }

        ULONGLONG testQw1 = winreg::GetQwordValue(key.Get(), L"TestValueQword");
        if (testQw1 != testQw)
        {
            wcout << "winreg::GetQwordValue failed.\n";
        }

        typeId = winreg::QueryValueType(key.Get(), L"TestValueQword");
        if (typeId != REG_QWORD)
        {
            wcout << "winreg::QueryValueType failed for REG_QWORD.\n";
        }

        wstring testSz1 = winreg::GetStringValue(key.Get(), L"TestValueString");
        if (testSz1 != testSz) 
        {
            wcout << "winreg::GetStringValue failed.\n";
        }

        typeId = winreg::QueryValueType(key.Get(), L"TestValueString");
        if (typeId != REG_SZ)
        {
            wcout << "winreg::QueryValueType failed for REG_SZ.\n";
        }

        wstring testExpandSz1 = winreg::GetExpandStringValue(key.Get(), L"TestValueExpandString");
        if (testExpandSz1 != testExpandSz)
        {
            wcout << "winreg::GetExpandStringValue failed.\n";
        }

        typeId = winreg::QueryValueType(key.Get(), L"TestValueExpandString");
        if (typeId != REG_EXPAND_SZ)
        {
            // NOTE: This seems a bug in the RegGetValue API, *not* in my wrapper function.
            // In fact, I tried direct RegGetValue calls for REG_EXPAND_SZ values,
            // and the API returns REG_SZ instead :(
            wcout << "winreg::QueryValueType failed for REG_EXPAND_SZ.\n";
        }

        vector<wstring> testMultiSz1 = winreg::GetMultiStringValue(key.Get(), L"TestValueMultiString");
        if (testMultiSz1 != testMultiSz)
        {
            wcout << "winreg::GetMultiStringValue failed.\n";
        }

        typeId = winreg::QueryValueType(key.Get(), L"TestValueMultiString");
        if (typeId != REG_MULTI_SZ)
        {
            wcout << "winreg::QueryValueType failed for REG_MULTI_SZ.\n";
        }

        vector<BYTE> testBinary1 = winreg::GetBinaryValue(key.Get(), L"TestValueBinary");
        if (testBinary1 != testBinary)
        {
            wcout << "winreg::GetBinaryValue failed.\n";
        }

        typeId = winreg::QueryValueType(key.Get(), L"TestValueBinary");
        if (typeId != REG_BINARY)
        {
            wcout << "winreg::QueryValueType failed for REG_BINARY.\n";
        }


        // Remove some test values
        winreg::DeleteValue(key.Get(), L"TestValueDword");
        winreg::DeleteValue(key.Get(), L"TestValueQword");
        winreg::DeleteValue(key.Get(), L"TestValueString");
        winreg::DeleteValue(key.Get(), L"TestValueExpandString");
        winreg::DeleteValue(key.Get(), L"TestValueMultiString");
        winreg::DeleteValue(key.Get(), L"TestValueBinary");
    }
    catch (const winreg::RegException& e)
    {
        cout << "\n*** Registry Exception: " << e.what();
        cout << "\n*** [Windows API error code = " << e.ErrorCode() << "\n\n";
        return kExitError;
    }
    catch (const exception& e)
    {
        cout << "\n*** ERROR: " << e.what() << '\n';
        return kExitError;
    }

    return kExitOk;
}

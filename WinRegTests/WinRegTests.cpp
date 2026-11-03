////////////////////////////////////////////////////////////////////////////////
// FILE: WinRegTests.cpp
// DESC: This file contains test cases for WinReg.
// AUTHOR: Giovanni Dicanio
////////////////////////////////////////////////////////////////////////////////


//==============================================================================
// There is no 'main' function here, as Catch2's catch_amalgamated.cpp
// already provides main().
//==============================================================================


//------------------------------------------------------------------------------
// The tests here use the Catch2 framework (v3.16.1)
//
// Catch2 GitHub link:
// https://github.com/catchorg/Catch2
//
// Release link:
// https://github.com/catchorg/Catch2/releases/tag/v3.16.1
//
//------------------------------------------------------------------------------


#include "catch_amalgamated.hpp"            // Catch2's amalgamated header

#include "WinReg/WinReg.hpp"                // WinReg's main header

#include <Windows.h>                        // Windows Native Platform API (Win32)

#include <string>
#include <system_error>
#include <vector>


// If this macro is defined, additional checks on the *exact* error return codes
// are executed.
#define WINREG_TESTS_CHECK_ERROR_CODES


//=============================================================================
//                          Test Infrastructure
//=============================================================================

//-----------------------------------------------------------------------------
// Helper macro to make sure that Win32 API calls succeed.
//-----------------------------------------------------------------------------
#define REQUIRE_WIN32_SUCCESS(expr)                             \
    do                                                          \
    {                                                           \
        const LSTATUS win32Ret_ = (expr);                       \
        INFO(#expr << "\n  returned " << win32Ret_ << ": "      \
             << std::system_category().message(win32Ret_));     \
        REQUIRE(win32Ret_ == ERROR_SUCCESS);                    \
    }                                                           \
    while (0)


//-----------------------------------------------------------------------------
// Using a test fixture:
// Catch2's TEST_CASE_METHOD builds a fresh instance for every test case,
// so each test starts clean.
// Moreover, if REQUIRE fails, Catch2 throws an exception; this will unwind
// the stack and run the destructor of the test fixture, so cleanup happens
// even when a test fails.
//-----------------------------------------------------------------------------
class RegTestFixture
{
public:

    // Init stuff
    RegTestFixture()
    {
        // Remove leftovers from a previous run that crashed or was killed
        // (destructors don't run in that case)
        DeleteTestKey();

        // (Re)create the test key
        m_key.Create(HKEY_CURRENT_USER, kTestKeyPath);
    }

    // Cleanup stuff
    ~RegTestFixture()
    {
        // Close our handle first
        m_key.Close();

        // Then delete the whole subtree
        DeleteTestKey();
    }

    //
    // We own the registry resource: disable copy
    //
    RegTestFixture(RegTestFixture const&) = delete;
    RegTestFixture& operator=(RegTestFixture const&) = delete;


protected:

    static constexpr const wchar_t* kTestKeyPath = L"Software\\GD_WinRegTests";

    winreg::RegKey& Key() noexcept { return m_key; }


private:

    static void DeleteTestKey() noexcept
    {
        // Raw Win32 API call on purpose: it reports errors through return codes,
        // so it can't throw from a destructor.
        // ERROR_FILE_NOT_FOUND (nothing to delete) is fine and ignored.
        ::RegDeleteTreeW(HKEY_CURRENT_USER, kTestKeyPath);
    }

    winreg::RegKey m_key;
};


//=============================================================================
//                          DWORD Tests
//=============================================================================

TEST_CASE_METHOD(RegTestFixture, "DWORD values round-trip", "[RegKey][REG_DWORD]")
{
    Key().SetDwordValue(L"TestDword", 64);
    REQUIRE(Key().GetDwordValue(L"TestDword") == 64);
}


TEST_CASE_METHOD(RegTestFixture, "Missing DWORD value throws", "[RegKey][REG_DWORD]")
{
    REQUIRE_THROWS_AS(Key().GetDwordValue(L"DoesNotExist"), winreg::RegException);
}


TEST_CASE_METHOD(RegTestFixture, "Malformed DWORD value throws", "[RegKey][REG_DWORD]")
{
    SECTION("DWORD with cb = 8 (ULONGLONG)")
    {
        // Write a malformed DWORD
        //
        // Create a REG_DWORD with cb = 8 (ULONLONG)
        //
        ULONGLONG testValue = 100;

        DWORD cbData = sizeof(testValue); // should be 4 for a DWORD...

        const wchar_t* valueName = L"TestWrongType_Dword_cb8";

        REQUIRE_WIN32_SUCCESS(::RegSetValueEx(
            Key().Get(),
            valueName,
            0,
            REG_DWORD,
            reinterpret_cast<const BYTE*>(&testValue),
            cbData
        ));

        REQUIRE_THROWS_AS(Key().GetDwordValue(valueName), winreg::RegException);

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        try
        {
            [[maybe_unused]] auto result = Key().GetDwordValue(valueName);
            FAIL("Expected winreg::RegException");
        }
        catch (const winreg::RegException& e)
        {
            REQUIRE(e.code().value() == ERROR_INVALID_DATA);
        }
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }


    SECTION("DWORD with cb = 2")
    {
        // Write a malformed DWORD
        //
        // Create a REG_DWORD with cb = 2

        // Store a REG_DWORD with only 2 bytes: 0x1234
        const WORD shortValue = 0x1234;

        DWORD cbData = sizeof(shortValue); // should be 4 for a DWORD...

        const wchar_t* valueName = L"TestWrongType_Dword_cb2";

        REQUIRE_WIN32_SUCCESS(::RegSetValueEx(
            Key().Get(),
            valueName,
            0,
            REG_DWORD,
            reinterpret_cast<const BYTE*>(&shortValue),
            cbData
        ));

        REQUIRE_THROWS_AS(Key().GetDwordValue(valueName), winreg::RegException);

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        try
        {
            [[maybe_unused]] auto result = Key().GetDwordValue(valueName);
            FAIL("Expected winreg::RegException");
        }
        catch (const winreg::RegException& e)
        {
            REQUIRE(e.code().value() == ERROR_INVALID_DATA);
        }
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }

    SECTION("another type read as a DWORD")
    {
        // Create a REG_SZ, and try to read it back as a DWORD

        constexpr wchar_t testString[] = L"Hello world! This is a test.";

        const wchar_t* valueName = L"TestWrongType_Dword_RegSZ";

        REQUIRE_WIN32_SUCCESS(::RegSetValueEx(
            Key().Get(),
            valueName,
            0,
            REG_SZ,
            reinterpret_cast<const BYTE*>(testString),
            sizeof(testString)
        ));

        REQUIRE_THROWS_AS(Key().GetDwordValue(valueName), winreg::RegException);

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        try
        {
            [[maybe_unused]] auto result = Key().GetDwordValue(valueName);
            FAIL("Expected winreg::RegException");
        }
        catch (const winreg::RegException& e)
        {
            REQUIRE(e.code().value() == ERROR_UNSUPPORTED_TYPE);
        }
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }
}


//=============================================================================
//                          QWORD Tests
//=============================================================================

TEST_CASE_METHOD(RegTestFixture, "QWORD values round-trip", "[RegKey][REG_QWORD]")
{
    constexpr ULONGLONG kTestValue = 0x11223344'AABBCCDD;
    Key().SetQwordValue(L"TestQword", kTestValue);
    REQUIRE(Key().GetQwordValue(L"TestQword") == kTestValue);
}


TEST_CASE_METHOD(RegTestFixture, "Missing QWORD value throws", "[RegKey][REG_QWORD]")
{
    REQUIRE_THROWS_AS(Key().GetQwordValue(L"DoesNotExist"), winreg::RegException);
}


TEST_CASE_METHOD(RegTestFixture, "Malformed QWORD value throws", "[RegKey][REG_QWORD]")
{
    SECTION("QWORD with cb = 16")
    {
        // Write a malformed QWORD
        //
        // Create a REG_QWORD with cb = 16
        //
        BYTE testValue[16] = { 0 };

        DWORD cbData = sizeof(testValue); // should be 8 for a QWORD...

        const wchar_t* valueName = L"TestWrongType_Qword_cb16";

        REQUIRE_WIN32_SUCCESS(::RegSetValueEx(
            Key().Get(),
            valueName,
            0,
            REG_QWORD,
            reinterpret_cast<const BYTE*>(&testValue),
            cbData
        ));

        REQUIRE_THROWS_AS(Key().GetQwordValue(valueName), winreg::RegException);

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        try
        {
            [[maybe_unused]] auto result = Key().GetQwordValue(valueName);
            FAIL("Expected winreg::RegException");
        }
        catch (const winreg::RegException& e)
        {
            REQUIRE(e.code().value() == ERROR_INVALID_DATA);
        }
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }

    SECTION("QWORD with cb = 4")
    {
        // Write a malformed QWORD
        //
        // Create a REG_QWORD with cb = 4

        // Store a REG_QWORD with only 4 bytes
        const DWORD shortValue = 0x1234ABCD;

        DWORD cbData = sizeof(shortValue); // should be 8 for a QWORD...

        const wchar_t* valueName = L"TestWrongType_Qword_cb4";

        REQUIRE_WIN32_SUCCESS(::RegSetValueEx(
            Key().Get(),
            valueName,
            0,
            REG_QWORD,
            reinterpret_cast<const BYTE*>(&shortValue),
            cbData
        ));

        REQUIRE_THROWS_AS(Key().GetQwordValue(valueName), winreg::RegException);

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        try
        {
            [[maybe_unused]] auto result = Key().GetQwordValue(valueName);
            FAIL("Expected winreg::RegException");
        }
        catch (const winreg::RegException& e)
        {
            REQUIRE(e.code().value() == ERROR_INVALID_DATA);
        }
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }

    SECTION("another type read as a QWORD")
    {
        // Create a REG_SZ, and try to read it back as a QWORD

        constexpr wchar_t testString[] = L"Hello world! This is a test.";

        const wchar_t* valueName = L"TestWrongType_Qword_RegSZ";

        REQUIRE_WIN32_SUCCESS(::RegSetValueEx(
            Key().Get(),
            valueName,
            0,
            REG_SZ,
            reinterpret_cast<const BYTE*>(testString),
            sizeof(testString)
        ));

        REQUIRE_THROWS_AS(Key().GetQwordValue(valueName), winreg::RegException);

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        try
        {
            [[maybe_unused]] auto result = Key().GetQwordValue(valueName);
            FAIL("Expected winreg::RegException");
        }
        catch (const winreg::RegException& e)
        {
            REQUIRE(e.code().value() == ERROR_UNSUPPORTED_TYPE);
        }
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }
}


//=============================================================================
//                          String Tests
//=============================================================================

TEST_CASE_METHOD(RegTestFixture, "String values", "[RegKey][REG_SZ]")
{
    SECTION("plain string")
    {
        Key().SetStringValue(L"Name", L"Connie");
        REQUIRE(Key().GetStringValue(L"Name") == L"Connie");
    }

    SECTION("empty string")
    {
        Key().SetStringValue(L"Empty", L"");
        REQUIRE(Key().GetStringValue(L"Empty").empty());
    }
}


TEST_CASE_METHOD(RegTestFixture, "Missing string value throws", "[RegKey][REG_SZ]")
{
    REQUIRE_THROWS_AS(Key().GetStringValue(L"DoesNotExist"), winreg::RegException);
}


TEST_CASE_METHOD(RegTestFixture, "Malformed strings throw", "[RegKey][REG_SZ]")
{
    // Odd byte length
    SECTION("odd byte length string")
    {
        const BYTE data[] = {'A', 'b', 'c', 'd', '\0'};

        const wchar_t* valueName = L"OddByteLengthString";

        REQUIRE_WIN32_SUCCESS(::RegSetValueExW(
            Key().Get(),
            valueName,
            0,
            REG_SZ,
            data,
            sizeof(data)
        ));

        REQUIRE_THROWS_AS(Key().GetStringValue(valueName), winreg::RegException);

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        try
        {
            [[maybe_unused]] auto result = Key().GetStringValue(valueName);
            FAIL("Expected winreg::RegException");
        }
        catch (const winreg::RegException& e)
        {
            REQUIRE(e.code().value() == ERROR_INVALID_DATA);
        }
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }

    // Not REG_SZ
    SECTION("not REG_SZ string")
    {
        const DWORD data = 0xABCDEF;

        const wchar_t* valueName = L"Not_REG_SZ";

        REQUIRE_WIN32_SUCCESS( ::RegSetValueExW(
            Key().Get(),
            valueName,
            0,
            REG_DWORD,
            reinterpret_cast<const BYTE*>(&data),
            sizeof(data)
        ));

        REQUIRE_THROWS_AS(Key().GetStringValue(valueName), winreg::RegException);

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        try
        {
            [[maybe_unused]] auto result = Key().GetStringValue(valueName);
            FAIL("Expected winreg::RegException");
        }
        catch (const winreg::RegException& e)
        {
            REQUIRE(e.code().value() == ERROR_UNSUPPORTED_TYPE);
        }
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }
}


//=============================================================================
//                          Multi-String Tests
//=============================================================================

TEST_CASE_METHOD(RegTestFixture, "Multi-String values", "[RegKey][REG_MULTI_SZ]")
{
    SECTION("plain multi-string")
    {
        const std::vector<std::wstring> names = { L"Connie", L"Gio", L"C64" };
        Key().SetMultiStringValue(L"Names", names);
        REQUIRE(Key().GetMultiStringValue(L"Names") == names);
    }

    SECTION("empty multi-string")
    {
        Key().SetMultiStringValue(L"Empty", {});
        REQUIRE(Key().GetMultiStringValue(L"Empty").empty());
    }
}


TEST_CASE_METHOD(RegTestFixture, "Missing multi-string value throws", "[RegKey][REG_MULTI_SZ]")
{
    REQUIRE_THROWS_AS(Key().GetMultiStringValue(L"DoesNotExist"), winreg::RegException);
}


//=============================================================================
//                          Binary Data Tests
//=============================================================================

TEST_CASE_METHOD(RegTestFixture, "Binary values", "[RegKey][REG_BINARY]")
{
    SECTION("plain binary data")
    {
        const std::vector<BYTE> data = { 0x11, 0x22, 0x33, 0x44 };
        Key().SetBinaryValue(L"TestBinaryData", data);
        REQUIRE(Key().GetBinaryValue(L"TestBinaryData") == data);
    }

    SECTION("empty binary")
    {
        Key().SetBinaryValue(L"Empty", std::vector<BYTE>{});
        REQUIRE(Key().GetBinaryValue(L"Empty").empty());
    }
}


TEST_CASE_METHOD(RegTestFixture, "Missing binary value throws", "[RegKey][REG_BINARY]")
{
    REQUIRE_THROWS_AS(Key().GetBinaryValue(L"DoesNotExist"), winreg::RegException);
}


//=============================================================================
//                          Enum Value Tests
//=============================================================================

TEST_CASE_METHOD(RegTestFixture, "Enum values", "[RegKey][EnumValues]")
{
    //
    // Create some values, and enumerate them back
    //

    Key().SetDwordValue(L"Value1_Dword", 10);
    Key().SetDwordValue(L"Value2_Dword", 20);
    Key().SetDwordValue(L"Value3_Dword", 0xABCD'EF12);
    Key().SetQwordValue(L"Value4_Qword", 0xABCD'EF12'3456'0102);
    Key().SetStringValue(L"Value5_String", L"Test string #1");
    Key().SetStringValue(L"Value6_String", L"Test string #2");

    // It's important that the elements in this expected values array are sorted by value name
    const std::vector<std::pair<std::wstring, DWORD>> expectedValues = {
        {L"Value1_Dword",   REG_DWORD},
        {L"Value2_Dword",   REG_DWORD},
        {L"Value3_Dword",   REG_DWORD},
        {L"Value4_Qword",   REG_QWORD},
        {L"Value5_String",  REG_SZ},
        {L"Value6_String",  REG_SZ}
    };

    auto values = Key().EnumValues();

    REQUIRE(values.size() == expectedValues.size());

    // Make sure the elements are sorted before comparing element-wise
    std::sort(values.begin(), values.end());

    int index = 0;
    for (const auto& [valueName, valueType] : values)
    {
        REQUIRE(valueName == expectedValues[index].first);
        REQUIRE(valueType == expectedValues[index].second);

        index++;
    }
}


//=============================================================================
//                          Contains Tests
//=============================================================================

TEST_CASE_METHOD(RegTestFixture, "Key contains value", "[RegKey][ContainsValue]")
{
    //
    // Create some values for testing
    //
    Key().SetDwordValue(L"TestDword", 10);
    Key().SetStringValue(L"Name", L"Connie");

    // Check if existing values are contained in the key
    REQUIRE(Key().ContainsValue(L"TestDword"));
    REQUIRE(Key().ContainsValue(L"Name"));

    // Check that a not-existing value is not contained in the key
    REQUIRE(! Key().ContainsValue(L"ThisDoesNotExist123"));
}



//=============================================================================
//                          Testing TryXxx Methods
//=============================================================================


//=============================================================================
//                          DWORD Tests [try]
//=============================================================================

TEST_CASE_METHOD(RegTestFixture, "DWORD values round-trip", "[RegKey][REG_DWORD][try]")
{
    auto result = Key().TrySetDwordValue(L"TestDword", 64);
    REQUIRE(result.IsOk());

    auto resultRead = Key().TryGetDwordValue(L"TestDword");
    REQUIRE(resultRead.has_value());
    REQUIRE(resultRead.value() == 64);
}


TEST_CASE_METHOD(RegTestFixture, "Missing DWORD value returns error", "[RegKey][REG_DWORD][try]")
{
    auto result = Key().TryGetDwordValue(L"DoesNotExist");
    REQUIRE(! result.has_value());
}


TEST_CASE_METHOD(RegTestFixture, "Malformed DWORD value returns error", "[RegKey][REG_DWORD][try]")
{
    SECTION("DWORD with cb = 8 (ULONGLONG)")
    {
        // Write a malformed DWORD
        //
        // Create a REG_DWORD with cb = 8 (ULONLONG)
        //
        ULONGLONG testValue = 100;

        DWORD cbData = sizeof(testValue); // should be 4 for a DWORD...

        const wchar_t* valueName = L"TestWrongType_Dword_cb8";

        REQUIRE_WIN32_SUCCESS(::RegSetValueEx(
            Key().Get(),
            valueName,
            0,
            REG_DWORD,
            reinterpret_cast<const BYTE*>(&testValue),
            cbData
        ));

        auto result = Key().TryGetDwordValue(valueName);
        REQUIRE(!result.has_value());

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        REQUIRE(result.error().Code() == ERROR_INVALID_DATA);
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }


    SECTION("DWORD with cb = 2")
    {
        // Write a malformed DWORD
        //
        // Create a REG_DWORD with cb = 2

        // Store a REG_DWORD with only 2 bytes: 0x1234
        const WORD shortValue = 0x1234;

        DWORD cbData = sizeof(shortValue); // should be 4 for a DWORD...

        const wchar_t* valueName = L"TestWrongType_Dword_cb2";

        REQUIRE_WIN32_SUCCESS(::RegSetValueEx(
            Key().Get(),
            valueName,
            0,
            REG_DWORD,
            reinterpret_cast<const BYTE*>(&shortValue),
            cbData
        ));

        auto result = Key().TryGetDwordValue(valueName);
        REQUIRE(!result.has_value());

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        REQUIRE(result.error().Code() == ERROR_INVALID_DATA);
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }

    SECTION("another type read as a DWORD")
    {
        // Create a REG_SZ, and try to read it back as a DWORD

        constexpr wchar_t testString[] = L"Hello world! This is a test.";

        const wchar_t* valueName = L"TestWrongType_Dword_RegSZ";

        REQUIRE_WIN32_SUCCESS(::RegSetValueEx(
            Key().Get(),
            valueName,
            0,
            REG_SZ,
            reinterpret_cast<const BYTE*>(testString),
            sizeof(testString)
        ));

        auto result = Key().TryGetDwordValue(valueName);
        REQUIRE(!result.has_value());

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        REQUIRE(result.error().Code() == ERROR_UNSUPPORTED_TYPE);
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }
}


//=============================================================================
//                          QWORD Tests [try]
//=============================================================================

TEST_CASE_METHOD(RegTestFixture, "QWORD values round-trip", "[RegKey][REG_QWORD][try]")
{
    constexpr ULONGLONG kTestValue = 0x11223344'AABBCCDD;

    auto result = Key().TrySetQwordValue(L"TestQword", kTestValue);
    REQUIRE(result.IsOk());

    auto resultRead = Key().TryGetQwordValue(L"TestQword");
    REQUIRE(resultRead.has_value());
    REQUIRE(resultRead.value() == kTestValue);
}


TEST_CASE_METHOD(RegTestFixture, "Missing QWORD value returns error", "[RegKey][REG_QWORD][try]")
{
    auto result = Key().TryGetQwordValue(L"DoesNotExist");
    REQUIRE(! result.has_value());
}


TEST_CASE_METHOD(RegTestFixture, "Malformed QWORD value returns error", "[RegKey][REG_QWORD][try]")
{
    SECTION("QWORD with cb = 16")
    {
        // Write a malformed QWORD
        //
        // Create a REG_QWORD with cb = 16
        //
        BYTE testValue[16] = { 0 };

        DWORD cbData = sizeof(testValue); // should be 8 for a QWORD...

        const wchar_t* valueName = L"TestWrongType_Qword_cb16";

        REQUIRE_WIN32_SUCCESS(::RegSetValueEx(
            Key().Get(),
            valueName,
            0,
            REG_QWORD,
            reinterpret_cast<const BYTE*>(&testValue),
            cbData
        ));

        auto result = Key().TryGetQwordValue(valueName);
        REQUIRE(!result.has_value());

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        REQUIRE(result.error().Code() == ERROR_INVALID_DATA);
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }

    SECTION("QWORD with cb = 4")
    {
        // Write a malformed QWORD
        //
        // Create a REG_QWORD with cb = 4

        // Store a REG_QWORD with only 4 bytes
        const DWORD shortValue = 0x1234ABCD;

        DWORD cbData = sizeof(shortValue); // should be 8 for a QWORD...

        const wchar_t* valueName = L"TestWrongType_Qword_cb4";

        REQUIRE_WIN32_SUCCESS(::RegSetValueEx(
            Key().Get(),
            valueName,
            0,
            REG_QWORD,
            reinterpret_cast<const BYTE*>(&shortValue),
            cbData
        ));

        auto result = Key().TryGetQwordValue(valueName);
        REQUIRE(! result.has_value());

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        REQUIRE(result.error().Code() == ERROR_INVALID_DATA);
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }

    SECTION("another type read as a QWORD")
    {
        // Create a REG_SZ, and try to read it back as a QWORD

        constexpr wchar_t testString[] = L"Hello world! This is a test.";

        const wchar_t* valueName = L"TestWrongType_Qword_RegSZ";

        REQUIRE_WIN32_SUCCESS(::RegSetValueEx(
            Key().Get(),
            valueName,
            0,
            REG_SZ,
            reinterpret_cast<const BYTE*>(testString),
            sizeof(testString)
        ));

        auto result = Key().TryGetQwordValue(valueName);
        REQUIRE(!result.has_value());

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        REQUIRE(result.error().Code() == ERROR_UNSUPPORTED_TYPE);
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }
}


//=============================================================================
//                          String Tests [try]
//=============================================================================

TEST_CASE_METHOD(RegTestFixture, "String values", "[RegKey][REG_SZ][try]")
{
    SECTION("plain string")
    {
        auto result = Key().TrySetStringValue(L"Name", L"Connie");
        REQUIRE(result.IsOk());

        auto resultRead = Key().TryGetStringValue(L"Name");
        REQUIRE(resultRead.has_value());
        REQUIRE(resultRead.value() == L"Connie");
    }

    SECTION("empty string")
    {
        auto result = Key().TrySetStringValue(L"Empty", L"");
        REQUIRE(result.IsOk());

        auto resultRead = Key().TryGetStringValue(L"Empty");
        REQUIRE(resultRead.has_value());
        REQUIRE(resultRead.value().empty());
    }
}


TEST_CASE_METHOD(RegTestFixture, "Missing string value returns error", "[RegKey][REG_SZ][try]")
{
    auto result = Key().TryGetStringValue(L"DoesNotExist");
    REQUIRE(!result.has_value());
}


TEST_CASE_METHOD(RegTestFixture, "Malformed strings return error", "[RegKey][REG_SZ][try]")
{
    // Odd byte length
    SECTION("odd byte length string")
    {
        const BYTE data[] = { 'A', 'b', 'c', 'd', '\0' };

        const wchar_t* valueName = L"OddByteLengthString";

        REQUIRE_WIN32_SUCCESS(::RegSetValueExW(
            Key().Get(),
            valueName,
            0,
            REG_SZ,
            data,
            sizeof(data)
        ));

        auto result = Key().TryGetStringValue(valueName);
        REQUIRE(! result.has_value());

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        REQUIRE(result.error().Code() == ERROR_INVALID_DATA);
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }

    // Not REG_SZ
    SECTION("not REG_SZ string")
    {
        const DWORD data = 0xABCDEF;

        const wchar_t* valueName = L"Not_REG_SZ";

        REQUIRE_WIN32_SUCCESS(::RegSetValueExW(
            Key().Get(),
            valueName,
            0,
            REG_DWORD,
            reinterpret_cast<const BYTE*>(&data),
            sizeof(data)
        ));

        auto result = Key().TryGetStringValue(valueName);
        REQUIRE(!result.has_value());

#ifdef WINREG_TESTS_CHECK_ERROR_CODES
        REQUIRE(result.error().Code() == ERROR_UNSUPPORTED_TYPE);
#endif // WINREG_TESTS_CHECK_ERROR_CODES

    }
}


//=============================================================================
//                          Multi-String Tests [try]
//=============================================================================

TEST_CASE_METHOD(RegTestFixture, "Multi-String values", "[RegKey][REG_MULTI_SZ][try]")
{
    SECTION("plain multi-string")
    {
        const std::vector<std::wstring> names = { L"Connie", L"Gio", L"C64" };
        auto result = Key().TrySetMultiStringValue(L"Names", names);
        REQUIRE(result.IsOk());

        auto resultRead = Key().TryGetMultiStringValue(L"Names");
        REQUIRE(resultRead.has_value());
        REQUIRE(resultRead.value() == names);
    }

    SECTION("empty multi-string")
    {
        auto result = Key().TrySetMultiStringValue(L"Empty", {});
        REQUIRE(result.IsOk());

        auto resultRead = Key().TryGetMultiStringValue(L"Empty");
        REQUIRE(resultRead.has_value());
        REQUIRE(resultRead.value().empty());
    }
}


TEST_CASE_METHOD(RegTestFixture, "Missing multi-string value returns error", "[RegKey][REG_MULTI_SZ][try]")
{
    auto result = Key().TryGetMultiStringValue(L"DoesNotExist");
    REQUIRE(!result.has_value());
}


//=============================================================================
//                          Binary Data Tests [try]
//=============================================================================

TEST_CASE_METHOD(RegTestFixture, "Binary values", "[RegKey][REG_BINARY][try]")
{
    SECTION("plain binary data")
    {
        const std::vector<BYTE> data = { 0x11, 0x22, 0x33, 0x44 };
        auto result = Key().TrySetBinaryValue(L"TestBinaryData", data);
        REQUIRE(result.IsOk());

        auto resultRead = Key().TryGetBinaryValue(L"TestBinaryData");
        REQUIRE(resultRead.has_value());
        REQUIRE(resultRead.value() == data);
    }

    SECTION("empty binary")
    {
        auto result = Key().TrySetBinaryValue(L"Empty", std::vector<BYTE>{});
        REQUIRE(result.IsOk());

        auto resultRead = Key().TryGetBinaryValue(L"Empty");
        REQUIRE(resultRead.has_value());
        REQUIRE(resultRead.value().empty());
    }
}


TEST_CASE_METHOD(RegTestFixture, "Missing binary value returns error", "[RegKey][REG_BINARY][try]")
{
    auto result = Key().TryGetBinaryValue(L"DoesNotExist");
    REQUIRE(!result.has_value());
}


//=============================================================================
//                          Enum Value Tests [try]
//=============================================================================

TEST_CASE_METHOD(RegTestFixture, "Enum values", "[RegKey][EnumValues][try]")
{
    //
    // Create some values, and enumerate them back
    //

    // NOTE:
    // I'm assuming all the Set method invocations succeed,
    // so I'm using the throwing variants (SetXxxxValue).
    // If some of these fail, an exception will be thrown
    // (but that is not expected, and not what this test case is testing)
    Key().SetDwordValue(L"Value1_Dword", 10);
    Key().SetDwordValue(L"Value2_Dword", 20);
    Key().SetDwordValue(L"Value3_Dword", 0xABCD'EF12);
    Key().SetQwordValue(L"Value4_Qword", 0xABCD'EF12'3456'0102);
    Key().SetStringValue(L"Value5_String", L"Test string #1");
    Key().SetStringValue(L"Value6_String", L"Test string #2");

    // It's important that the elements in this expected values array are sorted by value name
    const std::vector<std::pair<std::wstring, DWORD>> expectedValues = {
        {L"Value1_Dword",   REG_DWORD},
        {L"Value2_Dword",   REG_DWORD},
        {L"Value3_Dword",   REG_DWORD},
        {L"Value4_Qword",   REG_QWORD},
        {L"Value5_String",  REG_SZ},
        {L"Value6_String",  REG_SZ}
    };

    auto result = Key().TryEnumValues();
    REQUIRE(result.has_value());            // TryEnumValus() call must succeed

    // The result vector and the expected vector must have the same number of elements
    REQUIRE(result.value().size() == expectedValues.size());

    // Make sure the elements are sorted before comparing element-wise
    std::sort(result.value().begin(), result.value().end());

    int index = 0;
    for (const auto& [valueName, valueType] : result.value())
    {
        REQUIRE(valueName == expectedValues[index].first);
        REQUIRE(valueType == expectedValues[index].second);

        index++;
    }
}


// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started:
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file

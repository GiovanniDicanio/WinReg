#ifndef GIOVANNI_DICANIO_WINREG_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_HPP_INCLUDED


////////////////////////////////////////////////////////////////////////////////
//
//      *** Modern C++ Wrappers Around Windows Registry C API ***
//
//               Copyright (C) by Giovanni Dicanio
//
// First version: 2017, January 22nd
// Last update:   2024, March 5th
//
// E-mail: <first name>.<last name> AT REMOVE_THIS gmail.com
//
// Registry key handles are safely and conveniently wrapped
// in the RegKey resource manager C++ class.
//
// Many methods are available in two forms:
//
// - One form that signals errors throwing exceptions
//   of class RegException (e.g. RegKey::Open)
//
// - Another form that returns RegResult objects (e.g. RegKey::TryOpen)
//
// In addition, there are also some methods named like TryGet...Value
// (e.g. TryGetDwordValue), that _try_ to perform the given query,
// and return a RegExpected object. On success, that object contains
// the value read from the registry. On failure, the returned RegExpected object
// contains a RegResult storing the return code from the Windows Registry API call.
//
// Unicode UTF-16 strings are represented using the std::wstring class;
// ATL's CString is not used, to avoid dependencies from ATL or MFC.
//
// Compiler: Visual Studio 2019
// C++ Language Standard: C++17 (/std:c++17)
// Code compiles cleanly at warning level 4 (/W4) on both 32-bit and 64-bit builds.
//
// Requires building in Unicode mode (which has been the default since VS2005).
//
// ===========================================================================
//
// The MIT License(MIT)
//
// Copyright(c) 2017-2024 by Giovanni Dicanio
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files(the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and / or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions :
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
//
////////////////////////////////////////////////////////////////////////////////


#include <Windows.h>        // Windows Platform SDK
#include <crtdbg.h>         // _ASSERTE

#include <memory>           // std::unique_ptr, std::make_unique
#include <stdexcept>        // std::invalid_argument
#include <string>           // std::wstring
#include <system_error>     // std::system_error
#include <utility>          // std::swap, std::pair, std::move
#include <vector>           // std::vector


//
// *** Windows XP Compatibility Mode ***
//
// Define the following macro to enable code that works with Windows XP,
// using APIs like RegQueryValueEx instead of the Vista+ RegGetValue.
//
//#define GIOVANNI_DICANIO_WINREG_WINDOWS_XP_COMPATIBLE   1


namespace winreg
{

//
// Forward Class Declarations
//

class RegException;
class RegResult;

template <typename T>
class RegExpected;


//
// Class Declarations
//

//------------------------------------------------------------------------------
//
// Safe, efficient and convenient C++ wrapper around HKEY registry key handles.
//
// This class is movable but not copyable.
//
// This class is designed to be very *efficient* and low-overhead, for example:
// non-throwing operations are carefully marked as noexcept, so the C++ compiler
// can emit optimized code.
//
// Moreover, this class just wraps a raw HKEY handle, without any
// shared-ownership overhead like in std::shared_ptr; you can think of this
// class kind of like a std::unique_ptr for HKEYs.
//
// The class is also swappable (defines a custom non-member swap);
// relational operators are properly overloaded as well.
//
//------------------------------------------------------------------------------
class RegKey
{
public:

    //
    // Construction/Destruction
    //

    // Initialize as an empty key handle
    RegKey() noexcept = default;

    // Take ownership of the input key handle
    explicit RegKey(HKEY hKey) noexcept;

    // Open the given registry key if it exists, else create a new key.
    // Uses default KEY_READ|KEY_WRITE|KEY_WOW64_64KEY access.
    // For finer grained control, call the Create() method overloads.
    // Throw RegException on failure.
    RegKey(HKEY hKeyParent, const std::wstring& subKey);

    // Open the given registry key if it exists, else create a new key.
    // Allow the caller to specify the desired access to the key
    // (e.g. KEY_READ|KEY_WOW64_64KEY for read-only access).
    // For finer grained control, call the Create() method overloads.
    // Throw RegException on failure.
    RegKey(HKEY hKeyParent, const std::wstring& subKey, REGSAM desiredAccess);


    // Take ownership of the input key handle.
    // The input key handle wrapper is reset to an empty state.
    RegKey(RegKey&& other) noexcept;

    // Move-assign from the input key handle.
    // Properly check against self-move-assign (which is safe and does nothing).
    RegKey& operator=(RegKey&& other) noexcept;

    // Ban copy
    RegKey(const RegKey&) = delete;
    RegKey& operator=(const RegKey&) = delete;

    // Safely close the wrapped key handle (if any)
    ~RegKey() noexcept;


    //
    // Properties
    //

    // Access the wrapped raw HKEY handle
    [[nodiscard]] HKEY Get() const noexcept;

    // Is the wrapped HKEY handle valid?
    [[nodiscard]] bool IsValid() const noexcept;

    // Same as IsValid(), but allow a short "if (regKey)" syntax
    [[nodiscard]] explicit operator bool() const noexcept;

    // Is the wrapped handle a predefined handle (e.g.HKEY_CURRENT_USER) ?
    [[nodiscard]] bool IsPredefined() const noexcept;


    //
    // Operations
    //

    // Close current HKEY handle.
    // If there's no valid handle, do nothing.
    // This method doesn't close predefined HKEY handles (e.g. HKEY_CURRENT_USER).
    void Close() noexcept;

    // Transfer ownership of current HKEY to the caller.
    // Note that the caller is responsible for closing the key handle!
    [[nodiscard]] HKEY Detach() noexcept;

    // Take ownership of the input HKEY handle.
    // Safely close any previously open handle.
    // Input key handle can be nullptr.
    void Attach(HKEY hKey) noexcept;

    // Non-throwing swap;
    // Note: There's also a non-member swap overload
    void SwapWith(RegKey& other) noexcept;


    //
    // Wrappers around Windows Registry APIs.
    // See the official MSDN documentation for these APIs for detailed explanations
    // of the wrapper method parameters.
    //

    //
    // NOTE on the KEY_WOW64_64KEY flag
    // ================================
    //
    // By default, a 32-bit application running on 64-bit Windows accesses the 32-bit registry view
    // and a 64-bit application accesses the 64-bit registry view.
    // Using this KEY_WOW64_64KEY flag, both 32-bit or 64-bit applications access the 64-bit
    // registry view.
    //
    // MSDN documentation:
    // https://docs.microsoft.com/en-us/windows/win32/winprog64/accessing-an-alternate-registry-view
    //
    // If you want to use the default Windows API behavior, don't OR (|) the KEY_WOW64_64KEY flag
    // when specifying the desired access (e.g. just pass KEY_READ | KEY_WRITE as the desired
    // access parameter).
    //

    // Wrapper around RegCreateKeyEx, that allows you to specify desired access
    void Create(
        HKEY hKeyParent,
        const std::wstring& subKey,
        REGSAM desiredAccess = KEY_READ | KEY_WRITE | KEY_WOW64_64KEY
    );

    // Wrapper around RegCreateKeyEx
    void Create(
        HKEY hKeyParent,
        const std::wstring& subKey,
        REGSAM desiredAccess,
        DWORD options,
        SECURITY_ATTRIBUTES* securityAttributes,
        DWORD* disposition
    );

    // Wrapper around RegOpenKeyEx
    void Open(
        HKEY hKeyParent,
        const std::wstring& subKey,
        REGSAM desiredAccess = KEY_READ | KEY_WRITE | KEY_WOW64_64KEY
    );

    // Wrapper around RegCreateKeyEx, that allows you to specify desired access
    [[nodiscard]] RegResult TryCreate(
        HKEY hKeyParent,
        const std::wstring& subKey,
        REGSAM desiredAccess = KEY_READ | KEY_WRITE | KEY_WOW64_64KEY
    ) noexcept;

    // Wrapper around RegCreateKeyEx
    [[nodiscard]] RegResult TryCreate(
        HKEY hKeyParent,
        const std::wstring& subKey,
        REGSAM desiredAccess,
        DWORD options,
        SECURITY_ATTRIBUTES* securityAttributes,
        DWORD* disposition
    ) noexcept;

    // Wrapper around RegOpenKeyEx
    [[nodiscard]] RegResult TryOpen(
        HKEY hKeyParent,
        const std::wstring& subKey,
        REGSAM desiredAccess = KEY_READ | KEY_WRITE | KEY_WOW64_64KEY
    ) noexcept;


    //
    // Registry Value Setters
    //

    void SetDwordValue(const std::wstring& valueName, DWORD data);
    void SetQwordValue(const std::wstring& valueName, const ULONGLONG& data);
    void SetStringValue(const std::wstring& valueName, const std::wstring& data);
    void SetExpandStringValue(const std::wstring& valueName, const std::wstring& data);
    void SetMultiStringValue(const std::wstring& valueName, const std::vector<std::wstring>& data);
    void SetBinaryValue(const std::wstring& valueName, const std::vector<BYTE>& data);
    void SetBinaryValue(const std::wstring& valueName, const void* data, DWORD dataSize);


    //
    // Registry Value Setters Returning RegResult
    // (instead of throwing RegException on error)
    //

    [[nodiscard]] RegResult TrySetDwordValue(const std::wstring& valueName, DWORD data) noexcept;

    [[nodiscard]] RegResult TrySetQwordValue(const std::wstring& valueName,
                                             const ULONGLONG& data) noexcept;

    [[nodiscard]] RegResult TrySetStringValue(const std::wstring& valueName,
                                              const std::wstring& data) noexcept;

    [[nodiscard]] RegResult TrySetExpandStringValue(const std::wstring& valueName,
                                                    const std::wstring& data) noexcept;

    [[nodiscard]] RegResult TrySetMultiStringValue(const std::wstring& valueName,
                                                   const std::vector<std::wstring>& data);
    // Note: The TrySetMultiStringValue method CANNOT be marked noexcept,
    // because internally the method *dynamically allocates memory* for creating the multi-string
    // that will be stored in the Registry.

    [[nodiscard]] RegResult TrySetBinaryValue(const std::wstring& valueName,
                                              const std::vector<BYTE>& data) noexcept;

    [[nodiscard]] RegResult TrySetBinaryValue(const std::wstring& valueName,
                                              const void* data,
                                              DWORD dataSize) noexcept;


    //
    // Registry Value Getters
    //

    [[nodiscard]] DWORD GetDwordValue(const std::wstring& valueName) const;
    [[nodiscard]] ULONGLONG GetQwordValue(const std::wstring& valueName) const;
    [[nodiscard]] std::wstring GetStringValue(const std::wstring& valueName) const;

    enum class ExpandStringOption
    {
        DontExpand,
        Expand
    };

    [[nodiscard]] std::wstring GetExpandStringValue(
        const std::wstring& valueName,
        ExpandStringOption expandOption = ExpandStringOption::DontExpand
    ) const;

    [[nodiscard]] std::vector<std::wstring> GetMultiStringValue(const std::wstring& valueName) const;
    [[nodiscard]] std::vector<BYTE> GetBinaryValue(const std::wstring& valueName) const;


    //
    // Registry Value Getters Returning RegExpected<T>
    // (instead of throwing RegException on error)
    //

    [[nodiscard]] RegExpected<DWORD> TryGetDwordValue(const std::wstring& valueName) const;
    [[nodiscard]] RegExpected<ULONGLONG> TryGetQwordValue(const std::wstring& valueName) const;
    [[nodiscard]] RegExpected<std::wstring> TryGetStringValue(const std::wstring& valueName) const;

    [[nodiscard]] RegExpected<std::wstring> TryGetExpandStringValue(
        const std::wstring& valueName,
        ExpandStringOption expandOption = ExpandStringOption::DontExpand
    ) const;

    [[nodiscard]] RegExpected<std::vector<std::wstring>>
            TryGetMultiStringValue(const std::wstring& valueName) const;

    [[nodiscard]] RegExpected<std::vector<BYTE>>
            TryGetBinaryValue(const std::wstring& valueName) const;


    //
    // Query Operations
    //

    // Information about a registry key (retrieved by QueryInfoKey)
    struct InfoKey
    {
        DWORD    NumberOfSubKeys;
        DWORD    NumberOfValues;
        FILETIME LastWriteTime;

        // Clear the structure fields
        InfoKey() noexcept
            : NumberOfSubKeys{0}
            , NumberOfValues{0}
        {
            LastWriteTime.dwHighDateTime = LastWriteTime.dwLowDateTime = 0;
        }

        InfoKey(DWORD numberOfSubKeys, DWORD numberOfValues, FILETIME lastWriteTime) noexcept
            : NumberOfSubKeys{ numberOfSubKeys }
            , NumberOfValues{ numberOfValues }
            , LastWriteTime{ lastWriteTime }
        {
        }
    };

    // Retrieve information about the registry key
    [[nodiscard]] InfoKey QueryInfoKey() const;

    // Return the DWORD type ID for the input registry value
    [[nodiscard]] DWORD QueryValueType(const std::wstring& valueName) const;


    enum class KeyReflection
    {
        ReflectionEnabled,
        ReflectionDisabled
    };

    // Determines whether reflection has been disabled or enabled for the specified key
    [[nodiscard]] KeyReflection QueryReflectionKey() const;

    // Enumerate the subkeys of the registry key, using RegEnumKeyEx
    [[nodiscard]] std::vector<std::wstring> EnumSubKeys() const;

    // Enumerate the values under the registry key, using RegEnumValue.
    // Returns a vector of pairs: In each pair, the wstring is the value name,
    // the DWORD is the value type.
    [[nodiscard]] std::vector<std::pair<std::wstring, DWORD>> EnumValues() const;


    //
    // Query Operations Returning RegExpected
    // (instead of throwing RegException on error)
    //

    // Retrieve information about the registry key
    [[nodiscard]] RegExpected<InfoKey> TryQueryInfoKey() const;

    // Return the DWORD type ID for the input registry value
    [[nodiscard]] RegExpected<DWORD> TryQueryValueType(const std::wstring& valueName) const;


    // Determines whether reflection has been disabled or enabled for the specified key
    [[nodiscard]] RegExpected<KeyReflection> TryQueryReflectionKey() const;

    // Enumerate the subkeys of the registry key, using RegEnumKeyEx
    [[nodiscard]] RegExpected<std::vector<std::wstring>> TryEnumSubKeys() const;

    // Enumerate the values under the registry key, using RegEnumValue.
    // Returns a vector of pairs: In each pair, the wstring is the value name,
    // the DWORD is the value type.
    [[nodiscard]] RegExpected<std::vector<std::pair<std::wstring, DWORD>>> TryEnumValues() const;


    //
    // Misc Registry API Wrappers
    //

    void DeleteValue(const std::wstring& valueName);
    void DeleteKey(const std::wstring& subKey, REGSAM desiredAccess);
    void DeleteTree(const std::wstring& subKey);
    void CopyTree(const std::wstring& sourceSubKey, const RegKey& destKey);
    void FlushKey();
    void LoadKey(const std::wstring& subKey, const std::wstring& filename);
    void SaveKey(const std::wstring& filename, SECURITY_ATTRIBUTES* securityAttributes) const;
    void EnableReflectionKey();
    void DisableReflectionKey();
    void ConnectRegistry(const std::wstring& machineName, HKEY hKeyPredefined);


    //
    // Misc Registry API Wrappers Returning RegResult Status
    // (instead of throwing RegException on error)
    //

    [[nodiscard]] RegResult TryDeleteValue(const std::wstring& valueName) noexcept;
    [[nodiscard]] RegResult TryDeleteKey(const std::wstring& subKey, REGSAM desiredAccess) noexcept;
    [[nodiscard]] RegResult TryDeleteTree(const std::wstring& subKey) noexcept;

    [[nodiscard]] RegResult TryCopyTree(const std::wstring& sourceSubKey,
                                        const RegKey& destKey) noexcept;

    [[nodiscard]] RegResult TryFlushKey() noexcept;

    [[nodiscard]] RegResult TryLoadKey(const std::wstring& subKey,
                                       const std::wstring& filename) noexcept;

    [[nodiscard]] RegResult TrySaveKey(const std::wstring& filename,
                                       SECURITY_ATTRIBUTES* securityAttributes) const noexcept;

    [[nodiscard]] RegResult TryEnableReflectionKey() noexcept;
    [[nodiscard]] RegResult TryDisableReflectionKey() noexcept;

    [[nodiscard]] RegResult TryConnectRegistry(const std::wstring& machineName,
                                               HKEY hKeyPredefined) noexcept;


    // Return a string representation of Windows registry types
    [[nodiscard]] static std::wstring RegTypeToString(DWORD regType);


    //
    // Relational comparison operators are overloaded as non-members
    // ==, !=, <, <=, >, >=
    //


    //
    // Private Implementation
    //

private:
    // The wrapped registry key handle
    HKEY m_hKey{ nullptr };


    //
    // Private Helpers
    //

    //
    // These GetDword/QWord/String/etc./ValueInternal methods implement the code
    // that interacts directly with Windows Registry C-interface APIs
    // to get the desired values from the registry.
    //
    // Public methods like GetDword/Qword/String/etc./Value and the corresponding
    // TryGetXxxxValue methods simply call these "GetXxxxValueInternal" methods.
    //
    // Note that in all these GetXxxxValueInternal methods the errorMessage parameter
    // is an *optional* string. This error message string is used by the public GetXxxxValue
    // methods that return an exception on error; however, the corresponding TryGetXxxxValue
    // methods do not use this error message string as they don't throw a RegException on error.
    // (They could potentially use the error message string to log the error,
    // but they do not in this current implementation.)
    //

    [[nodiscard]] LSTATUS GetDwordValueInternal(
        DWORD& result,                           // [out] to be read from the registry
        const std::wstring& valueName,           // [in]  name of the DWORD value
        std::string* errorMessage = nullptr      // [out, opt] a descriptive error message
    ) const;

    [[nodiscard]] LSTATUS GetQwordValueInternal(
        ULONGLONG& result,                       // [out] to be read from the registry
        const std::wstring& valueName,           // [in]  name of the QWORD value
        std::string* errorMessage = nullptr      // [out, opt] a descriptive error message
    ) const;

    [[nodiscard]] LSTATUS GetStringValueInternal(
        std::wstring&       result,                     // [out] to be read from the registry
        const std::wstring& valueName,                  // [in]  name of the string value
        std::string*        errorMessage = nullptr      // [out, opt] a descriptive error message
    ) const;

    [[nodiscard]] LSTATUS GetExpandStringValueInternal(
        std::wstring& result,                     // [out] to be read from the registry
        const std::wstring& valueName,            // [in]  name of the expand string value
        ExpandStringOption  expandOption = ExpandStringOption::Expand, // [in] Expand option
        std::string* errorMessage = nullptr      // [out, opt] a descriptive error message
    ) const;

    [[nodiscard]] LSTATUS GetMultiStringValueInternal(
        std::vector<std::wstring>& result,              // [out] to be read from the registry
        const std::wstring& valueName,                  // [in]  name of the multi-string value
        std::string* errorMessage = nullptr             // [out, opt] a descriptive error message
    ) const;

    [[nodiscard]] LSTATUS RegKey::GetBinaryValueInternal(
        std::vector<BYTE>& result,                      // [out] to be read from the registry
        const std::wstring& valueName,                  // [in]  name of the binary value
        std::string* const errorMessage = nullptr       // [out, opt] a descriptive error message
    ) const;
};


//------------------------------------------------------------------------------
// An exception representing an error with the registry operations
//------------------------------------------------------------------------------
class RegException
    : public std::system_error
{
public:
    RegException(LSTATUS errorCode, const char* message);
    RegException(LSTATUS errorCode, const std::string& message);
};


//------------------------------------------------------------------------------
// A tiny wrapper around LSTATUS return codes used by the Windows Registry API.
//------------------------------------------------------------------------------
class RegResult
{
public:

    // Initialize to success code (ERROR_SUCCESS)
    RegResult() noexcept = default;

    // Initialize with specific Windows Registry API LSTATUS return code
    explicit RegResult(LSTATUS result) noexcept;

    // Is the wrapped code a success code?
    [[nodiscard]] bool IsOk() const noexcept;

    // Is the wrapped error code a failure code?
    [[nodiscard]] bool Failed() const noexcept;

    // Is the wrapped code a success code?
    [[nodiscard]] explicit operator bool() const noexcept;

    // Get the wrapped Win32 code
    [[nodiscard]] LSTATUS Code() const noexcept;

    // Return the system error message associated to the current error code
    [[nodiscard]] std::wstring ErrorMessage() const;

    // Return the system error message associated to the current error code,
    // using the given input language identifier
    [[nodiscard]] std::wstring ErrorMessage(DWORD languageId) const;

private:
    // Error code returned by Windows Registry C API;
    // default initialized to success code.
    LSTATUS m_result{ ERROR_SUCCESS };
};


//------------------------------------------------------------------------------
// A class template that stores a value of type T (e.g. DWORD, std::wstring)
// on success, or a RegResult on error.
//
// Used as the return value of some Registry RegKey::TryGetXxxValue() methods
// as an alternative to exception-throwing methods.
//------------------------------------------------------------------------------
template <typename T>
class RegExpected
{
public:
    // Initialize the object with an error code
    explicit RegExpected(const RegResult& errorCode) noexcept;

    // Initialize the object with a value (the success case)
    explicit RegExpected(const T& value);

    // Initialize the object with a value (the success case),
    // optimized for move semantics
    explicit RegExpected(T&& value);

    // Does this object contain a valid value?
    [[nodiscard]] explicit operator bool() const noexcept;

    // Does this object contain a valid value?
    [[nodiscard]] bool IsValid() const noexcept;

    // Access the value (if the object contains a valid value).
    // Throws an exception if the object is in invalid state.
    [[nodiscard]] const T& GetValue() const;

    // Access the error code (if the object contains an error status)
    // Can also be called if the object contains a valid value
    // (in this case, the returned error code is ERROR_SUCCESS).
    [[nodiscard]] RegResult GetError() const;


private:
    RegResult m_result{ ERROR_SUCCESS };

    // Value is valid only if m_result is ERROR_SUCCESS
    T m_value{};
};


//------------------------------------------------------------------------------
//          Overloads of relational comparison operators for RegKey
//------------------------------------------------------------------------------

inline bool operator==(const RegKey& a, const RegKey& b) noexcept
{
    return a.Get() == b.Get();
}

inline bool operator!=(const RegKey& a, const RegKey& b) noexcept
{
    return a.Get() != b.Get();
}

inline bool operator<(const RegKey& a, const RegKey& b) noexcept
{
    return a.Get() < b.Get();
}

inline bool operator<=(const RegKey& a, const RegKey& b) noexcept
{
    return a.Get() <= b.Get();
}

inline bool operator>(const RegKey& a, const RegKey& b) noexcept
{
    return a.Get() > b.Get();
}

inline bool operator>=(const RegKey& a, const RegKey& b) noexcept
{
    return a.Get() >= b.Get();
}


//------------------------------------------------------------------------------
//                  Private Helper Classes and Functions
//------------------------------------------------------------------------------

//
// Note: Naming this private namespace `winreg_internal` instead of just
// `internal` (or `detail`) helps protecting against client code
// that does something like:
//
//      using namespace winreg;
//
// In such cases, WinReg's internal private helper code is still *protected*
// under the `winreg_internal` namespace, and will not collide against
// other libraries' internal/detail namespaces.
//

namespace winreg_internal
{

//------------------------------------------------------------------------------
// Simple scoped-based RAII wrapper that *automatically* invokes LocalFree()
// in its destructor.
//------------------------------------------------------------------------------
template <typename T>
class ScopedLocalFree
{
public:

    typedef T  Type;
    typedef T* TypePtr;


    // Init wrapped pointer to nullptr
    ScopedLocalFree() noexcept = default;

    // Automatically and safely invoke ::LocalFree()
    ~ScopedLocalFree() noexcept
    {
        Free();
    }

    //
    // Ban copy and move operations
    //
    ScopedLocalFree(const ScopedLocalFree&) = delete;
    ScopedLocalFree(ScopedLocalFree&&) = delete;
    ScopedLocalFree& operator=(const ScopedLocalFree&) = delete;
    ScopedLocalFree& operator=(ScopedLocalFree&&) = delete;


    // Read-only access to the wrapped pointer
    [[nodiscard]] T* Get() const noexcept
    {
        return m_ptr;
    }

    // Writable access to the wrapped pointer
    [[nodiscard]] T** AddressOf() noexcept
    {
        return &m_ptr;
    }

    // Explicit pointer conversion to bool
    explicit operator bool() const noexcept
    {
        return (m_ptr != nullptr);
    }

    // Safely invoke ::LocalFree() on the wrapped pointer
    void Free() noexcept
    {
        if (m_ptr != nullptr)
        {
            ::LocalFree(m_ptr);
            m_ptr = nullptr;
        }
    }


    //
    // IMPLEMENTATION
    //
private:
    T* m_ptr{ nullptr };
};


//------------------------------------------------------------------------------
// Helper function to build a multi-string from a vector<wstring>.
//
// A multi-string is a sequence of contiguous NUL-terminated strings,
// that terminates with an additional NUL.
// Basically, considered as a whole, the sequence is terminated by two NULs.
// E.g.:
//          Hello\0World\0\0
//------------------------------------------------------------------------------
[[nodiscard]] inline std::vector<wchar_t> BuildMultiString(const std::vector<std::wstring>& data)
{
    // Special case of the empty multi-string
    if (data.empty())
    {
        // Build a vector containing just two NULs
        return std::vector<wchar_t>(2, L'\0');
    }

    // Get the total length in wchar_ts of the multi-string
    size_t totalLen = 0;
    for (const auto& s : data)
    {
        // Add one to current string's length for the terminating NUL
        totalLen += (s.length() + 1);
    }

    // Add one for the last NUL terminator (making the whole structure double-NUL terminated)
    totalLen++;

    // Allocate a buffer to store the multi-string
    std::vector<wchar_t> multiString;

    // Reserve room in the vector to speed up the following insertion loop
    multiString.reserve(totalLen);

    // Copy the single strings into the multi-string
    for (const auto& s : data)
    {
        if (!s.empty())
        {
            // Copy current string's content
            multiString.insert(multiString.end(), s.begin(), s.end());
        }

        // Don't forget to NUL-terminate the current string
        // (or just insert L'\0' for empty strings)
        multiString.emplace_back(L'\0');
    }

    // Add the last NUL-terminator
    multiString.emplace_back(L'\0');

    return multiString;
}


//------------------------------------------------------------------------------
// Return true if the wchar_t sequence stored in 'data' terminates
// with two null (L'\0') wchar_t's
//------------------------------------------------------------------------------
[[nodiscard]] inline bool IsDoubleNullTerminated(const std::vector<wchar_t>& data)
{
    // First check that there's enough room for at least two nulls
    if (data.size() < 2)
    {
        return false;
    }

    // Check that the sequence terminates with two nulls (L'\0', L'\0')
    const size_t lastPosition = data.size() - 1;
    return ((data[lastPosition]     == L'\0')  &&
            (data[lastPosition - 1] == L'\0')) ? true : false;
}


//------------------------------------------------------------------------------
// Make sure to double-null-terminate the vector.
// If the vector is *already* double-null-terminated, leave it as is.
//------------------------------------------------------------------------------
inline void DoubleNullTerminate(std::vector<wchar_t>& data)
{
    // If the vector is *already* double-null-terminated, leave it as is.
    if (IsDoubleNullTerminated(data))
    {
        return;
    }


    // The vector is not double-null-terminated.
    // We have to add one or two null-terminators.

    // Is the vector empty?
    if (data.empty())
    {
        // Add the double-null terminator
        data.push_back(L'\0');
        data.push_back(L'\0');

        _ASSERTE(IsDoubleNullTerminated(data));
        return;
    }

    // The vector is not empty.
    // Start checking the element in the last position.
    const size_t lastPosition = data.size() - 1;
    if (data[lastPosition] != L'\0')
    {
        // The last element is not a null (L'\0').
        // Se we need to add *two* null terminators to make the vector double-null-terminated.
        data.push_back(L'\0');
        data.push_back(L'\0');

        _ASSERTE(IsDoubleNullTerminated(data));
        return;
    }

    // The last element is null.
    // If that is the *only* element, add another null and we are done.
    if (data.size() == 1)
    {
        data.push_back(L'\0');

        _ASSERTE(IsDoubleNullTerminated(data));
        return;
    }

    // Now check the element *before* the last one.
    // (If we are here, there are at least two elements.)
    _ASSERTE(data.size() >= 2);
    if (data[lastPosition - 1] != L'\0')
    {
        // The last element is null, but the previous element is *not* null.
        //
        // C o n n i e 0
        //           X *<--- null
        //           |
        //           not-null
        //
        // So, just add another null at the end of the vector, and we'll make it
        // double-null terminated.
        data.push_back(L'\0');

        _ASSERTE(IsDoubleNullTerminated(data));
        return;
    }

    _ASSERTE(IsDoubleNullTerminated(data));
}


//------------------------------------------------------------------------------
// Given a sequence of wchar_ts representing a double-null-terminated string,
// returns a vector of wstrings that represent the single strings.
//
// Also supports embedded empty strings in the sequence.
//------------------------------------------------------------------------------
[[nodiscard]] inline std::vector<std::wstring> ParseMultiString(const std::vector<wchar_t>& data)
{
    // Make sure that there are two terminating L'\0's at the end of the sequence
    if (!IsDoubleNullTerminated(data))
    {
        throw RegException{ ERROR_INVALID_DATA, "Not a double-null terminated string." };
    }

    // Parse the double-NUL-terminated string into a vector<wstring>,
    // which will be returned to the caller
    std::vector<std::wstring> result;

    //
    // Note on Embedded Empty Strings
    // ==============================
    //
    // Below commented-out there is the previous parsing code,
    // that assumes that an empty string *terminates* the sequence.
    //
    // In fact, according to the official Microsoft MSDN documentation,
    // an empty string is treated as a sequence terminator,
    // so you can't have empty strings inside the sequence.
    //
    // Source: https://docs.microsoft.com/en-us/windows/win32/sysinfo/registry-value-types
    //      "A REG_MULTI_SZ string ends with a string of length 0.
    //       Therefore, it is not possible to include a zero-length string
    //       in the sequence. An empty sequence would be defined as follows: \0."
    //
    // Unfortunately, it seems that Microsoft violates its own rule, for example
    // in the PendingFileRenameOperations value under the
    // "SYSTEM\CurrentControlSet\Control\Session Manager" key.
    // This is a REG_MULTI_SZ value that does contain embedded empty strings.
    //
    // So, I changed the previous parsing code to support also embedded empty strings.
    //
    // -------------------------------------------------------------------------
    //// *** Previous parsing code - Assumes an empty string terminates the sequence ***
    //
    //const wchar_t* currStringPtr = data.data();
    //while (*currStringPtr != L'\0')
    //{
    //    // Current string is NUL-terminated, so get its length calling wcslen
    //    const size_t currStringLength = wcslen(currStringPtr);
    //
    //    // Add current string to the result vector
    //    result.emplace_back(currStringPtr, currStringLength);
    //
    //    // Move to the next string
    //    currStringPtr += currStringLength + 1;
    //}
    // -------------------------------------------------------------------------
    //

    const wchar_t* currStringPtr = data.data();
    const wchar_t* const endPtr  = data.data() + data.size() - 1;

    while (currStringPtr < endPtr)
    {
        // Current string is NUL-terminated, so get its length calling wcslen
        const size_t currStringLength = wcslen(currStringPtr);

        // Add current string to the result vector
        if (currStringLength > 0)
        {
            result.emplace_back(currStringPtr, currStringLength);
        }
        else
        {
            // Insert empty strings, as well
            result.emplace_back(std::wstring{});
        }

        // Move to the next string, skipping the terminating NUL
        currStringPtr += currStringLength + 1;
    }

    return result;
}


//------------------------------------------------------------------------------
// Builds a RegExpected object that stores an error code
//------------------------------------------------------------------------------
template <typename T>
[[nodiscard]] inline RegExpected<T> MakeRegExpectedWithError(const LSTATUS retCode)
{
    return RegExpected<T>{ RegResult{ retCode } };
}


} // namespace winreg_internal


//------------------------------------------------------------------------------
//                          RegKey Inline Methods
//------------------------------------------------------------------------------

inline RegKey::RegKey(const HKEY hKey) noexcept
    : m_hKey{ hKey }
{
}


inline RegKey::RegKey(const HKEY hKeyParent, const std::wstring& subKey)
{
    Create(hKeyParent, subKey);
}


inline RegKey::RegKey(const HKEY hKeyParent, const std::wstring& subKey, const REGSAM desiredAccess)
{
    Create(hKeyParent, subKey, desiredAccess);
}


inline RegKey::RegKey(RegKey&& other) noexcept
    : m_hKey{ other.m_hKey }
{
    // Other doesn't own the handle anymore
    other.m_hKey = nullptr;
}


inline RegKey& RegKey::operator=(RegKey&& other) noexcept
{
    // Prevent self-move-assign
    if ((this != &other) && (m_hKey != other.m_hKey))
    {
        // Close current
        Close();

        // Move from other (i.e. take ownership of other's raw handle)
        m_hKey = other.m_hKey;
        other.m_hKey = nullptr;
    }
    return *this;
}


inline RegKey::~RegKey() noexcept
{
    // Release the owned handle (if any)
    Close();
}


inline HKEY RegKey::Get() const noexcept
{
    return m_hKey;
}


inline void RegKey::Close() noexcept
{
    if (IsValid())
    {
        // Do not call RegCloseKey on predefined keys
        if (! IsPredefined())
        {
            ::RegCloseKey(m_hKey);
        }

        // Avoid dangling references
        m_hKey = nullptr;
    }
}


inline bool RegKey::IsValid() const noexcept
{
    return m_hKey != nullptr;
}


inline RegKey::operator bool() const noexcept
{
    return IsValid();
}


inline bool RegKey::IsPredefined() const noexcept
{
    // Predefined keys
    // https://msdn.microsoft.com/en-us/library/windows/desktop/ms724836(v=vs.85).aspx

    if (   (m_hKey == HKEY_CURRENT_USER)
        || (m_hKey == HKEY_LOCAL_MACHINE)
        || (m_hKey == HKEY_CLASSES_ROOT)
        || (m_hKey == HKEY_CURRENT_CONFIG)
        || (m_hKey == HKEY_CURRENT_USER_LOCAL_SETTINGS)
        || (m_hKey == HKEY_PERFORMANCE_DATA)
        || (m_hKey == HKEY_PERFORMANCE_NLSTEXT)
        || (m_hKey == HKEY_PERFORMANCE_TEXT)
        || (m_hKey == HKEY_USERS))
    {
        return true;
    }

    return false;
}


inline HKEY RegKey::Detach() noexcept
{
    HKEY hKey = m_hKey;

    // We don't own the HKEY handle anymore
    m_hKey = nullptr;

    // Transfer ownership to the caller
    return hKey;
}


inline void RegKey::Attach(const HKEY hKey) noexcept
{
    // Prevent self-attach
    if (m_hKey != hKey)
    {
        // Close any open registry handle
        Close();

        // Take ownership of the input hKey
        m_hKey = hKey;
    }
}


inline void RegKey::SwapWith(RegKey& other) noexcept
{
    // Enable ADL (not necessary in this case, but good practice)
    using std::swap;

    // Swap the raw handle members
    swap(m_hKey, other.m_hKey);
}


inline void swap(RegKey& a, RegKey& b) noexcept
{
    a.SwapWith(b);
}


inline void RegKey::Create(
    const HKEY                  hKeyParent,
    const std::wstring&         subKey,
    const REGSAM                desiredAccess
)
{
    constexpr DWORD kDefaultOptions = REG_OPTION_NON_VOLATILE;

    Create(hKeyParent, subKey, desiredAccess, kDefaultOptions,
        nullptr, // no security attributes,
        nullptr  // no disposition
    );
}


inline void RegKey::Create(
    const HKEY                  hKeyParent,
    const std::wstring&         subKey,
    const REGSAM                desiredAccess,
    const DWORD                 options,
    SECURITY_ATTRIBUTES* const  securityAttributes,
    DWORD* const                disposition
)
{
    HKEY hKey = nullptr;
    LSTATUS retCode = ::RegCreateKeyExW(
        hKeyParent,
        subKey.c_str(),
        0,          // reserved
        REG_NONE,   // user-defined class type parameter not supported
        options,
        desiredAccess,
        securityAttributes,
        &hKey,
        disposition
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegCreateKeyExW failed." };
    }

    // Safely close any previously opened key
    Close();

    // Take ownership of the newly created key
    m_hKey = hKey;
}


inline void RegKey::Open(
    const HKEY              hKeyParent,
    const std::wstring&     subKey,
    const REGSAM            desiredAccess
)
{
    HKEY hKey = nullptr;
    LSTATUS retCode = ::RegOpenKeyExW(
        hKeyParent,
        subKey.c_str(),
        REG_NONE,           // default options
        desiredAccess,
        &hKey
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegOpenKeyExW failed." };
    }

    // Safely close any previously opened key
    Close();

    // Take ownership of the newly created key
    m_hKey = hKey;
}


inline RegResult RegKey::TryCreate(
    const HKEY          hKeyParent,
    const std::wstring& subKey,
    const REGSAM        desiredAccess
) noexcept
{
    constexpr DWORD kDefaultOptions = REG_OPTION_NON_VOLATILE;

    return TryCreate(hKeyParent, subKey, desiredAccess, kDefaultOptions,
        nullptr, // no security attributes,
        nullptr  // no disposition
    );
}


inline RegResult RegKey::TryCreate(
    const HKEY                  hKeyParent,
    const std::wstring&         subKey,
    const REGSAM                desiredAccess,
    const DWORD                 options,
    SECURITY_ATTRIBUTES* const  securityAttributes,
    DWORD* const                disposition
) noexcept
{
    HKEY hKey = nullptr;
    RegResult retCode{ ::RegCreateKeyExW(
        hKeyParent,
        subKey.c_str(),
        0,          // reserved
        REG_NONE,   // user-defined class type parameter not supported
        options,
        desiredAccess,
        securityAttributes,
        &hKey,
        disposition
    ) };
    if (retCode.Failed())
    {
        return retCode;
    }

    // Safely close any previously opened key
    Close();

    // Take ownership of the newly created key
    m_hKey = hKey;

    _ASSERTE(retCode.IsOk());
    return retCode;
}


inline RegResult RegKey::TryOpen(
    const HKEY          hKeyParent,
    const std::wstring& subKey,
    const REGSAM        desiredAccess
) noexcept
{
    HKEY hKey = nullptr;
    RegResult retCode{ ::RegOpenKeyExW(
        hKeyParent,
        subKey.c_str(),
        REG_NONE,           // default options
        desiredAccess,
        &hKey
    ) };
    if (retCode.Failed())
    {
        return retCode;
    }

    // Safely close any previously opened key
    Close();

    // Take ownership of the newly created key
    m_hKey = hKey;

    _ASSERTE(retCode.IsOk());
    return retCode;
}


inline void RegKey::SetDwordValue(const std::wstring& valueName, const DWORD data)
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_DWORD,
        reinterpret_cast<const BYTE*>(&data),
        sizeof(data)
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot write DWORD value: RegSetValueExW failed." };
    }
}


inline void RegKey::SetQwordValue(const std::wstring& valueName, const ULONGLONG& data)
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_QWORD,
        reinterpret_cast<const BYTE*>(&data),
        sizeof(data)
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot write QWORD value: RegSetValueExW failed." };
    }
}


inline void RegKey::SetStringValue(const std::wstring& valueName, const std::wstring& data)
{
    _ASSERTE(IsValid());

    // String size including the terminating NUL, in bytes
    const DWORD dataSize = static_cast<DWORD>((data.length() + 1) * sizeof(wchar_t));

    LSTATUS retCode = ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_SZ,
        reinterpret_cast<const BYTE*>(data.c_str()),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot write string value: RegSetValueExW failed." };
    }
}


inline void RegKey::SetExpandStringValue(const std::wstring& valueName, const std::wstring& data)
{
    _ASSERTE(IsValid());

    // String size including the terminating NUL, in bytes
    const DWORD dataSize = static_cast<DWORD>((data.length() + 1) * sizeof(wchar_t));

    LSTATUS retCode = ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_EXPAND_SZ,
        reinterpret_cast<const BYTE*>(data.c_str()),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot write expand string value: RegSetValueExW failed." };
    }
}


inline void RegKey::SetMultiStringValue(
    const std::wstring& valueName,
    const std::vector<std::wstring>& data
)
{
    _ASSERTE(IsValid());

    // First, we have to build a double-NUL-terminated multi-string from the input data
    const std::vector<wchar_t> multiString = winreg_internal::BuildMultiString(data);

    // Total size, in bytes, of the whole multi-string structure
    const DWORD dataSize = static_cast<DWORD>(multiString.size() * sizeof(wchar_t));

    LSTATUS retCode = ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_MULTI_SZ,
        reinterpret_cast<const BYTE*>(multiString.data()),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot write multi-string value: RegSetValueExW failed." };
    }
}


inline void RegKey::SetBinaryValue(const std::wstring& valueName, const std::vector<BYTE>& data)
{
    _ASSERTE(IsValid());

    // Total data size, in bytes
    const DWORD dataSize = static_cast<DWORD>(data.size());

    LSTATUS retCode = ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_BINARY,
        data.data(),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot write binary data value: RegSetValueExW failed." };
    }
}


inline void RegKey::SetBinaryValue(
    const std::wstring& valueName,
    const void* const data,
    const DWORD dataSize
)
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_BINARY,
        static_cast<const BYTE*>(data),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot write binary data value: RegSetValueExW failed." };
    }
}


inline RegResult RegKey::TrySetDwordValue(const std::wstring& valueName, const DWORD data) noexcept
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_DWORD,
        reinterpret_cast<const BYTE*>(&data),
        sizeof(data)
    ) };
}


inline RegResult RegKey::TrySetQwordValue(const std::wstring& valueName,
                                          const ULONGLONG& data) noexcept
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_QWORD,
        reinterpret_cast<const BYTE*>(&data),
        sizeof(data)
    ) };
}


inline RegResult RegKey::TrySetStringValue(const std::wstring& valueName,
                                           const std::wstring& data) noexcept
{
    _ASSERTE(IsValid());

    // String size including the terminating NUL, in bytes
    const DWORD dataSize = static_cast<DWORD>((data.length() + 1) * sizeof(wchar_t));

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_SZ,
        reinterpret_cast<const BYTE*>(data.c_str()),
        dataSize
    ) };
}


inline RegResult RegKey::TrySetExpandStringValue(const std::wstring& valueName,
                                                 const std::wstring& data) noexcept
{
    _ASSERTE(IsValid());

    // String size including the terminating NUL, in bytes
    const DWORD dataSize = static_cast<DWORD>((data.length() + 1) * sizeof(wchar_t));

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_EXPAND_SZ,
        reinterpret_cast<const BYTE*>(data.c_str()),
        dataSize
    ) };
}


inline RegResult RegKey::TrySetMultiStringValue(const std::wstring& valueName,
                                                const std::vector<std::wstring>& data)
{
    _ASSERTE(IsValid());

    // First, we have to build a double-NUL-terminated multi-string from the input data.
    //
    // NOTE: This is the reason why I *cannot* mark this method noexcept,
    // since a *dynamic allocation* happens for creating the std::vector in BuildMultiString.
    // And, if dynamic memory allocations fail, an exception is thrown.
    //
    const std::vector<wchar_t> multiString = winreg_internal::BuildMultiString(data);

    // Total size, in bytes, of the whole multi-string structure
    const DWORD dataSize = static_cast<DWORD>(multiString.size() * sizeof(wchar_t));

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_MULTI_SZ,
        reinterpret_cast<const BYTE*>(multiString.data()),
        dataSize
    ) };
}


inline RegResult RegKey::TrySetBinaryValue(const std::wstring& valueName,
                                           const std::vector<BYTE>& data) noexcept
{
    _ASSERTE(IsValid());

    // Total data size, in bytes
    const DWORD dataSize = static_cast<DWORD>(data.size());

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_BINARY,
        data.data(),
        dataSize
    ) };
}


inline RegResult RegKey::TrySetBinaryValue(const std::wstring& valueName,
                                           const void* const data,
                                           const DWORD dataSize) noexcept
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_BINARY,
        static_cast<const BYTE*>(data),
        dataSize
    ) };
}


inline DWORD RegKey::GetDwordValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    DWORD result{};    // to be read from the registry
    std::string errorMessage;
    LSTATUS retCode = GetDwordValueInternal(result, valueName, &errorMessage);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, errorMessage };
    }

    return result;
}


inline ULONGLONG RegKey::GetQwordValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    ULONGLONG result{};    // to be read from the registry
    std::string errorMessage;
    LSTATUS retCode = GetQwordValueInternal(result, valueName, &errorMessage);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, errorMessage };
    }

    return result;
}


inline std::wstring RegKey::GetStringValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    std::wstring result;    // to be read from the registry
    std::string errorMessage;
    LSTATUS retCode = GetStringValueInternal(result, valueName, &errorMessage);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, errorMessage };
    }

    return result;
}


inline std::wstring RegKey::GetExpandStringValue(
    const std::wstring& valueName,
    const ExpandStringOption expandOption
) const
{
    _ASSERTE(IsValid());

    std::wstring result;    // to be read from the registry
    std::string errorMessage;
    LSTATUS retCode = GetExpandStringValueInternal(result, valueName, expandOption, &errorMessage);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, errorMessage };
    }

    return result;
}


inline std::vector<std::wstring> RegKey::GetMultiStringValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    std::vector<std::wstring> result;    // to be read from the registry
    std::string errorMessage;
    LSTATUS retCode = GetMultiStringValueInternal(result, valueName, &errorMessage);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, errorMessage };
    }

    return result;
}


inline std::vector<BYTE> RegKey::GetBinaryValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    std::vector<BYTE> result;    // to be read from the registry
    std::string errorMessage;
    LSTATUS retCode = GetBinaryValueInternal(result, valueName, &errorMessage);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, errorMessage };
    }

    return result;
}


inline RegExpected<DWORD> RegKey::TryGetDwordValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    using RegValueType = DWORD;

    DWORD result = 0;    // to be read from the registry
    LSTATUS retCode = GetDwordValueInternal(result, valueName);
    if (retCode != ERROR_SUCCESS)
    {
        return winreg_internal::MakeRegExpectedWithError<RegValueType>(retCode);
    }

    return RegExpected<RegValueType>{ result };
}


inline RegExpected<ULONGLONG> RegKey::TryGetQwordValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    using RegValueType = ULONGLONG;

    ULONGLONG result = 0;    // to be read from the registry
    LSTATUS retCode = GetQwordValueInternal(result, valueName);
    if (retCode != ERROR_SUCCESS)
    {
        return winreg_internal::MakeRegExpectedWithError<RegValueType>(retCode);
    }

    return RegExpected<RegValueType>{ result };
}


inline RegExpected<std::wstring> RegKey::TryGetStringValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    using RegValueType = std::wstring;

    std::wstring result;    // to be read from the registry
    LSTATUS retCode = GetStringValueInternal(result, valueName);
    if (retCode != ERROR_SUCCESS)
    {
        return winreg_internal::MakeRegExpectedWithError<RegValueType>(retCode);
    }

    return RegExpected<RegValueType>{ result };
}


inline RegExpected<std::wstring> RegKey::TryGetExpandStringValue(
    const std::wstring& valueName,
    const ExpandStringOption expandOption
) const
{
    _ASSERTE(IsValid());

    using RegValueType = std::wstring;

    std::wstring result;    // to be read from the registry
    LSTATUS retCode = GetExpandStringValueInternal(result, valueName, expandOption);
    if (retCode != ERROR_SUCCESS)
    {
        return winreg_internal::MakeRegExpectedWithError<RegValueType>(retCode);
    }

    return RegExpected<RegValueType>{ result };
}


inline RegExpected<std::vector<std::wstring>>
    RegKey::TryGetMultiStringValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    using RegValueType = std::vector<std::wstring>;

    std::vector<std::wstring> result;    // to be read from the registry
    LSTATUS retCode = GetMultiStringValueInternal(result, valueName);
    if (retCode != ERROR_SUCCESS)
    {
        return winreg_internal::MakeRegExpectedWithError<RegValueType>(retCode);
    }

    return RegExpected<RegValueType>{ result };
}


inline RegExpected<std::vector<BYTE>>
RegKey::TryGetBinaryValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    using RegValueType = std::vector<BYTE>;

    std::vector<BYTE> result;    // to be read from the registry
    LSTATUS retCode = GetBinaryValueInternal(result, valueName);
    if (retCode != ERROR_SUCCESS)
    {
        throw winreg_internal::MakeRegExpectedWithError<RegValueType>(retCode);
    }

    return RegExpected<RegValueType>{ result };
}


inline std::vector<std::wstring> RegKey::EnumSubKeys() const
{
    _ASSERTE(IsValid());

    // Get some useful enumeration info, like the total number of subkeys
    // and the maximum length of the subkey names
    DWORD subKeyCount = 0;
    DWORD maxSubKeyNameLen = 0;
    LSTATUS retCode = ::RegQueryInfoKeyW(
        m_hKey,
        nullptr,    // no user-defined class
        nullptr,    // no user-defined class size
        nullptr,    // reserved
        &subKeyCount,
        &maxSubKeyNameLen,
        nullptr,    // no subkey class length
        nullptr,    // no value count
        nullptr,    // no value name max length
        nullptr,    // no max value length
        nullptr,    // no security descriptor
        nullptr     // no last write time
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{
            retCode,
            "RegQueryInfoKeyW failed while preparing for subkey enumeration."
        };
    }

    // NOTE: According to the MSDN documentation, the size returned for subkey name max length
    // does *not* include the terminating NUL, so let's add +1 to take it into account
    // when I allocate the buffer for reading subkey names.
    maxSubKeyNameLen++;

    // Preallocate a buffer for the subkey names
    auto nameBuffer = std::make_unique<wchar_t[]>(maxSubKeyNameLen);

    // The result subkey names will be stored here
    std::vector<std::wstring> subkeyNames;

    // Reserve room in the vector to speed up the following insertion loop
    subkeyNames.reserve(subKeyCount);

    // Enumerate all the subkeys
    for (DWORD index = 0; index < subKeyCount; index++)
    {
        // Get the name of the current subkey
        DWORD subKeyNameLen = maxSubKeyNameLen;
        retCode = ::RegEnumKeyExW(
            m_hKey,
            index,
            nameBuffer.get(),
            &subKeyNameLen,
            nullptr, // reserved
            nullptr, // no class
            nullptr, // no class
            nullptr  // no last write time
        );
        if (retCode != ERROR_SUCCESS)
        {
            throw RegException{ retCode, "Cannot enumerate subkeys: RegEnumKeyExW failed." };
        }

        // On success, the ::RegEnumKeyEx API writes the length of the
        // subkey name in the subKeyNameLen output parameter
        // (not including the terminating NUL).
        // So I can build a wstring based on that length.
        subkeyNames.emplace_back(nameBuffer.get(), subKeyNameLen);
    }

    return subkeyNames;
}


inline std::vector<std::pair<std::wstring, DWORD>> RegKey::EnumValues() const
{
    _ASSERTE(IsValid());

    // Get useful enumeration info, like the total number of values
    // and the maximum length of the value names
    DWORD valueCount = 0;
    DWORD maxValueNameLen = 0;
    LSTATUS retCode = ::RegQueryInfoKeyW(
        m_hKey,
        nullptr,    // no user-defined class
        nullptr,    // no user-defined class size
        nullptr,    // reserved
        nullptr,    // no subkey count
        nullptr,    // no subkey max length
        nullptr,    // no subkey class length
        &valueCount,
        &maxValueNameLen,
        nullptr,    // no max value length
        nullptr,    // no security descriptor
        nullptr     // no last write time
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{
            retCode,
            "RegQueryInfoKeyW failed while preparing for value enumeration."
        };
    }

    // NOTE: According to the MSDN documentation, the size returned for value name max length
    // does *not* include the terminating NUL, so let's add +1 to take it into account
    // when I allocate the buffer for reading value names.
    maxValueNameLen++;

    // Preallocate a buffer for the value names
    auto nameBuffer = std::make_unique<wchar_t[]>(maxValueNameLen);

    // The value names and types will be stored here
    std::vector<std::pair<std::wstring, DWORD>> valueInfo;

    // Reserve room in the vector to speed up the following insertion loop
    valueInfo.reserve(valueCount);

    // Enumerate all the values
    for (DWORD index = 0; index < valueCount; index++)
    {
        // Get the name and the type of the current value
        DWORD valueNameLen = maxValueNameLen;
        DWORD valueType = 0;
        retCode = ::RegEnumValueW(
            m_hKey,
            index,
            nameBuffer.get(),
            &valueNameLen,
            nullptr,    // reserved
            &valueType,
            nullptr,    // no data
            nullptr     // no data size
        );
        if (retCode != ERROR_SUCCESS)
        {
            throw RegException{ retCode, "Cannot enumerate values: RegEnumValueW failed." };
        }

        // On success, the RegEnumValue API writes the length of the
        // value name in the valueNameLen output parameter
        // (not including the terminating NUL).
        // So we can build a wstring based on that.
        valueInfo.emplace_back(
            std::wstring{ nameBuffer.get(), valueNameLen },
            valueType
        );
    }

    return valueInfo;
}


inline RegExpected<std::vector<std::wstring>> RegKey::TryEnumSubKeys() const
{
    _ASSERTE(IsValid());

    using ReturnType = std::vector<std::wstring>;

    // Get some useful enumeration info, like the total number of subkeys
    // and the maximum length of the subkey names
    DWORD subKeyCount = 0;
    DWORD maxSubKeyNameLen = 0;
    LSTATUS retCode = ::RegQueryInfoKeyW(
        m_hKey,
        nullptr,    // no user-defined class
        nullptr,    // no user-defined class size
        nullptr,    // reserved
        &subKeyCount,
        &maxSubKeyNameLen,
        nullptr,    // no subkey class length
        nullptr,    // no value count
        nullptr,    // no value name max length
        nullptr,    // no max value length
        nullptr,    // no security descriptor
        nullptr     // no last write time
    );
    if (retCode != ERROR_SUCCESS)
    {
        return winreg_internal::MakeRegExpectedWithError<ReturnType>(retCode);
    }

    // NOTE: According to the MSDN documentation, the size returned for subkey name max length
    // does *not* include the terminating NUL, so let's add +1 to take it into account
    // when I allocate the buffer for reading subkey names.
    maxSubKeyNameLen++;

    // Preallocate a buffer for the subkey names
    auto nameBuffer = std::make_unique<wchar_t[]>(maxSubKeyNameLen);

    // The result subkey names will be stored here
    std::vector<std::wstring> subkeyNames;

    // Reserve room in the vector to speed up the following insertion loop
    subkeyNames.reserve(subKeyCount);

    // Enumerate all the subkeys
    for (DWORD index = 0; index < subKeyCount; index++)
    {
        // Get the name of the current subkey
        DWORD subKeyNameLen = maxSubKeyNameLen;
        retCode = ::RegEnumKeyExW(
            m_hKey,
            index,
            nameBuffer.get(),
            &subKeyNameLen,
            nullptr, // reserved
            nullptr, // no class
            nullptr, // no class
            nullptr  // no last write time
        );
        if (retCode != ERROR_SUCCESS)
        {
            return winreg_internal::MakeRegExpectedWithError<ReturnType>(retCode);
        }

        // On success, the ::RegEnumKeyEx API writes the length of the
        // subkey name in the subKeyNameLen output parameter
        // (not including the terminating NUL).
        // So I can build a wstring based on that length.
        subkeyNames.emplace_back(nameBuffer.get(), subKeyNameLen);
    }

    return RegExpected<ReturnType>{ subkeyNames };
}


inline RegExpected<std::vector<std::pair<std::wstring, DWORD>>> RegKey::TryEnumValues() const
{
    _ASSERTE(IsValid());

    using ReturnType = std::vector<std::pair<std::wstring, DWORD>>;

    // Get useful enumeration info, like the total number of values
    // and the maximum length of the value names
    DWORD valueCount = 0;
    DWORD maxValueNameLen = 0;
    LSTATUS retCode = ::RegQueryInfoKeyW(
        m_hKey,
        nullptr,    // no user-defined class
        nullptr,    // no user-defined class size
        nullptr,    // reserved
        nullptr,    // no subkey count
        nullptr,    // no subkey max length
        nullptr,    // no subkey class length
        &valueCount,
        &maxValueNameLen,
        nullptr,    // no max value length
        nullptr,    // no security descriptor
        nullptr     // no last write time
    );
    if (retCode != ERROR_SUCCESS)
    {
        return winreg_internal::MakeRegExpectedWithError<ReturnType>(retCode);
    }

    // NOTE: According to the MSDN documentation, the size returned for value name max length
    // does *not* include the terminating NUL, so let's add +1 to take it into account
    // when I allocate the buffer for reading value names.
    maxValueNameLen++;

    // Preallocate a buffer for the value names
    auto nameBuffer = std::make_unique<wchar_t[]>(maxValueNameLen);

    // The value names and types will be stored here
    std::vector<std::pair<std::wstring, DWORD>> valueInfo;

    // Reserve room in the vector to speed up the following insertion loop
    valueInfo.reserve(valueCount);

    // Enumerate all the values
    for (DWORD index = 0; index < valueCount; index++)
    {
        // Get the name and the type of the current value
        DWORD valueNameLen = maxValueNameLen;
        DWORD valueType = 0;
        retCode = ::RegEnumValueW(
            m_hKey,
            index,
            nameBuffer.get(),
            &valueNameLen,
            nullptr,    // reserved
            &valueType,
            nullptr,    // no data
            nullptr     // no data size
        );
        if (retCode != ERROR_SUCCESS)
        {
            return winreg_internal::MakeRegExpectedWithError<ReturnType>(retCode);
        }

        // On success, the RegEnumValue API writes the length of the
        // value name in the valueNameLen output parameter
        // (not including the terminating NUL).
        // So we can build a wstring based on that.
        valueInfo.emplace_back(
            std::wstring{ nameBuffer.get(), valueNameLen },
            valueType
        );
    }

    return RegExpected<ReturnType>{ valueInfo };
}


inline DWORD RegKey::QueryValueType(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    DWORD typeId = 0;     // will be returned by RegQueryValueEx

    LSTATUS retCode = ::RegQueryValueExW(
        m_hKey,
        valueName.c_str(),
        nullptr,    // reserved
        &typeId,
        nullptr,    // not interested
        nullptr     // not interested
    );

    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot get the value type: RegQueryValueExW failed." };
    }

    return typeId;
}


inline RegExpected<DWORD> RegKey::TryQueryValueType(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    using ReturnType = DWORD;

    DWORD typeId = 0;     // will be returned by RegQueryValueEx

    LSTATUS retCode = ::RegQueryValueExW(
        m_hKey,
        valueName.c_str(),
        nullptr,    // reserved
        &typeId,
        nullptr,    // not interested
        nullptr     // not interested
    );

    if (retCode != ERROR_SUCCESS)
    {
        return winreg_internal::MakeRegExpectedWithError<ReturnType>(retCode);
    }

    return RegExpected<ReturnType>{ typeId };
}


inline RegKey::InfoKey RegKey::QueryInfoKey() const
{
    _ASSERTE(IsValid());

    InfoKey infoKey{};
    LSTATUS retCode = ::RegQueryInfoKeyW(
        m_hKey,
        nullptr,
        nullptr,
        nullptr,
        &(infoKey.NumberOfSubKeys),
        nullptr,
        nullptr,
        &(infoKey.NumberOfValues),
        nullptr,
        nullptr,
        nullptr,
        &(infoKey.LastWriteTime)
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegQueryInfoKeyW failed." };
    }

    return infoKey;
}


inline RegExpected<RegKey::InfoKey> RegKey::TryQueryInfoKey() const
{
    _ASSERTE(IsValid());

    using ReturnType = RegKey::InfoKey;

    InfoKey infoKey{};
    LSTATUS retCode = ::RegQueryInfoKeyW(
        m_hKey,
        nullptr,
        nullptr,
        nullptr,
        &(infoKey.NumberOfSubKeys),
        nullptr,
        nullptr,
        &(infoKey.NumberOfValues),
        nullptr,
        nullptr,
        nullptr,
        &(infoKey.LastWriteTime)
    );
    if (retCode != ERROR_SUCCESS)
    {
        return winreg_internal::MakeRegExpectedWithError<ReturnType>(retCode);
    }

    return RegExpected<ReturnType>{ infoKey };
}


inline RegKey::KeyReflection RegKey::QueryReflectionKey() const
{
    BOOL isReflectionDisabled = FALSE;
    LSTATUS retCode = ::RegQueryReflectionKey(m_hKey, &isReflectionDisabled);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegQueryReflectionKey failed." };
    }

    return (isReflectionDisabled ? KeyReflection::ReflectionDisabled
                                 : KeyReflection::ReflectionEnabled);
}


inline RegExpected<RegKey::KeyReflection> RegKey::TryQueryReflectionKey() const
{
    using ReturnType = RegKey::KeyReflection;

    BOOL isReflectionDisabled = FALSE;
    LSTATUS retCode = ::RegQueryReflectionKey(m_hKey, &isReflectionDisabled);
    if (retCode != ERROR_SUCCESS)
    {
        return winreg_internal::MakeRegExpectedWithError<ReturnType>(retCode);
    }

    KeyReflection keyReflection = isReflectionDisabled ? KeyReflection::ReflectionDisabled
                                                       : KeyReflection::ReflectionEnabled;
    return RegExpected<ReturnType>{ keyReflection };
}


inline void RegKey::DeleteValue(const std::wstring& valueName)
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegDeleteValueW(m_hKey, valueName.c_str());
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegDeleteValueW failed." };
    }
}


inline RegResult RegKey::TryDeleteValue(const std::wstring& valueName) noexcept
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegDeleteValueW(m_hKey, valueName.c_str()) };
}


inline void RegKey::DeleteKey(const std::wstring& subKey, const REGSAM desiredAccess)
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegDeleteKeyExW(m_hKey, subKey.c_str(), desiredAccess, 0);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegDeleteKeyExW failed." };
    }
}


inline RegResult RegKey::TryDeleteKey(const std::wstring& subKey,
                                      const REGSAM desiredAccess) noexcept
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegDeleteKeyExW(m_hKey, subKey.c_str(), desiredAccess, 0) };
}


inline void RegKey::DeleteTree(const std::wstring& subKey)
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegDeleteTreeW(m_hKey, subKey.c_str());
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegDeleteTreeW failed." };
    }
}


inline RegResult RegKey::TryDeleteTree(const std::wstring& subKey) noexcept
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegDeleteTreeW(m_hKey, subKey.c_str()) };
}


inline void RegKey::CopyTree(const std::wstring& sourceSubKey, const RegKey& destKey)
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegCopyTreeW(m_hKey, sourceSubKey.c_str(), destKey.Get());
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegCopyTreeW failed." };
    }
}


inline RegResult RegKey::TryCopyTree(const std::wstring& sourceSubKey,
                                     const RegKey& destKey) noexcept
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegCopyTreeW(m_hKey, sourceSubKey.c_str(), destKey.Get()) };
}


inline void RegKey::FlushKey()
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegFlushKey(m_hKey);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegFlushKey failed." };
    }
}


inline RegResult RegKey::TryFlushKey() noexcept
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegFlushKey(m_hKey) };
}


inline void RegKey::LoadKey(const std::wstring& subKey, const std::wstring& filename)
{
    Close();

    LSTATUS retCode = ::RegLoadKeyW(m_hKey, subKey.c_str(), filename.c_str());
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegLoadKeyW failed." };
    }
}


inline RegResult RegKey::TryLoadKey(const std::wstring& subKey,
                                    const std::wstring& filename) noexcept
{
    Close();

    return RegResult{ ::RegLoadKeyW(m_hKey, subKey.c_str(), filename.c_str()) };
}


inline void RegKey::SaveKey(
    const std::wstring& filename,
    SECURITY_ATTRIBUTES* const securityAttributes
) const
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegSaveKeyW(m_hKey, filename.c_str(), securityAttributes);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegSaveKeyW failed." };
    }
}


inline RegResult RegKey::TrySaveKey(
    const std::wstring& filename,
    SECURITY_ATTRIBUTES* const securityAttributes
) const noexcept
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegSaveKeyW(m_hKey, filename.c_str(), securityAttributes) };
}


inline void RegKey::EnableReflectionKey()
{
    LSTATUS retCode = ::RegEnableReflectionKey(m_hKey);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegEnableReflectionKey failed." };
    }
}


inline RegResult RegKey::TryEnableReflectionKey() noexcept
{
    return RegResult{ ::RegEnableReflectionKey(m_hKey) };
}


inline void RegKey::DisableReflectionKey()
{
    LSTATUS retCode = ::RegDisableReflectionKey(m_hKey);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegDisableReflectionKey failed." };
    }
}


inline RegResult RegKey::TryDisableReflectionKey() noexcept
{
    return RegResult{ ::RegDisableReflectionKey(m_hKey) };
}


inline void RegKey::ConnectRegistry(const std::wstring& machineName, const HKEY hKeyPredefined)
{
    // Safely close any previously opened key
    Close();

    HKEY hKeyResult = nullptr;
    LSTATUS retCode = ::RegConnectRegistryW(machineName.c_str(), hKeyPredefined, &hKeyResult);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegConnectRegistryW failed." };
    }

    // Take ownership of the result key
    m_hKey = hKeyResult;
}


inline RegResult RegKey::TryConnectRegistry(const std::wstring& machineName,
                                            const HKEY hKeyPredefined) noexcept
{
    // Safely close any previously opened key
    Close();

    HKEY hKeyResult = nullptr;
    RegResult retCode{ ::RegConnectRegistryW(machineName.c_str(), hKeyPredefined, &hKeyResult) };
    if (retCode.Failed())
    {
        return retCode;
    }

    // Take ownership of the result key
    m_hKey = hKeyResult;

    _ASSERTE(retCode.IsOk());
    return retCode;
}


inline std::wstring RegKey::RegTypeToString(const DWORD regType)
{
    switch (regType)
    {
        case REG_SZ:        return L"REG_SZ";
        case REG_EXPAND_SZ: return L"REG_EXPAND_SZ";
        case REG_MULTI_SZ:  return L"REG_MULTI_SZ";
        case REG_DWORD:     return L"REG_DWORD";
        case REG_QWORD:     return L"REG_QWORD";
        case REG_BINARY:    return L"REG_BINARY";

        default:            return L"Unknown/unsupported registry type";
    }
}


//------------------------------------------------------------------------------
//              Internal RegKey Helper Methods Implementation
//------------------------------------------------------------------------------


#if defined(GIOVANNI_DICANIO_WINREG_WINDOWS_XP_COMPATIBLE)


namespace winreg_internal
{

//
// In the version of the code that uses Vista+ RegGetValue (and as such is not compatible
// with Windows XP), I don't need to manually expand environment strings read from the registry,
// as RegGetValue automatically expands them, unless the RRF_NOEXPAND flag is passed.
// So, the following custom winreg_internal::ExpandEnvironmentStrings helper function is only needed
// in the Windows XP compatible build.
//

//------------------------------------------------------------------------------
// Expands environment-variable strings and replaces them with the values
// defined for the current user.
// Basically, wraps the ::ExpandEnvironmentStrings API.
//------------------------------------------------------------------------------
[[nodiscard]] inline std::wstring ExpandEnvironmentStrings(const std::wstring& source)
{
    DWORD retCode = 0;
    DWORD bufferLength = MAX_PATH; // An arbitrary initial length that might be reasonable
    std::unique_ptr<wchar_t[]> buffer;

    do {
        buffer = std::make_unique<wchar_t[]>(bufferLength);

        retCode = ::ExpandEnvironmentStringsW(source.c_str(), buffer.get(), bufferLength);
        if (retCode == 0)
        {
            DWORD errorCode = ::GetLastError();
            throw RegException{
                static_cast<LSTATUS>(errorCode),
                "Cannot expand environment strings: ExpandEnvironmentStringsW failed."
            };
        }

        // If ExpandEnvironmentStrings needs a destination buffer of larger size,
        // this larger size value is returned.
        if (retCode > bufferLength)
        {
            bufferLength = retCode;
        }

    } while (retCode > bufferLength);

    // 'result' includes the terminating NUL, so we have to remove it to build a wstring object;
    // this is why I used (retCode - 1)
    return std::wstring(buffer.get(), retCode - 1);
}

} // namespace winreg_internal


//------------------------------------------------------------------------
// Code that uses the RegQueryValueEx Registry API instead of RegGetValue,
// and as such is compatible with Windows XP.
//------------------------------------------------------------------------

inline LSTATUS RegKey::GetDwordValueInternal(
    DWORD& result,
    const std::wstring& valueName,
    std::string* errorMessage
) const
{
    _ASSERTE(IsValid());

    // Clear output parameters
    result = 0;
    if (errorMessage != nullptr)
    {
        errorMessage->clear();
    }

    DWORD data = 0;                  // to be read from the registry
    DWORD dataSize = sizeof(data);   // size of data, in bytes
    DWORD valueType = 0;

    LSTATUS retCode = ::RegQueryValueExW(
        m_hKey,
        valueName.c_str(),
        nullptr,        // reserved
        &valueType,
        reinterpret_cast<BYTE*>(&data),
        &dataSize
    );

    if (retCode != ERROR_SUCCESS)
    {
        if (errorMessage != nullptr)
        {
            *errorMessage = "Cannot get DWORD value: RegQueryValueExW failed.";
        }

        return retCode;
    }

    if (valueType != REG_DWORD)
    {
        if (errorMessage != nullptr)
        {
            *errorMessage = "Cannot get DWORD value: type is not REG_DWORD.";
        }

        return ERROR_INVALID_DATATYPE;
    }

    result = data;
    return ERROR_SUCCESS;
}


inline LSTATUS RegKey::GetQwordValueInternal(
    ULONGLONG& result,
    const std::wstring& valueName,
    std::string* errorMessage
) const
{
    _ASSERTE(IsValid());

    // Clear output parameters
    result = 0;
    if (errorMessage != nullptr)
    {
        errorMessage->clear();
    }

    ULONGLONG data = 0;              // to be read from the registry
    DWORD dataSize = sizeof(data);   // size of data, in bytes
    DWORD valueType = 0;

    LSTATUS retCode = ::RegQueryValueExW(
        m_hKey,
        valueName.c_str(),
        nullptr,        // reserved
        &valueType,
        reinterpret_cast<BYTE*>(&data),
        &dataSize
    );

    if (retCode != ERROR_SUCCESS)
    {
        if (errorMessage != nullptr)
        {
            *errorMessage = "Cannot get QWORD value: RegQueryValueExW failed.";
        }

        return retCode;
    }

    if (valueType != REG_QWORD)
    {
        if (errorMessage != nullptr)
        {
            *errorMessage = "Cannot get QWORD value: type is not REG_QWORD.";
        }

        return ERROR_INVALID_DATATYPE;
    }

    result = data;
    return ERROR_SUCCESS;
}


namespace winreg_internal
{

inline LSTATUS GetStringOrExpandStringValueWinXPCompatible(
    HKEY         const  hKey,               // [in]  Registry key to operate on
    std::wstring&       result,             // [out] String or expand string read from the registry
    const std::wstring& valueName,          // [in]  Name of the string value to be read
    DWORD        const  desiredValueType,   // [in]  REG_SZ or REG_EXPAND_SZ
    std::string* const  errorMessage        // [out, optional] Descriptive error message on failure
)
{
    // Clear output parameters
    result.clear();
    if (errorMessage != nullptr)
    {
        errorMessage->clear();
    }

    _ASSERTE((desiredValueType == REG_SZ) || (desiredValueType == REG_EXPAND_SZ));

    std::wstring stringData;    // to be read from the registry
    DWORD dataSize = 0;         // size of the string data, in bytes
    DWORD valueType = 0;

    LSTATUS retCode = ERROR_MORE_DATA;

    while (retCode == ERROR_MORE_DATA)
    {
        // Get the size of the result string
        retCode = ::RegQueryValueExW(
            hKey,
            valueName.c_str(),
            nullptr,            // reserved
            &valueType,
            nullptr,            // output buffer not needed now
            &dataSize
        );

        if (retCode != ERROR_SUCCESS)
        {
            if (errorMessage != nullptr)
            {
                *errorMessage = desiredValueType == REG_SZ ?
                    "Cannot get the size of the string value: RegQueryValueExW failed." :
                    "Cannot get the size of the expand string value: RegQueryValueExW failed.";
            }

            return retCode;
        }

        if (valueType != desiredValueType)
        {
            if (errorMessage != nullptr)
            {
                *errorMessage = desiredValueType == REG_SZ ?
                    "Cannot get string value: type is not REG_SZ." :
"Cannot get expand string value: type is not REG_EXPAND_SZ.";
            }

            return ERROR_INVALID_DATATYPE;
        }

        if (dataSize == 0)
        {
            // return std::wstring{};
            _ASSERTE(result.empty());
            return ERROR_SUCCESS;
        }

        // Make room for the NUL terminator, in case it's not written by RegQueryValueEx.
        // The MSDN documentation warns about RegQueryValueEx not always returning
        // NUL-terminated strings.
        dataSize += sizeof(wchar_t);

        // Allocate a string of proper size.
        // Note that dataSize is in bytes and includes the terminating NUL;
        // we have to convert the size from bytes to wchar_ts for wstring::resize.
        stringData.resize(dataSize / sizeof(wchar_t));

        // Call RegQueryValueEx for the second time to read the string's content
        retCode = ::RegQueryValueExW(
            hKey,
            valueName.c_str(),
            nullptr,                                     // reserved
            nullptr,                                     // type not needed
            reinterpret_cast<BYTE*>(stringData.data()),  // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        if (errorMessage != nullptr)
        {
            *errorMessage = desiredValueType == REG_SZ ?
                "Cannot get the string value: RegQueryValueExW failed." :
                "Cannot get the expand string value: RegQueryValueExW failed.";
        }

        return retCode;
    }

    // Resize the string to the actual characters written by RegQueryValueEx.
    // Note that dataSize is in bytes and may include the terminating NUL;
    // we have to convert the size from bytes to wchar_ts for wstring::resize.
    const DWORD wcharsWritten = dataSize / sizeof(wchar_t);
    stringData.resize(wcharsWritten);

    // Special case of empty string
    if (wcharsWritten == 0)
    {
        _ASSERTE(result.empty());
        return ERROR_SUCCESS;
    }


    //
    // Remove the NUL terminator scribbled by RegQueryValueEx from the wstring, *if* present
    //
    //
    // Case #1: RegQueryValueEx writes the \0:
    //
    // H e l l o \0  <--- \0 written by RegQueryValueEx
    // 0 1 2 3 4 5
    //
    // wcharsWritten = 6
    //
    // data[ wcharsWritten - 1 ] == '\0'  <-- we need to chop this \0 off
    //
    // We need to chop off the \0, as wstring objects *already* have their own \0.
    //
    //
    // Case #2: RegQueryValueEx does *not* write the \0:
    //
    // If RegQueryValueEx did *not* write a terminating NUL
    // (which is a possibility, according to the MSDN doc of RegQueryValueEx):
    //
    //      result[ wcharsWritten - 1 ] != '\0'
    //
    // In this case we don't have to remove any terminating '\0'
    // (wstring's internal '\0' will be fine).
    //

    if (stringData[wcharsWritten - 1] == L'\0')
    {
        // Chop off the null
        stringData.resize(wcharsWritten - 1);
    }

    // result <= stringData
    std::swap(result, stringData);
    return ERROR_SUCCESS;
}

} // namespace winreg_internal


inline LSTATUS RegKey::GetStringValueInternal(
    std::wstring&       result,
    const std::wstring& valueName,
    std::string* const  errorMessage
) const
{
    _ASSERTE(IsValid());

    // Clear output parameters
    result.clear();
    if (errorMessage != nullptr)
    {
        errorMessage->clear();
    }

    return winreg_internal::GetStringOrExpandStringValueWinXPCompatible(
        m_hKey,
        result,
        valueName,
        REG_SZ,
        errorMessage
    );
}


inline LSTATUS RegKey::GetExpandStringValueInternal(
    std::wstring&               result,
    const std::wstring&         valueName,
    ExpandStringOption const    expandOption,
    std::string*       const    errorMessage
) const
{
    _ASSERTE(IsValid());

    // Clear output parameters
    result.clear();
    if (errorMessage != nullptr)
    {
        errorMessage->clear();
    }

    std::wstring stringMayNeedExpansion;
    LSTATUS retCode = winreg_internal::GetStringOrExpandStringValueWinXPCompatible(
        m_hKey,
        stringMayNeedExpansion,
        valueName,
        REG_EXPAND_SZ,
        errorMessage
    );
    if (retCode != ERROR_SUCCESS)
    {
        return retCode;
    }

    // Expand the environment strings if required by the caller
    if (expandOption == ExpandStringOption::Expand)
    {
        result = winreg_internal::ExpandEnvironmentStrings(stringMayNeedExpansion);
    }
    else
    {
        // Environment string expansion not required by the caller
        result = stringMayNeedExpansion;
    }

    return ERROR_SUCCESS;
}


inline LSTATUS RegKey::GetMultiStringValueInternal(
    std::vector<std::wstring>& result,
    const std::wstring& valueName,
    std::string* const errorMessage
) const
{
    _ASSERTE(IsValid());

    // Clear output parameters
    result.clear();
    if (errorMessage != nullptr)
    {
        errorMessage->clear();
    }


    // Room for the result multi-string, to be read from the registry
    std::vector<wchar_t> multiString;

    // Size of the multi-string, in bytes
    DWORD dataSize = 0;

    DWORD valueType = 0;

    LSTATUS retCode = ERROR_MORE_DATA;

    while (retCode == ERROR_MORE_DATA)
    {
        // Request the size of the multi-string, in bytes
        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,            // reserved
            &valueType,
            nullptr,            // output buffer not needed now
            &dataSize
        );

        if (retCode != ERROR_SUCCESS)
        {
            if (errorMessage != nullptr)
            {
                *errorMessage = "Cannot get the size of the multi-string value: RegQueryValueExW failed.";
            }

            return retCode;
        }

        if (valueType != REG_MULTI_SZ)
        {
            if (errorMessage != nullptr)
            {
                *errorMessage = "Cannot get the multi-string string value : type is not REG_MULTI_SZ.";
            }

            return ERROR_INVALID_DATATYPE;
        }

        if (dataSize == 0)
        {
            result.clear();
            return ERROR_SUCCESS;
        }

        // Make room for two NUL terminators, in case they are not written by RegQueryValueEx.
        // The MSDN documentation warns about RegQueryValueEx not always returning
        // double-NUL-terminated strings.
        dataSize += (2 * sizeof(wchar_t));

        // Allocate room for the result multi-string.
        // Note that dataSize is in bytes, but our vector<wchar_t>::resize method requires size
        // to be expressed in wchar_ts.
        multiString.resize(dataSize / sizeof(wchar_t));

        // Call RegQueryValueEx for the second time to read the multi-string's content into the vector
        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,                                      // reserved
            nullptr,                                      // type not needed
            reinterpret_cast<BYTE*>(multiString.data()),  // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        if (errorMessage != nullptr)
        {
            *errorMessage = "Cannot get the multi-string value: RegQueryValueExW failed.";
        }

        return retCode;
    }

    // Resize vector to the actual size returned by the last call to RegQueryValueEx.
    // Note that the vector is a vector of wchar_ts, instead the size returned by RegQueryValueEx
    // is in bytes, so we have to scale from bytes to wchar_t count.
    multiString.resize(dataSize / sizeof(wchar_t));

    // Make sure that the vector is double-NUL-terminated.
    winreg_internal::DoubleNullTerminate(multiString);

    // Convert the double-null-terminated string structure to a vector<wstring>,
    // and return that back to the caller
    result = winreg_internal::ParseMultiString(multiString);

    return ERROR_SUCCESS;
}


inline LSTATUS RegKey::GetBinaryValueInternal(
    std::vector<BYTE>& result,
    const std::wstring& valueName,
    std::string* const errorMessage
) const
{
    _ASSERTE(IsValid());

    // Clear output parameters
    result.clear();
    if (errorMessage != nullptr)
    {
        errorMessage->clear();
    }


    // Room for the binary data, to be read from the registry
    std::vector<BYTE> binaryData;

    // Size of binary data, in bytes
    DWORD dataSize = 0;

    DWORD valueType = 0;

    LSTATUS retCode = ERROR_MORE_DATA;

    while (retCode == ERROR_MORE_DATA)
    {
        // Request the size of the binary data, in bytes
        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,            // reserved
            &valueType,
            nullptr,            // output buffer not needed now
            &dataSize
        );

        if (retCode != ERROR_SUCCESS)
        {
            if (errorMessage != nullptr)
            {
                *errorMessage = "Cannot get the size of the binary data: RegQueryValueExW failed.";
            }

            return retCode;
        }

        if (valueType != REG_BINARY)
        {
            if (errorMessage != nullptr)
            {
                *errorMessage = "Cannot get the binary data: type is not REG_BINARY.";
            }

            return ERROR_INVALID_DATATYPE;
        }

        if (dataSize == 0)
        {
            result.clear();
            return ERROR_SUCCESS;
        }

        // Allocate a buffer of proper size to store the binary data
        binaryData.resize(dataSize);

        // Call RegQueryValueEx for the second time to read the binary data content into the vector
        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,               // reserved
            nullptr,               // type not needed
            binaryData.data(),     // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        if (errorMessage != nullptr)
        {
            *errorMessage = "Cannot get binary data: RegQueryValueExW failed.";
        }

        return retCode;
    }

    // Resize vector to the actual size returned by the last call to RegQueryValueEx
    binaryData.resize(dataSize);

    // result <= binaryData
    std::swap(result, binaryData);

    return ERROR_SUCCESS;
}

#else


//------------------------------------------------------------------------
// Code that uses the RegGetValue Registry API available in Vista+,
// and as such is *not* compatible with Windows XP.
//------------------------------------------------------------------------

inline LSTATUS RegKey::GetDwordValueInternal(
    DWORD& result,
    const std::wstring& valueName,
    std::string* errorMessage
) const
{
    _ASSERTE(IsValid());

    // Clear output parameters
    result = 0;
    if (errorMessage != nullptr)
    {
        errorMessage->clear();
    }

    DWORD data = 0;                  // to be read from the registry
    DWORD dataSize = sizeof(data);   // size of data, in bytes

    constexpr DWORD flags = RRF_RT_REG_DWORD;
    LSTATUS retCode = ::RegGetValueW(
        m_hKey,
        nullptr, // no subkey
        valueName.c_str(),
        flags,
        nullptr, // type not required
        &data,
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        // throw RegException{ retCode, "Cannot get DWORD value: RegGetValueW failed." };
        if (errorMessage != nullptr)
        {
            *errorMessage = "Cannot get DWORD value: RegGetValueW failed.";
        }

        return retCode;
    }

    result = data;
    return ERROR_SUCCESS;
}


inline LSTATUS RegKey::GetQwordValueInternal(
    ULONGLONG& result,
    const std::wstring& valueName,
    std::string* errorMessage
) const
{
    _ASSERTE(IsValid());

    // Clear output parameters
    result = 0;
    if (errorMessage != nullptr)
    {
        errorMessage->clear();
    }

    ULONGLONG data = 0;              // to be read from the registry
    DWORD dataSize = sizeof(data);   // size of data, in bytes

    constexpr DWORD flags = RRF_RT_REG_QWORD;
    LSTATUS retCode = ::RegGetValueW(
        m_hKey,
        nullptr, // no subkey
        valueName.c_str(),
        flags,
        nullptr, // type not required
        &data,
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        // throw RegException{ retCode, "Cannot get QWORD value: RegGetValueW failed." };
        if (errorMessage != nullptr)
        {
            *errorMessage = "Cannot get QWORD value: RegGetValueW failed.";
        }

        return retCode;
    }

    result = data;
    return ERROR_SUCCESS;
}


inline LSTATUS RegKey::GetStringValueInternal(
    std::wstring&       result,
    const std::wstring& valueName,
    std::string* const  errorMessage
) const
{
    _ASSERTE(IsValid());

    // Clear output parameters
    result.clear();
    if (errorMessage != nullptr)
    {
        errorMessage->clear();
    }


    std::wstring stringData;    // to be read from the registry
    DWORD dataSize = 0;         // size of the string data, in bytes

    constexpr DWORD flags = RRF_RT_REG_SZ;

    LSTATUS retCode = ERROR_MORE_DATA;

    while (retCode == ERROR_MORE_DATA)
    {
        // Get the size of the result string
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueName.c_str(),
            flags,
            nullptr,    // type not required
            nullptr,    // output buffer not needed now
            &dataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            // throw RegException{ retCode, "Cannot get the size of the string value: RegGetValueW failed." };
            if (errorMessage != nullptr)
            {
                *errorMessage = "Cannot get the size of the string value: RegGetValueW failed.";
            }

            return retCode;
        }

        // Allocate a string of proper size.
        // Note that dataSize is in bytes and includes the terminating NUL;
        // we have to convert the size from bytes to wchar_ts for wstring::resize.
        stringData.resize(dataSize / sizeof(wchar_t));

        // Call RegGetValue for the second time to read the string's content
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,            // no subkey
            valueName.c_str(),
            flags,
            nullptr,            // type not required
            stringData.data(),  // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        // throw RegException{ retCode, "Cannot get the string value: RegGetValueW failed." };
        if (errorMessage != nullptr)
        {
            *errorMessage = "Cannot get the string value: RegGetValueW failed.";
        }

        return retCode;
    }

    // Remove the NUL terminator scribbled by RegGetValue from the wstring
    stringData.resize((dataSize / sizeof(wchar_t)) - 1);


    // result <= stringData
    std::swap(result, stringData);
    return ERROR_SUCCESS;
}


inline LSTATUS RegKey::GetExpandStringValueInternal(
    std::wstring& result,
    const std::wstring& valueName,
    ExpandStringOption expandOption,
    std::string* const  errorMessage
) const
{
    _ASSERTE(IsValid());

    // Clear output parameters
    result.clear();
    if (errorMessage != nullptr)
    {
        errorMessage->clear();
    }


    std::wstring stringData;    // to be read from the registry
    DWORD dataSize = 0;         // size of the string data, in bytes

    const DWORD flags = [expandOption]() {
        DWORD flags = RRF_RT_REG_EXPAND_SZ;

        // Adjust the flag for RegGetValue considering the expand string option specified by the caller
        if (expandOption == ExpandStringOption::DontExpand)
        {
            flags |= RRF_NOEXPAND;
        }

        return flags;
    }();

    LSTATUS retCode = ERROR_MORE_DATA;

    while (retCode == ERROR_MORE_DATA)
    {
        // Get the size of the result string
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueName.c_str(),
            flags,
            nullptr,    // type not required
            nullptr,    // output buffer not needed now
            &dataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            if (errorMessage != nullptr)
            {
                *errorMessage = "Cannot get the size of the expand string value: RegGetValueW failed.";
            }

            return retCode;
        }

        // Allocate a string of proper size.
        // Note that dataSize is in bytes and includes the terminating NUL;
        // we have to convert the size from bytes to wchar_ts for wstring::resize.
        stringData.resize(dataSize / sizeof(wchar_t));

        // Call RegGetValue for the second time to read the string's content
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,            // no subkey
            valueName.c_str(),
            flags,
            nullptr,            // type not required
            stringData.data(),  // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        if (errorMessage != nullptr)
        {
            *errorMessage = "Cannot get the expand string value: RegGetValueW failed.";
        }

        return retCode;
    }

    // Remove the NUL terminator scribbled by RegGetValue from the wstring
    stringData.resize((dataSize / sizeof(wchar_t)) - 1);


    // result <= stringData
    std::swap(result, stringData);
    return ERROR_SUCCESS;
}


inline LSTATUS RegKey::GetMultiStringValueInternal(
    std::vector<std::wstring>& result,
    const std::wstring& valueName,
    std::string* const errorMessage
) const
{
    _ASSERTE(IsValid());

    // Clear output parameters
    result.clear();
    if (errorMessage != nullptr)
    {
        errorMessage->clear();
    }


    // Room for the result multi-string, to be read from the registry
    std::vector<wchar_t> multiString;

    // Size of the multi-string, in bytes
    DWORD dataSize = 0;

    constexpr DWORD flags = RRF_RT_REG_MULTI_SZ;

    LSTATUS retCode = ERROR_MORE_DATA;


    while (retCode == ERROR_MORE_DATA)
    {
        // Request the size of the multi-string, in bytes
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueName.c_str(),
            flags,
            nullptr,    // type not required
            nullptr,    // output buffer not needed now
            &dataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            if (errorMessage != nullptr)
            {
                *errorMessage = "Cannot get the size of the multi-string value: RegGetValueW failed.";
            }

            return retCode;
        }

        // Allocate room for the result multi-string.
        // Note that dataSize is in bytes, but our vector<wchar_t>::resize method requires size
        // to be expressed in wchar_ts.
        multiString.resize(dataSize / sizeof(wchar_t));

        // Call RegGetValue for the second time to read the multi-string's content into the vector
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,                // no subkey
            valueName.c_str(),
            flags,
            nullptr,                // type not required
            multiString.data(),     // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        if (errorMessage != nullptr)
        {
            *errorMessage = "Cannot get the size of the multi-string value: RegGetValueW failed.";
        }

        return retCode;
    }

    // Resize vector to the actual size returned by the last call to RegGetValue.
    // Note that the vector is a vector of wchar_ts, instead the size returned by RegGetValue
    // is in bytes, so we have to scale from bytes to wchar_t count.
    multiString.resize(dataSize / sizeof(wchar_t));

    // Convert the double-null-terminated string structure to a vector<wstring>,
    // and return that back to the caller
    result = winreg_internal::ParseMultiString(multiString);

    return ERROR_SUCCESS;
}


inline LSTATUS RegKey::GetBinaryValueInternal(
    std::vector<BYTE>& result,
    const std::wstring& valueName,
    std::string* const errorMessage
) const
{
    _ASSERTE(IsValid());

    // Clear output parameters
    result.clear();
    if (errorMessage != nullptr)
    {
        errorMessage->clear();
    }


    // Room for the binary data, to be read from the registry
    std::vector<BYTE> binaryData;

    // Size of binary data, in bytes
    DWORD dataSize = 0;

    constexpr DWORD flags = RRF_RT_REG_BINARY;

    LSTATUS retCode = ERROR_MORE_DATA;

    while (retCode == ERROR_MORE_DATA)
    {
        // Request the size of the binary data, in bytes
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueName.c_str(),
            flags,
            nullptr,    // type not required
            nullptr,    // output buffer not needed now
            &dataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            if (errorMessage != nullptr)
            {
                *errorMessage = "Cannot get the size of the binary data: RegGetValueW failed.";
            }

            return retCode;
        }

        // Allocate a buffer of proper size to store the binary data
        binaryData.resize(dataSize);

        // Handle the special case of zero-length binary data:
        // If the binary data value in the registry is empty, just return an empty vector.
        if (dataSize == 0)
        {
            _ASSERTE(binaryData.empty());
            _ASSERTE(result.empty());
            return ERROR_SUCCESS;
        }

        // Call RegGetValue for the second time to read the binary data content into the vector
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,            // no subkey
            valueName.c_str(),
            flags,
            nullptr,            // type not required
            binaryData.data(),  // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        if (errorMessage != nullptr)
        {
            *errorMessage = "Cannot get the binary data: RegGetValueW failed.";
        }

        return retCode;
    }

    // Resize vector to the actual size returned by the last call to RegGetValue
    binaryData.resize(dataSize);

    // result <= binaryData
    std::swap(result, binaryData);

    return ERROR_SUCCESS;
}


#endif // GIOVANNI_DICANIO_WINREG_WINDOWS_XP_COMPATIBLE


//------------------------------------------------------------------------------
//                          RegException Inline Methods
//------------------------------------------------------------------------------

inline RegException::RegException(const LSTATUS errorCode, const char* const message)
    : std::system_error{ errorCode, std::system_category(), message }
{}


inline RegException::RegException(const LSTATUS errorCode, const std::string& message)
    : std::system_error{ errorCode, std::system_category(), message }
{}


//------------------------------------------------------------------------------
//                          RegResult Inline Methods
//------------------------------------------------------------------------------

inline RegResult::RegResult(const LSTATUS result) noexcept
    : m_result{ result }
{}


inline bool RegResult::IsOk() const noexcept
{
    return m_result == ERROR_SUCCESS;
}


inline bool RegResult::Failed() const noexcept
{
    return m_result != ERROR_SUCCESS;
}


inline RegResult::operator bool() const noexcept
{
    return IsOk();
}


inline LSTATUS RegResult::Code() const noexcept
{
    return m_result;
}


inline std::wstring RegResult::ErrorMessage() const
{
    return ErrorMessage(MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT));
}


inline std::wstring RegResult::ErrorMessage(const DWORD languageId) const
{
    // Invoke FormatMessage() to retrieve the error message from Windows
    winreg_internal::ScopedLocalFree<wchar_t> messagePtr;
    DWORD retCode = ::FormatMessageW(
        FORMAT_MESSAGE_ALLOCATE_BUFFER |
        FORMAT_MESSAGE_FROM_SYSTEM |
        FORMAT_MESSAGE_IGNORE_INSERTS,
        nullptr,
        m_result,
        languageId,
        reinterpret_cast<LPWSTR>(messagePtr.AddressOf()),
        0,
        nullptr
    );
    if (retCode == 0)
    {
        // FormatMessage failed: return an empty string
        return std::wstring{};
    }

    // Safely copy the C-string returned by FormatMessage() into a std::wstring object,
    // and return it back to the caller.
    return std::wstring{ messagePtr.Get() };
}


//------------------------------------------------------------------------------
//                          RegExpected Inline Methods
//------------------------------------------------------------------------------

template <typename T>
inline RegExpected<T>::RegExpected(const RegResult& errorCode) noexcept
    : m_result{ errorCode }
{}


template <typename T>
inline RegExpected<T>::RegExpected(const T& value)
    : m_value{ value }
{}


template <typename T>
inline RegExpected<T>::RegExpected(T&& value)
    : m_value{ std::move(value) }
{}


template <typename T>
inline RegExpected<T>::operator bool() const noexcept
{
    return IsValid();
}


template <typename T>
inline bool RegExpected<T>::IsValid() const noexcept
{
    return m_result.IsOk();
}


template <typename T>
inline const T& RegExpected<T>::GetValue() const
{
    // Check that the object stores a valid value
    _ASSERTE(IsValid());

    if (!IsValid())
    {
        // TODO: May use another exception?
        throw std::invalid_argument{
            "winreg::RegExpected<T>::GetValue() called on an instance in error state."
        };
    }

    // If the object is in a valid state, return the value
    return m_value;
}


template <typename T>
inline RegResult RegExpected<T>::GetError() const
{
    return m_result;
}


} // namespace winreg


#endif // GIOVANNI_DICANIO_WINREG_HPP_INCLUDED

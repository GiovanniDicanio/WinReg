#ifndef GIOVANNI_DICANIO_WINREG_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_HPP_INCLUDED


////////////////////////////////////////////////////////////////////////////////
//
//      *** Modern C++ Wrappers Around Windows Registry C API ***
//
//               Copyright (C) by Giovanni Dicanio
//
// First version: 2017, January 22nd
// Last update:   2020, May 8th
//
// E-mail: <first name>.<last name> AT REMOVE_THIS gmail.com
//
// Registry key handles are safely and conveniently wrapped
// in the RegKey resource manager C++ class.
//
// Errors are signaled throwing exceptions of class RegException.
// In addition, there are also methods named following the style of TryAction
// (e.g. TryGetDwordValue), that try to perform the given action, and on failure
// return Windows Registry API error codes wrapped in the RegResult class.
//
// Unicode UTF-16 strings are represented using the std::wstring class;
// ATL's CString is not used, to avoid dependencies from ATL or MFC.
//
// Compiler: Visual Studio 2019
// Code compiles cleanly at /W4 on both 32-bit and 64-bit builds.
//
// ===========================================================================
//
// The MIT License(MIT)
//
// Copyright(c) 2017-2020 by Giovanni Dicanio
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

#include <string>           // std::wstring
#include <system_error>     // std::system_error
#include <utility>          // std::swap, std::pair
#include <vector>           // std::vector


//
// Support for Building as DLL
// ===========================
//
// When building as DLL, #define WINREG_API as follow, *before* including this file:
//
//      #define WINREG_API __declspec(dllexport)
//
// When consuming as DLL, #define WINREG_API as follow, *before* including this file:
//
//      #define WINREG_API __declspec(dllimport)
//
//
#ifndef WINREG_API
#define WINREG_API
#endif // WINREG_API


namespace winreg
{

// Forward class declarations
class RegException;
class RegResult;


//------------------------------------------------------------------------------
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
//------------------------------------------------------------------------------
class WINREG_API RegKey
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
    // Uses default KEY_READ|KEY_WRITE access.
    // For finer grained control, call the Create() method overloads.
    // Throw RegException on failure.
    RegKey(HKEY hKeyParent, const std::wstring& subKey);

    // Open the given registry key if it exists, else create a new key.
    // Allow the caller to specify the desired access to the key (e.g. KEY_READ
    // for read-only access).
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

    // Wrapper around RegCreateKeyEx, that allows you to specify desired access
    void Create(
        HKEY hKeyParent,
        const std::wstring& subKey,
        REGSAM desiredAccess = KEY_READ | KEY_WRITE
    );

    // Wrapper around RegCreateKeyEx, that allows you to specify desired access
    [[nodiscard]] RegResult TryCreate(
        HKEY hKeyParent,
        const std::wstring& subKey,
        REGSAM desiredAccess = KEY_READ | KEY_WRITE
    ) noexcept;

    // Wrapper around RegCreateKeyEx
    void Create(
        HKEY hKeyParent,
        const std::wstring& subKey,
        REGSAM desiredAccess,
        DWORD options,
        SECURITY_ATTRIBUTES* securityAttributes,
        DWORD* disposition
    );

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
    void Open(
        HKEY hKeyParent,
        const std::wstring& subKey,
        REGSAM desiredAccess = KEY_READ | KEY_WRITE
    );

    // Wrapper around RegOpenKeyEx
    [[nodiscard]] RegResult TryOpen(
        HKEY hKeyParent,
        const std::wstring& subKey,
        REGSAM desiredAccess = KEY_READ | KEY_WRITE
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

    [[nodiscard]] RegResult TrySetDwordValue(const std::wstring& valueName, DWORD data) noexcept;
    [[nodiscard]] RegResult TrySetQwordValue(const std::wstring& valueName,
                                             const ULONGLONG& data) noexcept;
    [[nodiscard]] RegResult TrySetStringValue(const std::wstring& valueName,
                                              const std::wstring& data) noexcept;
    [[nodiscard]] RegResult TrySetExpandStringValue(const std::wstring& valueName,
                                                    const std::wstring& data) noexcept;
    [[nodiscard]] RegResult TrySetMultiStringValue(const std::wstring& valueName,
                                                   const std::vector<std::wstring>& data);
    [[nodiscard]] RegResult TrySetBinaryValue(const std::wstring& valueName,
                                              const std::vector<BYTE>& data) noexcept;
    [[nodiscard]] RegResult TrySetBinaryValue(const std::wstring& valueName,
                                              const void* data, DWORD dataSize) noexcept;


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


    [[nodiscard]] RegResult TryGetDwordValue(const std::wstring& valueName, DWORD& result) const noexcept;
    [[nodiscard]] RegResult TryGetQwordValue(const std::wstring& valueName, ULONGLONG& result) const noexcept;
    [[nodiscard]] RegResult TryGetStringValue(const std::wstring& valueName, std::wstring& result) const;

    [[nodiscard]] RegResult TryGetExpandStringValue(
        const std::wstring& valueName,
        std::wstring& result,
        ExpandStringOption expandOption = ExpandStringOption::DontExpand
    ) const;

    [[nodiscard]] RegResult TryGetMultiStringValue(const std::wstring& valueName,
                                                   std::vector<std::wstring>& result) const;
    [[nodiscard]] RegResult TryGetBinaryValue(const std::wstring& valueName,
                                              std::vector<BYTE>& result) const;


    //
    // Query Operations
    //

    void QueryInfoKey(DWORD& subKeys, DWORD &values, FILETIME& lastWriteTime) const;

    // Return the DWORD type ID for the input registry value
    [[nodiscard]] DWORD QueryValueType(const std::wstring& valueName) const;

    // Enumerate the subkeys of the registry key, using RegEnumKeyEx
    [[nodiscard]] std::vector<std::wstring> EnumSubKeys() const;

    // Enumerate the values under the registry key, using RegEnumValue.
    // Returns a vector of pairs: In each pair, the wstring is the value name,
    // the DWORD is the value type.
    [[nodiscard]] std::vector<std::pair<std::wstring, DWORD>> EnumValues() const;

    // Enumerate the subkeys of the registry key, using RegEnumKeyEx
    [[nodiscard]] RegResult TryEnumSubKeys(std::vector<std::wstring>& subKeys) const;

    // Enumerate the values under the registry key, using RegEnumValue.
    // Returns a vector of pairs: In each pair, the wstring is the value name,
    // the DWORD is the value type.
    [[nodiscard]] RegResult TryEnumValues(std::vector<std::pair<std::wstring, DWORD>>& values) const;


    [[nodiscard]] RegResult TryQueryInfoKey(DWORD& subKeys,
                              DWORD &values,
                              FILETIME& lastWriteTime) const noexcept;

    // Return the DWORD type ID for the input registry value
    [[nodiscard]] RegResult TryQueryValueType(const std::wstring& valueName, DWORD& typeId) const noexcept;


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
    [[nodiscard]] bool QueryReflectionKey() const;
    void ConnectRegistry(const std::wstring& machineName, HKEY hKeyPredefined);

    [[nodiscard]] RegResult TryDeleteValue(const std::wstring& valueName) noexcept;
    [[nodiscard]] RegResult TryDeleteKey(const std::wstring& subKey, REGSAM desiredAccess) noexcept;
    [[nodiscard]] RegResult TryDeleteTree(const std::wstring& subKey) noexcept;
    [[nodiscard]] RegResult TryCopyTree(const std::wstring& sourceSubKey, const RegKey& destKey) noexcept;
    [[nodiscard]] RegResult TryFlushKey() noexcept;
    [[nodiscard]] RegResult TryLoadKey(const std::wstring& subKey, const std::wstring& filename) noexcept;
    [[nodiscard]] RegResult TrySaveKey(const std::wstring& filename,
                                       SECURITY_ATTRIBUTES* securityAttributes) const noexcept;
    [[nodiscard]] RegResult TryEnableReflectionKey() noexcept;
    [[nodiscard]] RegResult TryDisableReflectionKey() noexcept;
    [[nodiscard]] RegResult TryQueryReflectionKey(bool& isReflected) const noexcept;
    [[nodiscard]] RegResult TryConnectRegistry(const std::wstring& machineName, HKEY hKeyPredefined) noexcept;


    // Return a string representation of Windows registry types
    [[nodiscard]] static std::wstring RegTypeToString(DWORD regType);

    //
    // Relational comparison operators are overloaded as non-members
    // ==, !=, <, <=, >, >=
    //


private:
    // The wrapped registry key handle
    HKEY m_hKey = nullptr;
};


//------------------------------------------------------------------------------
// A simple wrapper around LONG return codes used by the Windows Registry API.
//------------------------------------------------------------------------------
class WINREG_API RegResult
{
public:

    // Initialize to success code (ERROR_SUCCESS)
    RegResult() noexcept = default;

    // Conversion constructor, *not* marked "explicit" on purpose,
    // allows easy and convenient conversion from Win32 API return code type
    // to this C++ wrapper.
    RegResult(LONG result) noexcept;

    // Is the wrapped code a success code?
    [[nodiscard]] bool IsOk() const noexcept;

    // Is the wrapped error code a failure code?
    [[nodiscard]] bool Failed() const noexcept;

    // Is the wrapped code a success code?
    explicit operator bool() const noexcept;

    // Get the wrapped Win32 code
    [[nodiscard]] LONG Code() const noexcept;

    // Return the system error message associated to the current error code
    [[nodiscard]] std::wstring ErrorMessage() const;

    // Return the system error message associated to the current error code,
    // using the given input language identifier
    [[nodiscard]] std::wstring ErrorMessage(DWORD languageId) const;

private:
    // Error code returned by Windows Registry C API;
    // default initialized to success code.
    LONG m_result = ERROR_SUCCESS;
};


//------------------------------------------------------------------------------
// An exception representing an error with the registry operations
//------------------------------------------------------------------------------
class WINREG_API RegException
    : public std::system_error
{
public:
    RegException(LONG errorCode, const char* message);
    RegException(LONG errorCode, const std::string& message);
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
//                          RegKey Inline Methods
//------------------------------------------------------------------------------

inline RegKey::RegKey(const HKEY hKey) noexcept
    : m_hKey(hKey)
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
    : m_hKey(other.m_hKey)
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
    const DWORD kDefaultOptions = REG_OPTION_NON_VOLATILE;

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
    LONG retCode = ::RegCreateKeyExW(
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
        throw RegException(retCode, "RegCreateKeyEx failed.");
    }

    // Safely close any previously opened key
    Close();

    // Take ownership of the newly created key
    m_hKey = hKey;
}


inline RegResult RegKey::TryCreate(
    const HKEY                  hKeyParent,
    const std::wstring&         subKey,
    const REGSAM                desiredAccess
) noexcept
{
    const DWORD kDefaultOptions = REG_OPTION_NON_VOLATILE;

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
    RegResult retCode = ::RegCreateKeyExW(
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


inline void RegKey::Open(
    const HKEY              hKeyParent,
    const std::wstring&     subKey,
    const REGSAM            desiredAccess
)
{
    HKEY hKey = nullptr;
    LONG retCode = ::RegOpenKeyExW(
        hKeyParent,
        subKey.c_str(),
        REG_NONE,           // default options
        desiredAccess,
        &hKey
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "RegOpenKeyExW failed.");
    }

    // Safely close any previously opened key
    Close();

    // Take ownership of the newly created key
    m_hKey = hKey;
}


inline RegResult RegKey::TryOpen(
    const HKEY              hKeyParent,
    const std::wstring&     subKey,
    const REGSAM            desiredAccess
) noexcept
{
    HKEY hKey = nullptr;
    RegResult retCode = ::RegOpenKeyExW(
        hKeyParent,
        subKey.c_str(),
        REG_NONE,           // default options
        desiredAccess,
        &hKey
    );
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

    LONG retCode = ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_DWORD,
        reinterpret_cast<const BYTE*>(&data),
        sizeof(data)
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot write DWORD value: RegSetValueExW failed.");
    }
}


inline void RegKey::SetQwordValue(const std::wstring& valueName, const ULONGLONG& data)
{
    _ASSERTE(IsValid());

    LONG retCode = ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_QWORD,
        reinterpret_cast<const BYTE*>(&data),
        sizeof(data)
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot write QWORD value: RegSetValueExW failed.");
    }
}


inline void RegKey::SetStringValue(const std::wstring& valueName, const std::wstring& data)
{
    _ASSERTE(IsValid());

    // String size including the terminating NUL, in bytes
    const DWORD dataSize = static_cast<DWORD>((data.length() + 1) * sizeof(wchar_t));

    LONG retCode = ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_SZ,
        reinterpret_cast<const BYTE*>(data.c_str()),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot write string value: RegSetValueExW failed.");
    }
}


inline void RegKey::SetExpandStringValue(const std::wstring& valueName, const std::wstring& data)
{
    _ASSERTE(IsValid());

    // String size including the terminating NUL, in bytes
    const DWORD dataSize = static_cast<DWORD>((data.length() + 1) * sizeof(wchar_t));

    LONG retCode = ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_EXPAND_SZ,
        reinterpret_cast<const BYTE*>(data.c_str()),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot write expand string value: RegSetValueExW failed.");
    }
}




inline void RegKey::SetBinaryValue(const std::wstring& valueName, const std::vector<BYTE>& data)
{
    _ASSERTE(IsValid());

    // Total data size, in bytes
    const DWORD dataSize = static_cast<DWORD>(data.size());

    LONG retCode = ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_BINARY,
        &data[0],
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot write binary data value: RegSetValueExW failed.");
    }
}


inline void RegKey::SetBinaryValue(
    const std::wstring& valueName,
    const void* const data,
    const DWORD dataSize
)
{
    _ASSERTE(IsValid());

    LONG retCode = ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_BINARY,
        static_cast<const BYTE*>(data),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot write binary data value: RegSetValueExW failed.");
    }
}


inline RegResult RegKey::TrySetDwordValue(const std::wstring& valueName, const DWORD data) noexcept
{
    _ASSERTE(IsValid());

    return ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_DWORD,
        reinterpret_cast<const BYTE*>(&data),
        sizeof(data)
    );
}


inline RegResult RegKey::TrySetQwordValue(
    const std::wstring& valueName,
    const ULONGLONG& data
) noexcept
{
    _ASSERTE(IsValid());

    return ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_QWORD,
        reinterpret_cast<const BYTE*>(&data),
        sizeof(data)
    );
}


inline RegResult RegKey::TrySetStringValue(
    const std::wstring& valueName,
    const std::wstring& data
) noexcept
{
    _ASSERTE(IsValid());

    // String size including the terminating NUL, in bytes
    const DWORD dataSize = static_cast<DWORD>((data.length() + 1) * sizeof(wchar_t));

    return ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_SZ,
        reinterpret_cast<const BYTE*>(data.c_str()),
        dataSize
    );
}


inline RegResult RegKey::TrySetExpandStringValue(
    const std::wstring& valueName,
    const std::wstring& data
) noexcept
{
    _ASSERTE(IsValid());

    // String size including the terminating NUL, in bytes
    const DWORD dataSize = static_cast<DWORD>((data.length() + 1) * sizeof(wchar_t));

    return ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_EXPAND_SZ,
        reinterpret_cast<const BYTE*>(data.c_str()),
        dataSize
    );
}


inline RegResult RegKey::TrySetBinaryValue(
    const std::wstring& valueName,
    const std::vector<BYTE>& data
) noexcept
{
    _ASSERTE(IsValid());

    // Total data size, in bytes
    const DWORD dataSize = static_cast<DWORD>(data.size());

    return ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_BINARY,
        &data[0],
        dataSize
    );
}


inline RegResult RegKey::TrySetBinaryValue(
    const std::wstring& valueName,
    const void* const data,
    const DWORD dataSize
) noexcept
{
    _ASSERTE(IsValid());

    return ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_BINARY,
        static_cast<const BYTE*>(data),
        dataSize
    );
}


inline DWORD RegKey::GetDwordValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    DWORD data = 0;                  // to be read from the registry
    DWORD dataSize = sizeof(data);   // size of data, in bytes

    const DWORD flags = RRF_RT_REG_DWORD;
    LONG retCode = ::RegGetValueW(
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
        throw RegException(retCode, "Cannot get DWORD value: RegGetValueW failed.");
    }

    return data;
}


inline ULONGLONG RegKey::GetQwordValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    ULONGLONG data = 0;              // to be read from the registry
    DWORD dataSize = sizeof(data);   // size of data, in bytes

    const DWORD flags = RRF_RT_REG_QWORD;
    LONG retCode = ::RegGetValueW(
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
        throw RegException(retCode, "Cannot get QWORD value: RegGetValueW failed.");
    }

    return data;
}


inline RegResult RegKey::TryGetDwordValue(
    const std::wstring& valueName,
    DWORD& result
) const noexcept
{
    _ASSERTE(IsValid());

    result = 0;

    DWORD data = 0;                   // to be read from the registry
    DWORD dataSize = sizeof(data);    // size of data, in bytes

    const DWORD flags = RRF_RT_REG_DWORD;
    RegResult retCode = ::RegGetValueW(
        m_hKey,
        nullptr, // no subkey
        valueName.c_str(),
        flags,
        nullptr, // type not required
        &data,
        &dataSize
    );
    if (retCode.Failed())
    {
        result = 0;
        return retCode;
    }

    result = data;
    _ASSERTE(retCode.IsOk());
    return retCode;
}


inline RegResult RegKey::TryGetQwordValue(
    const std::wstring& valueName,
    ULONGLONG& result
) const noexcept
{
    _ASSERTE(IsValid());

    result = 0;

    ULONGLONG data = 0;                 // to be read from the registry
    DWORD dataSize = sizeof(data);      // size of data, in bytes

    const DWORD flags = RRF_RT_REG_QWORD;
    RegResult retCode = ::RegGetValueW(
        m_hKey,
        nullptr, // no subkey
        valueName.c_str(),
        flags,
        nullptr, // type not required
        &data,
        &dataSize
    );
    if (retCode.Failed())
    {
        result = 0;
        return retCode;
    }

    result = data;
    _ASSERTE(retCode.IsOk());
    return retCode;
}


inline DWORD RegKey::QueryValueType(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    DWORD typeId = 0;     // will be returned by RegQueryValueEx

    LONG retCode = ::RegQueryValueExW(
        m_hKey,
        valueName.c_str(),
        nullptr,    // reserved
        &typeId,
        nullptr,    // not interested
        nullptr     // not interested
    );

    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot get the value type: RegQueryValueExW failed.");
    }

    return typeId;
}


inline void RegKey::QueryInfoKey(DWORD& subKeys, DWORD &values, FILETIME& lastWriteTime) const
{
    _ASSERTE(IsValid());

    subKeys = 0;
    values = 0;
    lastWriteTime.dwLowDateTime = lastWriteTime.dwHighDateTime = 0;

    LONG retCode = ::RegQueryInfoKeyW(
        m_hKey,
        nullptr,
        nullptr,
        nullptr,
        &subKeys,
        nullptr,
        nullptr,
        &values,
        nullptr,
        nullptr,
        nullptr,
        &lastWriteTime
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "RegQueryInfoKeyW failed.");
    }
}


inline RegResult RegKey::TryQueryValueType(
    const std::wstring& valueName,
    DWORD& typeId
) const noexcept
{
    _ASSERTE(IsValid());

    typeId = 0;

    return ::RegQueryValueExW(
        m_hKey,
        valueName.c_str(),
        nullptr,    // reserved
        &typeId,
        nullptr,    // not interested
        nullptr     // not interested
    );
}


inline RegResult RegKey::TryQueryInfoKey(
    DWORD& subKeys,
    DWORD &values,
    FILETIME& lastWriteTime
) const noexcept
{
    _ASSERTE(IsValid());

    subKeys = 0;
    values = 0;
    lastWriteTime.dwLowDateTime = lastWriteTime.dwHighDateTime = 0;

    return ::RegQueryInfoKeyW(
        m_hKey,
        nullptr,
        nullptr,
        nullptr,
        &subKeys,
        nullptr,
        nullptr,
        &values,
        nullptr,
        nullptr,
        nullptr,
        &lastWriteTime
    );
}


inline void RegKey::DeleteValue(const std::wstring& valueName)
{
    _ASSERTE(IsValid());

    LONG retCode = ::RegDeleteValueW(m_hKey, valueName.c_str());
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "RegDeleteValueW failed.");
    }
}


inline void RegKey::DeleteKey(const std::wstring& subKey, const REGSAM desiredAccess)
{
    _ASSERTE(IsValid());

    LONG retCode = ::RegDeleteKeyExW(m_hKey, subKey.c_str(), desiredAccess, 0);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "RegDeleteKeyExW failed.");
    }
}


inline void RegKey::DeleteTree(const std::wstring& subKey)
{
    _ASSERTE(IsValid());

    LONG retCode = ::RegDeleteTreeW(m_hKey, subKey.c_str());
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "RegDeleteTreeW failed.");
    }
}


inline void RegKey::CopyTree(const std::wstring& sourceSubKey, const RegKey& destKey)
{
    _ASSERTE(IsValid());

    LONG retCode = ::RegCopyTreeW(m_hKey, sourceSubKey.c_str(), destKey.Get());
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "RegCopyTreeW failed.");
    }
}


inline void RegKey::FlushKey()
{
    _ASSERTE(IsValid());

    LONG retCode = ::RegFlushKey(m_hKey);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "RegFlushKey failed.");
    }
}


inline void RegKey::LoadKey(const std::wstring& subKey, const std::wstring& filename)
{
    Close();

    LONG retCode = ::RegLoadKeyW(m_hKey, subKey.c_str(), filename.c_str());
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "RegLoadKeyW failed.");
    }
}


inline void RegKey::SaveKey(
    const std::wstring& filename,
    SECURITY_ATTRIBUTES* const securityAttributes
) const
{
    _ASSERTE(IsValid());

    LONG retCode = ::RegSaveKeyW(m_hKey, filename.c_str(), securityAttributes);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "RegSaveKeyW failed.");
    }
}


inline void RegKey::EnableReflectionKey()
{
    LONG retCode = ::RegEnableReflectionKey(m_hKey);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "RegEnableReflectionKey failed.");
    }
}


inline void RegKey::DisableReflectionKey()
{
    LONG retCode = ::RegDisableReflectionKey(m_hKey);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "RegDisableReflectionKey failed.");
    }
}


inline bool RegKey::QueryReflectionKey() const
{
    BOOL isReflectionDisabled = FALSE;
    LONG retCode = ::RegQueryReflectionKey(m_hKey, &isReflectionDisabled);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "RegQueryReflectionKey failed.");
    }

    return (isReflectionDisabled ? true : false);
}


inline void RegKey::ConnectRegistry(const std::wstring& machineName, const HKEY hKeyPredefined)
{
    // Safely close any previously opened key
    Close();

    HKEY hKeyResult = nullptr;
    LONG retCode = ::RegConnectRegistryW(machineName.c_str(), hKeyPredefined, &hKeyResult);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "RegConnectRegistryW failed.");
    }

    // Take ownership of the result key
    m_hKey = hKeyResult;
}


inline RegResult RegKey::TryDeleteValue(const std::wstring& valueName) noexcept
{
    _ASSERTE(IsValid());

    return ::RegDeleteValueW(m_hKey, valueName.c_str());
}


inline RegResult RegKey::TryDeleteKey(const std::wstring& subKey, REGSAM desiredAccess) noexcept
{
    _ASSERTE(IsValid());

    return ::RegDeleteKeyExW(m_hKey, subKey.c_str(), desiredAccess, 0);
}


inline RegResult RegKey::TryDeleteTree(const std::wstring& subKey) noexcept
{
    _ASSERTE(IsValid());

    return ::RegDeleteTreeW(m_hKey, subKey.c_str());
}


inline RegResult RegKey::TryCopyTree(
    const std::wstring& sourceSubKey,
    const RegKey& destKey
) noexcept
{
    _ASSERTE(IsValid());

    return ::RegCopyTreeW(m_hKey, sourceSubKey.c_str(), destKey.Get());
}


inline RegResult RegKey::TryFlushKey() noexcept
{
    _ASSERTE(IsValid());

    return ::RegFlushKey(m_hKey);
}


inline RegResult RegKey::TryLoadKey(
    const std::wstring& subKey,
    const std::wstring& filename
) noexcept
{
    Close();

    return ::RegLoadKeyW(m_hKey, subKey.c_str(), filename.c_str());
}


inline RegResult RegKey::TrySaveKey(
    const std::wstring& filename,
    SECURITY_ATTRIBUTES* const securityAttributes
) const noexcept
{
    _ASSERTE(IsValid());

    return ::RegSaveKeyW(m_hKey, filename.c_str(), securityAttributes);
}


inline RegResult RegKey::TryEnableReflectionKey() noexcept
{
    return ::RegEnableReflectionKey(m_hKey);
}


inline RegResult RegKey::TryDisableReflectionKey() noexcept
{
    return ::RegDisableReflectionKey(m_hKey);
}


inline RegResult RegKey::TryQueryReflectionKey(bool& reflectionDisabled) const noexcept
{
    BOOL isReflectionDisabled = FALSE;
    RegResult retCode = ::RegQueryReflectionKey(m_hKey, &isReflectionDisabled);
    if (retCode.Failed())
    {
        return retCode;
    }

    reflectionDisabled = isReflectionDisabled ? true : false;

    _ASSERTE(retCode.IsOk());
    return retCode;
}


inline RegResult RegKey::TryConnectRegistry(
    const std::wstring& machineName,
    const HKEY hKeyPredefined
) noexcept
{
    // Safely close any previously opened key
    Close();

    HKEY hKeyResult = nullptr;
    RegResult retCode = ::RegConnectRegistryW(machineName.c_str(), hKeyPredefined, &hKeyResult);
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
//                          RegResult Inline Methods
//------------------------------------------------------------------------------

inline RegResult::RegResult(const LONG result) noexcept
    : m_result(result)
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


inline LONG RegResult::Code() const noexcept
{
    return m_result;
}


inline std::wstring RegResult::ErrorMessage() const
{
    return ErrorMessage(MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT));
}


//------------------------------------------------------------------------------
//                          RegException Inline Methods
//------------------------------------------------------------------------------

inline RegException::RegException(const LONG errorCode, const char* const message)
    : std::system_error(errorCode, std::system_category(), message)
{}


inline RegException::RegException(const LONG errorCode, const std::string& message)
    : std::system_error(errorCode, std::system_category(), message)
{}


} // namespace winreg


#endif // GIOVANNI_DICANIO_WINREG_HPP_INCLUDED

#ifndef GIOVANNI_DICANIO_WINREG_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_HPP_INCLUDED


////////////////////////////////////////////////////////////////////////////////
//
//      *** Modern C++ Wrappers Around Windows Registry C API ***
//
//               Copyright (C) by Giovanni Dicanio
//
// ---------------------------------------------------------------------------
// FILE: WinReg.hpp
// DESC: Library's public header
// ---------------------------------------------------------------------------
//
// First version: 2017, January 22nd
// Last update:   2024, October 24
//                [forked from v6.3.1 header-only version]
//
// E-mail: <first name>.<last name> AT REMOVE_THIS gmail.com
//
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

#include <string>           // std::wstring
#include <system_error>     // std::system_error
#include <utility>          // std::swap, std::pair, std::move
#include <variant>          // std::variant
#include <vector>           // std::vector


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
                                              const std::wstring& data);
    // Note: The TrySetStringValue method CANNOT be marked noexcept,
    // because internally the method *may* throw an exception if the input string is too big
    // (size_t value overflowing a DWORD).

    [[nodiscard]] RegResult TrySetExpandStringValue(const std::wstring& valueName,
                                                    const std::wstring& data);
    // Note: The TrySetExpandStringValue method CANNOT be marked noexcept,
    // because internally the method *may* throw an exception if the input string is too big
    // (size_t value overflowing a DWORD).

    [[nodiscard]] RegResult TrySetMultiStringValue(const std::wstring& valueName,
                                                   const std::vector<std::wstring>& data);
    // Note: The TrySetMultiStringValue method CANNOT be marked noexcept,
    // because internally the method *dynamically allocates memory* for creating the multi-string
    // that will be stored in the Registry.

    [[nodiscard]] RegResult TrySetBinaryValue(const std::wstring& valueName,
                                              const std::vector<BYTE>& data);
    // Note: This overload of the TrySetBinaryValue method CANNOT be marked noexcept,
    // because internally the method *may* throw an exception if the input vector is too large
    // (vector::size size_t value overflowing a DWORD).


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

    // Check if the current key contains the specified value
    [[nodiscard]] bool ContainsValue(const std::wstring& valueName) const;

    // Check if the current key contains the specified sub-key
    [[nodiscard]] bool ContainsSubKey(const std::wstring& subKey) const;


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

    // Check if the current key contains the specified value
    [[nodiscard]] RegExpected<bool> TryContainsValue(const std::wstring& valueName) const;

    // Check if the current key contains the specified sub-key
    [[nodiscard]] RegExpected<bool> TryContainsSubKey(const std::wstring& subKey) const;


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
    // Throws an exception if the object is in valid state.
    [[nodiscard]] RegResult GetError() const;


private:
    // Stores a value of type T on success,
    // or RegResult on error
    std::variant<RegResult, T> m_var;
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



//------------------------------------------------------------------------------
//                          RegExpected Inline Methods
//------------------------------------------------------------------------------

template <typename T>
inline RegExpected<T>::RegExpected(const RegResult& errorCode) noexcept
    : m_var{ errorCode }
{}


template <typename T>
inline RegExpected<T>::RegExpected(const T& value)
    : m_var{ value }
{}


template <typename T>
inline RegExpected<T>::RegExpected(T&& value)
    : m_var{ std::move(value) }
{}


template <typename T>
inline RegExpected<T>::operator bool() const noexcept
{
    return IsValid();
}


template <typename T>
inline bool RegExpected<T>::IsValid() const noexcept
{
    return std::holds_alternative<T>(m_var);
}


template <typename T>
inline const T& RegExpected<T>::GetValue() const
{
    // Check that the object stores a valid value
    _ASSERTE(IsValid());

    // If the object is in a valid state, the variant stores an instance of T
    return std::get<T>(m_var);
}


template <typename T>
inline RegResult RegExpected<T>::GetError() const
{
    // Check that the object is in an invalid state
    _ASSERTE(!IsValid());

    // If the object is in an invalid state, the variant stores a RegResult
    // that represents an error code from the Windows Registry API
    return std::get<RegResult>(m_var);
}


} // namespace winreg


#endif // GIOVANNI_DICANIO_WINREG_HPP_INCLUDED

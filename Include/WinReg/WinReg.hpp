////////////////////////////////////////////////////////////////////////////////
// FILE: WinReg.hpp
// DESC: Public header of the WinReg library.
// AUTHOR: Giovanni Dicanio
////////////////////////////////////////////////////////////////////////////////

#ifndef GIOVANNI_DICANIO_WINREG_WINREG_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_WINREG_HPP_INCLUDED


////////////////////////////////////////////////////////////////////////////////
//
//      *** Modern C++ Wrappers Around Windows Registry C API ***
//
//               Copyright (C) by Giovanni Dicanio
//
// First version: 2017, January 22nd
// Last update:   2024, November 13th
//
// E-mail: <first name>.<last name> AT REMOVE_THIS gmail.com
//
// Registry key handles are safely and conveniently wrapped
// in the RegKeyT resource manager C++ template class.
//
// Many methods are available in two forms:
//
// - One form that signals errors throwing exceptions
//   of class RegException (e.g. RegKeyT::Open)
//
// - Another form that returns RegResult objects (e.g. RegKeyT::TryOpen)
//
// In addition, there are also some methods named like TryGet...Value
// (e.g. TryGetDwordValue), that _try_ to perform the given query,
// and return a RegExpected object. On success, that object contains
// the value read from the registry. On failure, the returned RegExpected object
// contains a RegResult storing the return code from the Windows Registry API call.
//
// Unicode UTF-16 strings are represented using the std::wstring class;
// ATL's CString is not used, to avoid dependencies from ATL or MFC.
// Unicode UTF-8 strings are represented using std::string.
//
// Compiler: Visual Studio 2019
// C++ Language Standard: C++17 (/std:c++17)
// Code compiles cleanly at warning level 4 (/W4) on both 32-bit and 64-bit builds.
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
#include <string>           // std::string, std::wstring
#include <type_traits>      // std::is_same_v
#include <utility>          // std::swap, std::pair, std::move
#include <vector>           // std::vector

#include "WinReg/Details.hpp"
#include "WinReg/RegException.hpp"
#include "WinReg/RegExpected.hpp"
#include "WinReg/RegResult.hpp"
#include "WinReg/Utf8Conv.hpp"


namespace winreg
{


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
// RegKeyT is actually a class template that is specialized to use
// UTF-8 std::strings or UTF-16 std::wstrings at its public interface.
// Client code should use RegKey or RegKeyUtf8 instead of this RegKeyT template.
//
//------------------------------------------------------------------------------
template <typename StringType, typename StringTraits>
class RegKeyT
{
public:

    typedef StringType StringType;
    typedef StringTraits StringTraits;


    //
    // Construction/Destruction
    //

    // Initialize as an empty key handle
    RegKeyT() noexcept = default;

    // Take ownership of the input key handle
    explicit RegKeyT(HKEY hKey) noexcept;

    // Open the given registry key if it exists, else create a new key.
    // Uses default KEY_READ|KEY_WRITE|KEY_WOW64_64KEY access.
    // For finer grained control, call the Create() method overloads.
    // Throw RegException on failure.
    RegKeyT(HKEY hKeyParent, const StringType& subKey);

    // Open the given registry key if it exists, else create a new key.
    // Allow the caller to specify the desired access to the key
    // (e.g. KEY_READ|KEY_WOW64_64KEY for read-only access).
    // For finer grained control, call the Create() method overloads.
    // Throw RegException on failure.
    RegKeyT(HKEY hKeyParent, const StringType& subKey, REGSAM desiredAccess);


    // Take ownership of the input key handle.
    // The input key handle wrapper is reset to an empty state.
    RegKeyT(RegKeyT&& other) noexcept;

    // Move-assign from the input key handle.
    // Properly check against self-move-assign (which is safe and does nothing).
    RegKeyT& operator=(RegKeyT&& other) noexcept;

    // Ban copy
    RegKeyT(const RegKeyT&) = delete;
    RegKeyT& operator=(const RegKeyT&) = delete;

    // Safely close the wrapped key handle (if any)
    ~RegKeyT() noexcept;


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
    void SwapWith(RegKeyT& other) noexcept;


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
        const StringType& subKey,
        REGSAM desiredAccess = KEY_READ | KEY_WRITE | KEY_WOW64_64KEY
    );

    // Wrapper around RegCreateKeyEx
    void Create(
        HKEY hKeyParent,
        const StringType& subKey,
        REGSAM desiredAccess,
        DWORD options,
        SECURITY_ATTRIBUTES* securityAttributes,
        DWORD* disposition
    );

    // Wrapper around RegOpenKeyEx
    void Open(
        HKEY hKeyParent,
        const StringType& subKey,
        REGSAM desiredAccess = KEY_READ | KEY_WRITE | KEY_WOW64_64KEY
    );

    // Wrapper around RegCreateKeyEx, that allows you to specify desired access
    [[nodiscard]] RegResult TryCreate(
        HKEY hKeyParent,
        const StringType& subKey,
        REGSAM desiredAccess = KEY_READ | KEY_WRITE | KEY_WOW64_64KEY
    );

    // Wrapper around RegCreateKeyEx
    [[nodiscard]] RegResult TryCreate(
        HKEY hKeyParent,
        const StringType& subKey,
        REGSAM desiredAccess,
        DWORD options,
        SECURITY_ATTRIBUTES* securityAttributes,
        DWORD* disposition
    );

    // Wrapper around RegOpenKeyEx
    [[nodiscard]] RegResult TryOpen(
        HKEY hKeyParent,
        const StringType& subKey,
        REGSAM desiredAccess = KEY_READ | KEY_WRITE | KEY_WOW64_64KEY
    );


    //
    // Registry Value Setters
    //

    void SetDwordValue(const StringType& valueName, DWORD data);
    void SetQwordValue(const StringType& valueName, const ULONGLONG& data);
    void SetStringValue(const StringType& valueName, const StringType& data);
    void SetExpandStringValue(const StringType& valueName, const StringType& data);
    void SetMultiStringValue(const StringType& valueName, const std::vector<StringType>& data);
    void SetBinaryValue(const StringType& valueName, const std::vector<BYTE>& data);
    void SetBinaryValue(const StringType& valueName, const void* data, DWORD dataSize);


    //
    // Registry Value Setters Returning RegResult
    // (instead of throwing RegException on error)
    //

    [[nodiscard]] RegResult TrySetDwordValue(const StringType& valueName, DWORD data);

    [[nodiscard]] RegResult TrySetQwordValue(const StringType& valueName,
                                             const ULONGLONG& data);

    [[nodiscard]] RegResult TrySetStringValue(const StringType& valueName,
                                              const StringType& data);

    [[nodiscard]] RegResult TrySetExpandStringValue(const StringType& valueName,
                                                    const StringType& data);

    [[nodiscard]] RegResult TrySetMultiStringValue(const StringType& valueName,
                                                   const std::vector<StringType>& data);

    [[nodiscard]] RegResult TrySetBinaryValue(const StringType& valueName,
                                              const std::vector<BYTE>& data);

    [[nodiscard]] RegResult TrySetBinaryValue(const StringType& valueName,
                                              const void* data,
                                              DWORD dataSize);


    //
    // Registry Value Getters
    //

    [[nodiscard]] DWORD GetDwordValue(const StringType& valueName) const;
    [[nodiscard]] ULONGLONG GetQwordValue(const StringType& valueName) const;
    [[nodiscard]] StringType GetStringValue(const StringType& valueName) const;

    enum class ExpandStringOption
    {
        DontExpand,
        Expand
    };

    [[nodiscard]] StringType GetExpandStringValue(
        const StringType& valueName,
        ExpandStringOption expandOption = ExpandStringOption::DontExpand
    ) const;

    [[nodiscard]] std::vector<StringType> GetMultiStringValue(const StringType& valueName) const;
    [[nodiscard]] std::vector<BYTE> GetBinaryValue(const StringType& valueName) const;


    //
    // Registry Value Getters Returning RegExpected<T>
    // (instead of throwing RegException on error)
    //

    [[nodiscard]] RegExpected<DWORD> TryGetDwordValue(const StringType& valueName) const;
    [[nodiscard]] RegExpected<ULONGLONG> TryGetQwordValue(const StringType& valueName) const;
    [[nodiscard]] RegExpected<StringType> TryGetStringValue(const StringType& valueName) const;

    [[nodiscard]] RegExpected<StringType> TryGetExpandStringValue(
        const StringType& valueName,
        ExpandStringOption expandOption = ExpandStringOption::DontExpand
    ) const;

    [[nodiscard]] RegExpected<std::vector<StringType>>
            TryGetMultiStringValue(const StringType& valueName) const;

    [[nodiscard]] RegExpected<std::vector<BYTE>>
            TryGetBinaryValue(const StringType& valueName) const;


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
    [[nodiscard]] DWORD QueryValueType(const StringType& valueName) const;


    enum class KeyReflection
    {
        ReflectionEnabled,
        ReflectionDisabled
    };

    // Determines whether reflection has been disabled or enabled for the specified key
    [[nodiscard]] KeyReflection QueryReflectionKey() const;

    // Enumerate the subkeys of the registry key, using RegEnumKeyEx
    [[nodiscard]] std::vector<StringType> EnumSubKeys() const;

    // Enumerate the values under the registry key, using RegEnumValue.
    // Returns a vector of pairs: In each pair, the wstring is the value name,
    // the DWORD is the value type.
    [[nodiscard]] std::vector<std::pair<StringType, DWORD>> EnumValues() const;

    // Check if the current key contains the specified value
    [[nodiscard]] bool ContainsValue(const StringType& valueName) const;

    // Check if the current key contains the specified sub-key
    [[nodiscard]] bool ContainsSubKey(const StringType& subKey) const;


    //
    // Query Operations Returning RegExpected
    // (instead of throwing RegException on error)
    //

    // Retrieve information about the registry key
    [[nodiscard]] RegExpected<InfoKey> TryQueryInfoKey() const;

    // Return the DWORD type ID for the input registry value
    [[nodiscard]] RegExpected<DWORD> TryQueryValueType(const StringType& valueName) const;


    // Determines whether reflection has been disabled or enabled for the specified key
    [[nodiscard]] RegExpected<KeyReflection> TryQueryReflectionKey() const;

    // Enumerate the subkeys of the registry key, using RegEnumKeyEx
    [[nodiscard]] RegExpected<std::vector<StringType>> TryEnumSubKeys() const;

    // Enumerate the values under the registry key, using RegEnumValue.
    // Returns a vector of pairs: In each pair, the wstring is the value name,
    // the DWORD is the value type.
    [[nodiscard]] RegExpected<std::vector<std::pair<StringType, DWORD>>> TryEnumValues() const;

    // Check if the current key contains the specified value
    [[nodiscard]] RegExpected<bool> TryContainsValue(const StringType& valueName) const;

    // Check if the current key contains the specified sub-key
    [[nodiscard]] RegExpected<bool> TryContainsSubKey(const StringType& subKey) const;


    //
    // Misc Registry API Wrappers
    //

    void DeleteValue(const StringType& valueName);
    void DeleteKey(const StringType& subKey, REGSAM desiredAccess);
    void DeleteTree(const StringType& subKey);
    void CopyTree(const StringType& sourceSubKey, const RegKeyT& destKey);
    void FlushKey();
    void LoadKey(const StringType& subKey, const StringType& filename);
    void SaveKey(const StringType& filename, SECURITY_ATTRIBUTES* securityAttributes) const;
    void EnableReflectionKey();
    void DisableReflectionKey();
    void ConnectRegistry(const StringType& machineName, HKEY hKeyPredefined);


    //
    // Misc Registry API Wrappers Returning RegResult Status
    // (instead of throwing RegException on error)
    //

    [[nodiscard]] RegResult TryDeleteValue(const StringType& valueName);
    [[nodiscard]] RegResult TryDeleteKey(const StringType& subKey, REGSAM desiredAccess);
    [[nodiscard]] RegResult TryDeleteTree(const StringType& subKey);

    [[nodiscard]] RegResult TryCopyTree(const StringType& sourceSubKey,
                                        const RegKeyT& destKey);

    [[nodiscard]] RegResult TryFlushKey();

    [[nodiscard]] RegResult TryLoadKey(const StringType& subKey,
                                       const StringType& filename);

    [[nodiscard]] RegResult TrySaveKey(const StringType& filename,
                                       SECURITY_ATTRIBUTES* securityAttributes) const;

    [[nodiscard]] RegResult TryEnableReflectionKey();
    [[nodiscard]] RegResult TryDisableReflectionKey();

    [[nodiscard]] RegResult TryConnectRegistry(const StringType& machineName,
                                               HKEY hKeyPredefined);


    // Return a string representation of Windows registry types
    [[nodiscard]] static StringType RegTypeToString(DWORD regType);


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

// RegKey has UTF-16 strings at his public interface.
// (I kept RegKey for backward compatibility with previous UTF-16-only versions of the library.)
using RegKey = RegKeyT<std::wstring, details::StringTraitUtf16>;

// RegKeyUtf8 uses UTF-8 strings at the public interface
// (and converts to UTF-16 internally for Windows API calls).
using RegKeyUtf8 = RegKeyT<std::string, details::StringTraitUtf8>;



//------------------------------------------------------------------------------
//          Overloads of relational comparison operators for RegKeyT
//------------------------------------------------------------------------------

template <typename StringType, typename StringTraits>
inline bool operator==(const RegKeyT<StringType, StringTraits>& a,
                       const RegKeyT<StringType, StringTraits>& b) noexcept
{
    return a.Get() == b.Get();
}

template <typename StringType, typename StringTraits>
inline bool operator!=(const RegKeyT<StringType, StringTraits>& a,
                       const RegKeyT<StringType, StringTraits>& b) noexcept
{
    return a.Get() != b.Get();
}

template <typename StringType, typename StringTraits>
inline bool operator<(const RegKeyT<StringType, StringTraits>& a,
                      const RegKeyT<StringType, StringTraits>& b) noexcept
{
    return a.Get() < b.Get();
}

template <typename StringType, typename StringTraits>
inline bool operator<=(const RegKeyT<StringType, StringTraits>& a,
                       const RegKeyT<StringType, StringTraits>& b) noexcept
{
    return a.Get() <= b.Get();
}

template <typename StringType, typename StringTraits>
inline bool operator>(const RegKeyT<StringType, StringTraits>& a,
                      const RegKeyT<StringType, StringTraits>& b) noexcept
{
    return a.Get() > b.Get();
}

template <typename StringType, typename StringTraits>
inline bool operator>=(const RegKeyT<StringType, StringTraits>& a,
                       const RegKeyT<StringType, StringTraits>& b) noexcept
{
    return a.Get() >= b.Get();
}


//------------------------------------------------------------------------------
//                          RegKeyT Inline Methods
//------------------------------------------------------------------------------

template <typename StringType, typename StringTraits>
inline RegKeyT<StringType, StringTraits>::RegKeyT(const HKEY hKey) noexcept
    : m_hKey{ hKey }
{
}


template <typename StringType, typename StringTraits>
inline RegKeyT<StringType, StringTraits>::RegKeyT(const HKEY hKeyParent, const StringType& subKey)
{
    Create(hKeyParent, subKey);
}


template <typename StringType, typename StringTraits>
inline RegKeyT<StringType, StringTraits>::RegKeyT(const HKEY hKeyParent,
                                    const StringType& subKey,
                                    const REGSAM desiredAccess)
{
    Create(hKeyParent, subKey, desiredAccess);
}


template <typename StringType, typename StringTraits>
inline RegKeyT<StringType, StringTraits>::RegKeyT(RegKeyT<StringType, StringTraits>&& other) noexcept
    : m_hKey{ other.m_hKey }
{
    // Other doesn't own the handle anymore
    other.m_hKey = nullptr;
}


template <typename StringType, typename StringTraits>
inline RegKeyT<StringType, StringTraits>& RegKeyT<StringType, StringTraits>::operator=(RegKeyT<StringType, StringTraits>&& other) noexcept
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


template <typename StringType, typename StringTraits>
inline RegKeyT<StringType, StringTraits>::~RegKeyT() noexcept
{
    // Release the owned handle (if any)
    Close();
}


template <typename StringType, typename StringTraits>
inline HKEY RegKeyT<StringType, StringTraits>::Get() const noexcept
{
    return m_hKey;
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::Close() noexcept
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


template <typename StringType, typename StringTraits>
inline bool RegKeyT<StringType, StringTraits>::IsValid() const noexcept
{
    return m_hKey != nullptr;
}


template <typename StringType, typename StringTraits>
inline RegKeyT<StringType, StringTraits>::operator bool() const noexcept
{
    return IsValid();
}


template <typename StringType, typename StringTraits>
inline bool RegKeyT<StringType, StringTraits>::IsPredefined() const noexcept
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


template <typename StringType, typename StringTraits>
inline HKEY RegKeyT<StringType, StringTraits>::Detach() noexcept
{
    HKEY hKey = m_hKey;

    // We don't own the HKEY handle anymore
    m_hKey = nullptr;

    // Transfer ownership to the caller
    return hKey;
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::Attach(const HKEY hKey) noexcept
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


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::SwapWith(RegKeyT<StringType, StringTraits>& other) noexcept
{
    // Enable ADL (not necessary in this case, but good practice)
    using std::swap;

    // Swap the raw handle members
    swap(m_hKey, other.m_hKey);
}


template <typename StringType, typename StringTraits>
inline void swap(RegKeyT<StringType, StringTraits>& a, RegKeyT<StringType, StringTraits>& b) noexcept
{
    a.SwapWith(b);
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::Create(
    const HKEY                  hKeyParent,
    const StringType&           subKey,
    const REGSAM                desiredAccess
)
{
    constexpr DWORD kDefaultOptions = REG_OPTION_NON_VOLATILE;

    Create(hKeyParent, subKey, desiredAccess, kDefaultOptions,
        nullptr, // no security attributes,
        nullptr  // no disposition
    );
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::Create(
    const HKEY                  hKeyParent,
    const StringType&           subKey,
    const REGSAM                desiredAccess,
    const DWORD                 options,
    SECURITY_ATTRIBUTES* const  securityAttributes,
    DWORD* const                disposition
)
{
    HKEY hKey = nullptr;
    LSTATUS retCode = ::RegCreateKeyExW(
        hKeyParent,
        StringTraits::ToUtf16(subKey).c_str(),
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


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::Open(
    const HKEY              hKeyParent,
    const StringType&     subKey,
    const REGSAM            desiredAccess
)
{
    HKEY hKey = nullptr;
    LSTATUS retCode = ::RegOpenKeyExW(
        hKeyParent,
        StringTraits::ToUtf16(subKey).c_str(),
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


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TryCreate(
    const HKEY          hKeyParent,
    const StringType&   subKey,
    const REGSAM        desiredAccess
)
{
    constexpr DWORD kDefaultOptions = REG_OPTION_NON_VOLATILE;

    return TryCreate(hKeyParent, subKey, desiredAccess, kDefaultOptions,
        nullptr, // no security attributes,
        nullptr  // no disposition
    );
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TryCreate(
    const HKEY                  hKeyParent,
    const StringType&           subKey,
    const REGSAM                desiredAccess,
    const DWORD                 options,
    SECURITY_ATTRIBUTES* const  securityAttributes,
    DWORD* const                disposition
)
{
    HKEY hKey = nullptr;
    RegResult retCode{ ::RegCreateKeyExW(
        hKeyParent,
        StringTraits::ToUtf16(subKey).c_str(),
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


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TryOpen(
    const HKEY          hKeyParent,
    const StringType&   subKey,
    const REGSAM        desiredAccess
)
{
    HKEY hKey = nullptr;
    RegResult retCode{ ::RegOpenKeyExW(
        hKeyParent,
        StringTraits::ToUtf16(subKey).c_str(),
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


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::SetDwordValue(const StringType& valueName,
                                                             const DWORD data)
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegSetValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
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


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::SetQwordValue(const StringType& valueName,
                                                             const ULONGLONG& data)
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegSetValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
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


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::SetStringValue(const StringType& valueName,
                                                              const StringType& data)
{
    _ASSERTE(IsValid());

    const std::wstring dataUtf16 = StringTraits::ToUtf16(data);

    // String size including the terminating NUL, in bytes
    const DWORD dataSize = details::SafeCastSizeToDword((dataUtf16.length() + 1) * sizeof(wchar_t));

    LSTATUS retCode = ::RegSetValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
        0, // reserved
        REG_SZ,
        reinterpret_cast<const BYTE*>(dataUtf16.c_str()),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot write string value: RegSetValueExW failed." };
    }
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::SetExpandStringValue(const StringType& valueName,
                                                                    const StringType& data)
{
    _ASSERTE(IsValid());

    const std::wstring dataUtf16 = StringTraits::ToUtf16(data);

    // String size including the terminating NUL, in bytes
    const DWORD dataSize = details::SafeCastSizeToDword((dataUtf16.length() + 1) * sizeof(wchar_t));

    LSTATUS retCode = ::RegSetValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
        0, // reserved
        REG_EXPAND_SZ,
        reinterpret_cast<const BYTE*>(dataUtf16.c_str()),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot write expand string value: RegSetValueExW failed." };
    }
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::SetMultiStringValue(
    const StringType& valueName,
    const std::vector<StringType>& data
)
{
    _ASSERTE(IsValid());

    // First, we have to build a double-NUL-terminated multi-string from the input data
    std::vector<wchar_t> multiString;
    if constexpr (std::is_same_v<StringTraits, details::StringTraitUtf16>)
    {
        // UTF-16 mode: the input data already contains UTF-16 strings
        multiString = details::BuildMultiString(data);
    }
    else if constexpr (std::is_same_v<StringTraits, details::StringTraitUtf8>)
    {
        // UTF-8 mode:
        // - First, we need to convert the input UTF-8 string vector into UTF-16
        // - Then, we can invoke the BuildMultiString helper function on it
        multiString = details::BuildMultiString(details::Utf16StringVectorFromUtf8(data));
    }
    else
    {
        static_assert("Unknown/unsupported string trait.");
    }

    // Total size, in bytes, of the whole multi-string structure
    const DWORD dataSize = details::SafeCastSizeToDword(multiString.size() * sizeof(wchar_t));

    LSTATUS retCode = ::RegSetValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
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


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::SetBinaryValue(const StringType& valueName,
                                                              const std::vector<BYTE>& data)
{
    _ASSERTE(IsValid());

    // Total data size, in bytes
    const DWORD dataSize = details::SafeCastSizeToDword(data.size());

    LSTATUS retCode = ::RegSetValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
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


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::SetBinaryValue(
    const StringType& valueName,
    const void* const data,
    const DWORD dataSize
)
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegSetValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
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


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TrySetDwordValue(const StringType& valueName, const DWORD data)
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
        0, // reserved
        REG_DWORD,
        reinterpret_cast<const BYTE*>(&data),
        sizeof(data)
    ) };
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TrySetQwordValue(const StringType& valueName,
                                                                     const ULONGLONG& data)
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
        0, // reserved
        REG_QWORD,
        reinterpret_cast<const BYTE*>(&data),
        sizeof(data)
    ) };
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TrySetStringValue(const StringType& valueName,
                                                                      const StringType& data)
{
    _ASSERTE(IsValid());

    const std::wstring dataUtf16 = StringTraits::ToUtf16(data);

    // String size including the terminating NUL, in bytes
    const DWORD dataSize = details::SafeCastSizeToDword((dataUtf16.length() + 1) * sizeof(wchar_t));

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
        0, // reserved
        REG_SZ,
        reinterpret_cast<const BYTE*>(dataUtf16.c_str()),
        dataSize
    ) };
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TrySetExpandStringValue(
    const StringType& valueName,
    const StringType& data
)
{
    _ASSERTE(IsValid());

    const std::wstring dataUtf16 = StringTraits::ToUtf16(data);

    // String size including the terminating NUL, in bytes
    const DWORD dataSize = details::SafeCastSizeToDword((dataUtf16.length() + 1) * sizeof(wchar_t));

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
        0, // reserved
        REG_EXPAND_SZ,
        reinterpret_cast<const BYTE*>(dataUtf16.c_str()),
        dataSize
    ) };
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TrySetMultiStringValue(
    const StringType& valueName,
    const std::vector<StringType>& data
)
{
    _ASSERTE(IsValid());

    // First, we have to build a double-NUL-terminated multi-string from the input data.

    std::vector<wchar_t> multiString;
    if constexpr (std::is_same_v<StringTraits, details::StringTraitUtf16>)
    {
        // UTF-16 mode: just build the multi-string from the input UTF-16 string vector
        multiString = details::BuildMultiString(data);
    }
    else if constexpr(std::is_same_v<StringTraits, details::StringTraitUtf8>)
    {
        // UTF-8 mode: first convert input UTF-8 string vector to UTF-16 string vector;
        // once you have the UTF-16 string vector, can convert it to multi-string
        multiString = details::BuildMultiString(details::Utf16StringVectorFromUtf8(data));
    }
    else
    {
        static_assert("Unknown/unsupported string trait.");
    }

    // Total size, in bytes, of the whole multi-string structure
    const DWORD dataSize = details::SafeCastSizeToDword(multiString.size() * sizeof(wchar_t));

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
        0, // reserved
        REG_MULTI_SZ,
        reinterpret_cast<const BYTE*>(multiString.data()),
        dataSize
    ) };
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TrySetBinaryValue(
    const StringType& valueName,
    const std::vector<BYTE>& data
)
{
    _ASSERTE(IsValid());

    // Total data size, in bytes
    const DWORD dataSize = details::SafeCastSizeToDword(data.size());

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
        0, // reserved
        REG_BINARY,
        data.data(),
        dataSize
    ) };
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TrySetBinaryValue(
    const StringType& valueName,
    const void* const data,
    const DWORD dataSize
)
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
        0, // reserved
        REG_BINARY,
        static_cast<const BYTE*>(data),
        dataSize
    ) };
}


template <typename StringType, typename StringTraits>
inline DWORD RegKeyT<StringType, StringTraits>::GetDwordValue(const StringType& valueName) const
{
    _ASSERTE(IsValid());

    DWORD data = 0;                  // to be read from the registry
    DWORD dataSize = sizeof(data);   // size of data, in bytes

    constexpr DWORD flags = RRF_RT_REG_DWORD;
    LSTATUS retCode = ::RegGetValueW(
        m_hKey,
        nullptr, // no subkey
        StringTraits::ToUtf16(valueName).c_str(),
        flags,
        nullptr, // type not required
        &data,
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot get DWORD value: RegGetValueW failed." };
    }

    return data;
}


template <typename StringType, typename StringTraits>
inline ULONGLONG RegKeyT<StringType, StringTraits>::GetQwordValue(const StringType& valueName) const
{
    _ASSERTE(IsValid());

    ULONGLONG data = 0;              // to be read from the registry
    DWORD dataSize = sizeof(data);   // size of data, in bytes

    constexpr DWORD flags = RRF_RT_REG_QWORD;
    LSTATUS retCode = ::RegGetValueW(
        m_hKey,
        nullptr, // no subkey
        StringTraits::ToUtf16(valueName).c_str(),
        flags,
        nullptr, // type not required
        &data,
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot get QWORD value: RegGetValueW failed." };
    }

    return data;
}


template <typename StringType, typename StringTraits>
inline StringType RegKeyT<StringType, StringTraits>::GetStringValue(const StringType& valueName) const
{
    _ASSERTE(IsValid());

    std::wstring result;    // to be read from the registry
    DWORD dataSize = 0;     // size of the string data, in bytes

    constexpr DWORD flags = RRF_RT_REG_SZ;

    const std::wstring valueNameUtf16 = StringTraits::ToUtf16(valueName);

    LSTATUS retCode = ERROR_MORE_DATA;

    while (retCode == ERROR_MORE_DATA)
    {
        // Get the size of the result string
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,    // type not required
            nullptr,    // output buffer not needed now
            &dataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            throw RegException{ retCode, "Cannot get the size of the string value: RegGetValueW failed." };
        }

        // Allocate a string of proper size.
        // Note that dataSize is in bytes and includes the terminating NUL;
        // we have to convert the size from bytes to wchar_ts for wstring::resize.
        result.resize(dataSize / sizeof(wchar_t));

        // Call RegGetValue for the second time to read the string's content
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,       // type not required
            result.data(), // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot get the string value: RegGetValueW failed." };
    }

    // Remove the NUL terminator scribbled by RegGetValue from the wstring
    result.resize((dataSize / sizeof(wchar_t)) - 1);

    return StringTraits::ConstructFromUtf16(result);
}


template <typename StringType, typename StringTraits>
inline StringType RegKeyT<StringType, StringTraits>::GetExpandStringValue(
    const StringType& valueName,
    const ExpandStringOption expandOption
) const
{
    _ASSERTE(IsValid());

    std::wstring result;    // to be read from the registry
    DWORD dataSize = 0;     // size of the expand string data, in bytes


    DWORD flags = RRF_RT_REG_EXPAND_SZ;

    // Adjust the flag for RegGetValue considering the expand string option specified by the caller
    if (expandOption == ExpandStringOption::DontExpand)
    {
        flags |= RRF_NOEXPAND;
    }


    const std::wstring valueNameUtf16 = StringTraits::ToUtf16(valueName);

    LSTATUS retCode = ERROR_MORE_DATA;

    while (retCode == ERROR_MORE_DATA)
    {
        // Get the size of the result string
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,    // type not required
            nullptr,    // output buffer not needed now
            &dataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            throw RegException{ retCode,
                                "Cannot get the size of the expand string value: RegGetValueW failed." };
        }

        // Allocate a string of proper size.
        // Note that dataSize is in bytes and includes the terminating NUL;
        // we have to convert the size from bytes to wchar_ts for wstring::resize.
        result.resize(dataSize / sizeof(wchar_t));

        // Call RegGetValue for the second time to read the string's content
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,       // type not required
            result.data(), // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot get the expand string value: RegGetValueW failed." };
    }

    // Remove the NUL terminator scribbled by RegGetValue from the wstring
    result.resize((dataSize / sizeof(wchar_t)) - 1);

    return StringTraits::ConstructFromUtf16(result);
}


template <typename StringType, typename StringTraits>
inline std::vector<StringType>
RegKeyT<StringType, StringTraits>::GetMultiStringValue(const StringType& valueName) const
{
    _ASSERTE(IsValid());

    // Room for the result multi-string, to be read from the registry
    std::vector<wchar_t> multiString;

    // Size of the multi-string, in bytes
    DWORD dataSize = 0;

    constexpr DWORD flags = RRF_RT_REG_MULTI_SZ;

    const std::wstring valueNameUtf16 = StringTraits::ToUtf16(valueName);

    LSTATUS retCode = ERROR_MORE_DATA;

    while (retCode == ERROR_MORE_DATA)
    {
        // Request the size of the multi-string, in bytes
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,    // type not required
            nullptr,    // output buffer not needed now
            &dataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            throw RegException{ retCode,
                                "Cannot get the size of the multi-string value: RegGetValueW failed." };
        }

        // Allocate room for the result multi-string.
        // Note that dataSize is in bytes, but our vector<wchar_t>::resize method requires size
        // to be expressed in wchar_ts.
        multiString.resize(dataSize / sizeof(wchar_t));

        // Call RegGetValue for the second time to read the multi-string's content into the vector
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,                // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,                // type not required
            multiString.data(),     // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot get the multi-string value: RegGetValueW failed." };
    }

    // Resize vector to the actual size returned by the last call to RegGetValue.
    // Note that the vector is a vector of wchar_ts, instead the size returned by RegGetValue
    // is in bytes, so we have to scale from bytes to wchar_t count.
    multiString.resize(dataSize / sizeof(wchar_t));

    // Convert the double-null-terminated string structure to a vector<wstring>,
    // and return that back to the caller
    const std::vector<std::wstring> resultStringsUtf16 = details::ParseMultiString(multiString);

    if constexpr (std::is_same_v<StringTraits, details::StringTraitUtf16>)
    {
        // Convert the double-null-terminated string structure to a vector<wstring>,
        // and return that back to the caller
        return details::ParseMultiString(multiString);
    }
    else if constexpr (std::is_same_v<StringTraits, details::StringTraitUtf8>)
    {
        // Convert the double-null-terminated string structure to a vector<wstring>
        // via ParseMultiString(),
        // and then convert them to UTF-8 for returning them back to the caller
        return details::Utf8StringVectorFromUtf16(details::ParseMultiString(multiString));
    }
    else
    {
        static_assert("Unknown/unsupported string traits.");
    }
}


template <typename StringType, typename StringTraits>
inline std::vector<BYTE>
RegKeyT<StringType, StringTraits>::GetBinaryValue(const StringType& valueName) const
{
    _ASSERTE(IsValid());

    // Room for the binary data, to be read from the registry
    std::vector<BYTE> binaryData;

    // Size of binary data, in bytes
    DWORD dataSize = 0;

    constexpr DWORD flags = RRF_RT_REG_BINARY;

    const std::wstring valueNameUtf16 = StringTraits::ToUtf16(valueName);

    LSTATUS retCode = ERROR_MORE_DATA;

    while (retCode == ERROR_MORE_DATA)
    {
        // Request the size of the binary data, in bytes
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,    // type not required
            nullptr,    // output buffer not needed now
            &dataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            throw RegException{ retCode,
                                "Cannot get the size of the binary data: RegGetValueW failed." };
        }

        // Allocate a buffer of proper size to store the binary data
        binaryData.resize(dataSize);

        // Handle the special case of zero-length binary data:
        // If the binary data value in the registry is empty, just return an empty vector.
        if (dataSize == 0)
        {
            _ASSERTE(binaryData.empty());
            return binaryData;
        }

        // Call RegGetValue for the second time to read the binary data content into the vector
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,            // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,            // type not required
            binaryData.data(),  // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot get the binary data: RegGetValueW failed." };
    }

    // Resize vector to the actual size returned by the last call to RegGetValue
    binaryData.resize(dataSize);

    return binaryData;
}


template <typename StringType, typename StringTraits>
inline RegExpected<DWORD>
RegKeyT<StringType, StringTraits>::TryGetDwordValue(const StringType& valueName) const
{
    _ASSERTE(IsValid());

    using RegValueType = DWORD;

    DWORD data = 0;                  // to be read from the registry
    DWORD dataSize = sizeof(data);   // size of data, in bytes

    constexpr DWORD flags = RRF_RT_REG_DWORD;
    LSTATUS retCode = ::RegGetValueW(
        m_hKey,
        nullptr, // no subkey
        StringTraits::ToUtf16(valueName).c_str(),
        flags,
        nullptr, // type not required
        &data,
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        return details::MakeRegExpectedWithError<RegValueType>(retCode);
    }

    return RegExpected<RegValueType>{ data };
}


template <typename StringType, typename StringTraits>
inline RegExpected<ULONGLONG>
RegKeyT<StringType, StringTraits>::TryGetQwordValue(const StringType& valueName) const
{
    _ASSERTE(IsValid());

    using RegValueType = ULONGLONG;

    ULONGLONG data = 0;              // to be read from the registry
    DWORD dataSize = sizeof(data);   // size of data, in bytes

    constexpr DWORD flags = RRF_RT_REG_QWORD;
    LSTATUS retCode = ::RegGetValueW(
        m_hKey,
        nullptr, // no subkey
        StringTraits::ToUtf16(valueName).c_str(),
        flags,
        nullptr, // type not required
        &data,
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        return details::MakeRegExpectedWithError<RegValueType>(retCode);
    }

    return RegExpected<RegValueType>{ data };
}


template <typename StringType, typename StringTraits>
inline RegExpected<StringType>
RegKeyT<StringType, StringTraits>::TryGetStringValue(const StringType& valueName) const
{
    _ASSERTE(IsValid());

    using RegValueType = StringType;

    constexpr DWORD flags = RRF_RT_REG_SZ;

    std::wstring result;

    DWORD dataSize = 0; // size of the string data, in bytes

    const std::wstring valueNameUtf16 = StringTraits::ToUtf16(valueName);

    LSTATUS retCode = ERROR_MORE_DATA;

    while (retCode == ERROR_MORE_DATA)
    {
        // Get the size of the result string
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,    // type not required
            nullptr,    // output buffer not needed now
            &dataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            return details::MakeRegExpectedWithError<RegValueType>(retCode);
        }

        // Allocate a string of proper size.
        // Note that dataSize is in bytes and includes the terminating NUL;
        // we have to convert the size from bytes to wchar_ts for wstring::resize.
        result.resize(dataSize / sizeof(wchar_t));

        // Call RegGetValue for the second time to read the string's content
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,       // type not required
            result.data(), // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        return details::MakeRegExpectedWithError<RegValueType>(retCode);
    }

    // Remove the NUL terminator scribbled by RegGetValue from the wstring
    result.resize((dataSize / sizeof(wchar_t)) - 1);

    return RegExpected<RegValueType>{ StringTraits::ConstructFromUtf16(result) };
}


template <typename StringType, typename StringTraits>
inline RegExpected<StringType> RegKeyT<StringType, StringTraits>::TryGetExpandStringValue(
    const StringType& valueName,
    const ExpandStringOption expandOption
) const
{
    _ASSERTE(IsValid());

    using RegValueType = StringType;

    DWORD flags = RRF_RT_REG_EXPAND_SZ;

    // Adjust the flag for RegGetValue considering the expand string option specified by the caller
    if (expandOption == ExpandStringOption::DontExpand)
    {
        flags |= RRF_NOEXPAND;
    }

    std::wstring result;
    DWORD dataSize = 0; // size of the expand string data, in bytes

    const std::wstring valueNameUtf16 = StringTraits::ToUtf16(valueName);

    LSTATUS retCode = ERROR_MORE_DATA;

    while (retCode == ERROR_MORE_DATA)
    {
        // Get the size of the result string
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,    // type not required
            nullptr,    // output buffer not needed now
            &dataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            return details::MakeRegExpectedWithError<RegValueType>(retCode);
        }

        // Allocate a string of proper size.
        // Note that dataSize is in bytes and includes the terminating NUL;
        // we have to convert the size from bytes to wchar_ts for wstring::resize.
        result.resize(dataSize / sizeof(wchar_t));

        // Call RegGetValue for the second time to read the string's content
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,       // type not required
            result.data(), // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        return details::MakeRegExpectedWithError<RegValueType>(retCode);
    }

    // Remove the NUL terminator scribbled by RegGetValue from the wstring
    result.resize((dataSize / sizeof(wchar_t)) - 1);

    return RegExpected<RegValueType>{ StringTraits::ConstructFromUtf16(result) };
}


template <typename StringType, typename StringTraits>
inline RegExpected<std::vector<StringType>>
RegKeyT<StringType, StringTraits>::TryGetMultiStringValue(const StringType& valueName) const
{
    _ASSERTE(IsValid());

    using RegValueType = std::vector<StringType>;

    constexpr DWORD flags = RRF_RT_REG_MULTI_SZ;

    // Room for the result multi-string
    std::vector<wchar_t> data;

    // Size of the multi-string, in bytes
    DWORD dataSize = 0;

    const std::wstring valueNameUtf16 = StringTraits::ToUtf16(valueName);

    LSTATUS retCode = ERROR_MORE_DATA;

    while (retCode == ERROR_MORE_DATA)
    {
        // Request the size of the multi-string, in bytes
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,    // type not required
            nullptr,    // output buffer not needed now
            &dataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            return details::MakeRegExpectedWithError<RegValueType>(retCode);
        }

        // Allocate room for the result multi-string.
        // Note that dataSize is in bytes, but our vector<wchar_t>::resize method requires size
        // to be expressed in wchar_ts.
        data.resize(dataSize / sizeof(wchar_t));

        // Call RegGetValue for the second time to read the multi-string's content into the vector
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,        // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,        // type not required
            data.data(),    // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        return details::MakeRegExpectedWithError<RegValueType>(retCode);
    }

    // Resize vector to the actual size returned by the last call to RegGetValue.
    // Note that the vector is a vector of wchar_ts, instead the size returned by RegGetValue
    // is in bytes, so we have to scale from bytes to wchar_t count.
    data.resize(dataSize / sizeof(wchar_t));

    // Convert the double-null-terminated string structure to a vector<wstring> or vector<string>,
    // and return that back to the caller
    if constexpr (std::is_same_v<StringTraits, details::StringTraitUtf16>)
    {
        return RegExpected<RegValueType>(details::ParseMultiString(data));
    }
    else if constexpr (std::is_same_v<StringTraits, details::StringTraitUtf8>)
    {
        return RegExpected<RegValueType>(
            details::Utf8StringVectorFromUtf16(details::ParseMultiString(data)));
    }
    else
    {
        static_assert("Unknown/unsupported string trait.")
    }
}


template <typename StringType, typename StringTraits>
inline RegExpected<std::vector<BYTE>>
RegKeyT<StringType, StringTraits>::TryGetBinaryValue(const StringType& valueName) const
{
    _ASSERTE(IsValid());

    using RegValueType = std::vector<BYTE>;

    constexpr DWORD flags = RRF_RT_REG_BINARY;

    // Room for the binary data
    std::vector<BYTE> data;

    DWORD dataSize = 0; // size of binary data, in bytes

    const std::wstring valueNameUtf16 = StringTraits::ToUtf16(valueName);

    LSTATUS retCode = ERROR_MORE_DATA;

    while (retCode == ERROR_MORE_DATA)
    {
        // Request the size of the binary data, in bytes
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,    // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,    // type not required
            nullptr,    // output buffer not needed now
            &dataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            return details::MakeRegExpectedWithError<RegValueType>(retCode);
        }

        // Allocate a buffer of proper size to store the binary data
        data.resize(dataSize);

        // Handle the special case of zero-length binary data:
        // If the binary data value in the registry is empty, just return
        if (dataSize == 0)
        {
            _ASSERTE(data.empty());
            return RegExpected<RegValueType>{ data };
        }

        // Call RegGetValue for the second time to read the binary data content into the vector
        retCode = ::RegGetValueW(
            m_hKey,
            nullptr,        // no subkey
            valueNameUtf16.c_str(),
            flags,
            nullptr,        // type not required
            data.data(),    // output buffer
            &dataSize
        );
    }

    if (retCode != ERROR_SUCCESS)
    {
        return details::MakeRegExpectedWithError<RegValueType>(retCode);
    }

    // Resize vector to the actual size returned by the last call to RegGetValue
    data.resize(dataSize);

    return RegExpected<RegValueType>{ data };
}


template <typename StringType, typename StringTraits>
inline std::vector<StringType> RegKeyT<StringType, StringTraits>::EnumSubKeys() const
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
    std::vector<StringType> subkeyNames;

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
        subkeyNames.emplace_back(StringTraits::ConstructFromUtf16({ nameBuffer.get(), subKeyNameLen }));
    }

    return subkeyNames;
}


template <typename StringType, typename StringTraits>
inline std::vector<std::pair<StringType, DWORD>>
RegKeyT<StringType, StringTraits>::EnumValues() const
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
    std::vector<std::pair<StringType, DWORD>> valueInfo;

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
            StringTraits::ConstructFromUtf16(std::wstring{ nameBuffer.get(), valueNameLen }),
            valueType
        );
    }

    return valueInfo;
}


template <typename StringType, typename StringTraits>
inline bool RegKeyT<StringType, StringTraits>::ContainsValue(const StringType& valueName) const
{
    _ASSERTE(IsValid());

    // Invoke RegGetValueW to just check if the input value exists under the current key
    LSTATUS retCode = ::RegGetValueW(
        m_hKey,             // current key
        nullptr,            // no subkey - check value in current key
        StringTraits::ToUtf16(valueName).c_str(),  // value name
        RRF_RT_ANY,         // no type restriction on this value
        nullptr,            // we don't need to know the type of the value
        nullptr, nullptr    // we don't need the actual value, just to check if it exists
    );
    if (retCode == ERROR_SUCCESS)
    {
        // The value exists under the current key
        return true;
    }
    else if (retCode == ERROR_FILE_NOT_FOUND)
    {
        // The value does *not* exist under the current key
        return false;
    }
    else
    {
        // Some other error occurred - signal it by throwing an exception
        throw RegException{
            retCode,
            "RegGetValueW failed when checking if the current key contains the specified value."
        };
    }
}


template <typename StringType, typename StringTraits>
inline bool RegKeyT<StringType, StringTraits>::ContainsSubKey(const StringType& subKey) const
{
    _ASSERTE(IsValid());

    // Let's try and open the specified subKey, then check the return code
    // of RegOpenKeyExW to figure out if the subKey exists or not.
    HKEY hSubKey = nullptr;
    LSTATUS retCode = ::RegOpenKeyExW(
        m_hKey,
        StringTraits::ToUtf16(subKey).c_str(),
        0,
        KEY_READ,
        &hSubKey
    );
    if (retCode == ERROR_SUCCESS)
    {
        // We were able to open the specified sub-key, so the sub-key does exist.
        //
        // Don't forget to close the sub-key opened for this testing purpose!
        ::RegCloseKey(hSubKey);
        hSubKey = nullptr;

        return true;
    }
    else if ((retCode == ERROR_FILE_NOT_FOUND) || (retCode == ERROR_PATH_NOT_FOUND))
    {
        // The specified sub-key does not exist
        return false;
    }
    else
    {
        // Some other error occurred - signal it by throwing an exception
        throw RegException{
            retCode,
            "RegOpenKeyExW failed when checking if the current key contains the specified sub-key."
        };
    }
}


template <typename StringType, typename StringTraits>
inline RegExpected<std::vector<StringType>> RegKeyT<StringType, StringTraits>::TryEnumSubKeys() const
{
    _ASSERTE(IsValid());

    using ReturnType = std::vector<StringType>;

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
        return details::MakeRegExpectedWithError<ReturnType>(retCode);
    }

    // NOTE: According to the MSDN documentation, the size returned for subkey name max length
    // does *not* include the terminating NUL, so let's add +1 to take it into account
    // when I allocate the buffer for reading subkey names.
    maxSubKeyNameLen++;

    // Preallocate a buffer for the subkey names
    auto nameBuffer = std::make_unique<wchar_t[]>(maxSubKeyNameLen);

    // The result subkey names will be stored here
    std::vector<StringType> subkeyNames;

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
            return details::MakeRegExpectedWithError<ReturnType>(retCode);
        }

        // On success, the ::RegEnumKeyEx API writes the length of the
        // subkey name in the subKeyNameLen output parameter
        // (not including the terminating NUL).
        // So I can build a wstring based on that length.
        subkeyNames.emplace_back(StringTraits::ConstructFromUtf16({ nameBuffer.get(), subKeyNameLen }));
    }

    return RegExpected<ReturnType>{ subkeyNames };
}


template <typename StringType, typename StringTraits>
inline RegExpected<std::vector<std::pair<StringType, DWORD>>>
RegKeyT<StringType, StringTraits>::TryEnumValues() const
{
    _ASSERTE(IsValid());

    using ReturnType = std::vector<std::pair<StringType, DWORD>>;

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
        return details::MakeRegExpectedWithError<ReturnType>(retCode);
    }

    // NOTE: According to the MSDN documentation, the size returned for value name max length
    // does *not* include the terminating NUL, so let's add +1 to take it into account
    // when I allocate the buffer for reading value names.
    maxValueNameLen++;

    // Preallocate a buffer for the value names
    auto nameBuffer = std::make_unique<wchar_t[]>(maxValueNameLen);

    // The value names and types will be stored here
    std::vector<std::pair<StringType, DWORD>> valueInfo;

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
            return details::MakeRegExpectedWithError<ReturnType>(retCode);
        }

        // On success, the RegEnumValue API writes the length of the
        // value name in the valueNameLen output parameter
        // (not including the terminating NUL).
        // So we can build a wstring based on that.
        valueInfo.emplace_back(
            StringTraits::ConstructFromUtf16(std::wstring{ nameBuffer.get(), valueNameLen }),
            valueType
        );
    }

    return RegExpected<ReturnType>{ valueInfo };
}


template <typename StringType, typename StringTraits>
inline RegExpected<bool>
RegKeyT<StringType, StringTraits>::TryContainsValue(const StringType& valueName) const
{
    _ASSERTE(IsValid());

    // Invoke RegGetValueW to just check if the input value exists under the current key
    LSTATUS retCode = ::RegGetValueW(
        m_hKey,             // current key
        nullptr,            // no subkey - check value in current key
        StringTraits::ToUtf16(valueName).c_str(),  // value name
        RRF_RT_ANY,         // no type restriction on this value
        nullptr,            // we don't need to know the type of the value
        nullptr, nullptr    // we don't need the actual value, just to check if it exists
    );
    if (retCode == ERROR_SUCCESS)
    {
        // The value exists under the current key
        return RegExpected<bool>{ true };
    }
    else if (retCode == ERROR_FILE_NOT_FOUND)
    {
        // The value does *not* exist under the current key
        return RegExpected<bool>{ false };
    }
    else
    {
        // Some other error occurred
        return details::MakeRegExpectedWithError<bool>(retCode);
    }
}


template <typename StringType, typename StringTraits>
inline RegExpected<bool>
RegKeyT<StringType, StringTraits>::TryContainsSubKey(const StringType& subKey) const
{
    _ASSERTE(IsValid());

    // Let's try and open the specified subKey, then check the return code
    // of RegOpenKeyExW to figure out if the subKey exists or not.
    HKEY hSubKey = nullptr;
    LSTATUS retCode = ::RegOpenKeyExW(
        m_hKey,
        StringTraits::ToUtf16(subKey).c_str(),
        0,
        KEY_READ,
        &hSubKey
    );
    if (retCode == ERROR_SUCCESS)
    {
        // We were able to open the specified sub-key, so the sub-key does exist.
        //
        // Don't forget to close the sub-key opened for this testing purpose!
        ::RegCloseKey(hSubKey);
        hSubKey = nullptr;

        return RegExpected<bool>{ true };
    }
    else if ((retCode == ERROR_FILE_NOT_FOUND) || (retCode == ERROR_PATH_NOT_FOUND))
    {
        // The specified sub-key does not exist
        return RegExpected<bool>{ false };
    }
    else
    {
        // Some other error occurred
        return details::MakeRegExpectedWithError<bool>(retCode);
    }
}


template <typename StringType, typename StringTraits>
inline DWORD RegKeyT<StringType, StringTraits>::QueryValueType(const StringType& valueName) const
{
    _ASSERTE(IsValid());

    DWORD typeId = 0;     // will be returned by RegQueryValueEx

    LSTATUS retCode = ::RegQueryValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
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


template <typename StringType, typename StringTraits>
inline RegExpected<DWORD>
RegKeyT<StringType, StringTraits>::TryQueryValueType(const StringType& valueName) const
{
    _ASSERTE(IsValid());

    using ReturnType = DWORD;

    DWORD typeId = 0;     // will be returned by RegQueryValueEx

    LSTATUS retCode = ::RegQueryValueExW(
        m_hKey,
        StringTraits::ToUtf16(valueName).c_str(),
        nullptr,    // reserved
        &typeId,
        nullptr,    // not interested
        nullptr     // not interested
    );

    if (retCode != ERROR_SUCCESS)
    {
        return details::MakeRegExpectedWithError<ReturnType>(retCode);
    }

    return RegExpected<ReturnType>{ typeId };
}


template <typename StringType, typename StringTraits>
inline typename RegKeyT<StringType, StringTraits>::InfoKey
RegKeyT<StringType, StringTraits>::QueryInfoKey() const
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


template <typename StringType, typename StringTraits>
inline typename RegExpected< typename RegKeyT<StringType, StringTraits>::InfoKey>
RegKeyT<StringType, StringTraits>::TryQueryInfoKey() const
{
    _ASSERTE(IsValid());

    using ReturnType = RegKeyT::InfoKey;

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
        return details::MakeRegExpectedWithError<ReturnType>(retCode);
    }

    return RegExpected<ReturnType>{ infoKey };
}


template <typename StringType, typename StringTraits>
inline typename RegKeyT<StringType, StringTraits>::KeyReflection
RegKeyT<StringType, StringTraits>::QueryReflectionKey() const
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


template <typename StringType, typename StringTraits>
inline RegExpected<typename RegKeyT<StringType, StringTraits>::KeyReflection>
RegKeyT<StringType, StringTraits>::TryQueryReflectionKey() const
{
    using ReturnType = RegKeyT::KeyReflection;

    BOOL isReflectionDisabled = FALSE;
    LSTATUS retCode = ::RegQueryReflectionKey(m_hKey, &isReflectionDisabled);
    if (retCode != ERROR_SUCCESS)
    {
        return details::MakeRegExpectedWithError<ReturnType>(retCode);
    }

    KeyReflection keyReflection = isReflectionDisabled ? KeyReflection::ReflectionDisabled
                                                       : KeyReflection::ReflectionEnabled;
    return RegExpected<ReturnType>{ keyReflection };
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::DeleteValue(const StringType& valueName)
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegDeleteValueW(m_hKey, StringTraits::ToUtf16(valueName).c_str());
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegDeleteValueW failed." };
    }
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TryDeleteValue(const StringType& valueName)
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegDeleteValueW(m_hKey, StringTraits::ToUtf16(valueName).c_str()) };
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::DeleteKey(const StringType& subKey,
                                                         const REGSAM desiredAccess)
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegDeleteKeyExW(m_hKey,
                                        StringTraits::ToUtf16(subKey).c_str(),
                                        desiredAccess,
                                        0);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegDeleteKeyExW failed." };
    }
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TryDeleteKey(const StringType& subKey,
                                                                 const REGSAM desiredAccess)
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegDeleteKeyExW(m_hKey,
                                        StringTraits::ToUtf16(subKey).c_str(),
                                        desiredAccess,
                                        0) };
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::DeleteTree(const StringType& subKey)
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegDeleteTreeW(m_hKey, StringTraits::ToUtf16(subKey).c_str());
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegDeleteTreeW failed." };
    }
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TryDeleteTree(const StringType& subKey)
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegDeleteTreeW(m_hKey, StringTraits::ToUtf16(subKey).c_str()) };
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::CopyTree(const StringType& sourceSubKey,
                                           const RegKeyT<StringType, StringTraits>& destKey)
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegCopyTreeW(
        m_hKey,
        StringTraits::ToUtf16(sourceSubKey).c_str(),
        destKey.Get()
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegCopyTreeW failed." };
    }
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TryCopyTree(
    const StringType& sourceSubKey,
    const RegKeyT<StringType, StringTraits>& destKey
)
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegCopyTreeW(
        m_hKey,
        StringTraits::ToUtf16(sourceSubKey).c_str(),
        destKey.Get())
    };
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::FlushKey()
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegFlushKey(m_hKey);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegFlushKey failed." };
    }
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TryFlushKey()
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegFlushKey(m_hKey) };
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::LoadKey(const StringType& subKey,
                                                       const StringType& filename)
{
    Close();

    LSTATUS retCode = ::RegLoadKeyW(
        m_hKey,
        StringTraits::ToUtf16(subKey),
        StringTraits::ToUtf16(filename)
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegLoadKeyW failed." };
    }
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TryLoadKey(
    const StringType& subKey,
    const StringType& filename)
{
    Close();

    return RegResult{ ::RegLoadKeyW(m_hKey,
                                    StringTraits::ToUtf16(subKey).c_str(),
                                    StringTraits::ToUtf16(filename).c_str()) };
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::SaveKey(
    const StringType& filename,
    SECURITY_ATTRIBUTES* const securityAttributes
) const
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegSaveKeyW(
        m_hKey,
        StringTraits::ToUtf16(filename).c_str(),
        securityAttributes
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegSaveKeyW failed." };
    }
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TrySaveKey(
    const StringType& filename,
    SECURITY_ATTRIBUTES* const securityAttributes
) const
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegSaveKeyW(m_hKey,
                                    StringTraits::ToUtf16(filename).c_str(),
                                    securityAttributes)
                    };
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::EnableReflectionKey()
{
    LSTATUS retCode = ::RegEnableReflectionKey(m_hKey);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegEnableReflectionKey failed." };
    }
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TryEnableReflectionKey()
{
    return RegResult{ ::RegEnableReflectionKey(m_hKey) };
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::DisableReflectionKey()
{
    LSTATUS retCode = ::RegDisableReflectionKey(m_hKey);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegDisableReflectionKey failed." };
    }
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TryDisableReflectionKey()
{
    return RegResult{ ::RegDisableReflectionKey(m_hKey) };
}


template <typename StringType, typename StringTraits>
inline void RegKeyT<StringType, StringTraits>::ConnectRegistry(
    const StringType& machineName,
    const HKEY hKeyPredefined
)
{
    // Safely close any previously opened key
    Close();

    HKEY hKeyResult = nullptr;
    LSTATUS retCode = ::RegConnectRegistryW(StringTraits::ToUtf16(machineName).c_str(),
                                            hKeyPredefined,
                                            &hKeyResult);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegConnectRegistryW failed." };
    }

    // Take ownership of the result key
    m_hKey = hKeyResult;
}


template <typename StringType, typename StringTraits>
inline RegResult RegKeyT<StringType, StringTraits>::TryConnectRegistry(
    const StringType& machineName,
    const HKEY hKeyPredefined
)
{
    // Safely close any previously opened key
    Close();

    HKEY hKeyResult = nullptr;
    RegResult retCode{ ::RegConnectRegistryW(StringTraits::ToUtf16(machineName).c_str(),
                                             hKeyPredefined,
                                             &hKeyResult) };
    if (retCode.Failed())
    {
        return retCode;
    }

    // Take ownership of the result key
    m_hKey = hKeyResult;

    _ASSERTE(retCode.IsOk());
    return retCode;
}


template <>
inline std::wstring
RegKeyT<std::wstring, details::StringTraitUtf16>::RegTypeToString(const DWORD regType)
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


template <>
inline std::string
RegKeyT<std::string, details::StringTraitUtf8>::RegTypeToString(const DWORD regType)
{
    switch (regType)
    {
        case REG_SZ:        return "REG_SZ";
        case REG_EXPAND_SZ: return "REG_EXPAND_SZ";
        case REG_MULTI_SZ:  return "REG_MULTI_SZ";
        case REG_DWORD:     return "REG_DWORD";
        case REG_QWORD:     return "REG_QWORD";
        case REG_BINARY:    return "REG_BINARY";

        default:            return "Unknown/unsupported registry type";
    }
}


} // namespace winreg


#endif // GIOVANNI_DICANIO_WINREG_WINREG_HPP_INCLUDED

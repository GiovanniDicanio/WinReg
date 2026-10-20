////////////////////////////////////////////////////////////////////////////////
//
//         *** Modern C++ Wrappers Around Windows Registry C API ***
//
//       Copyright (C) by Giovanni Dicanio (giovanni.dicanio@gmail.com)
//
//----------------------------------------------------------------------------
// SPDX-License-Identifier: MIT
////////////////////////////////////////////////////////////////////////////////

#ifndef GIOVANNI_DICANIO_WINREG_REGKEY_IMPL_HPP
#define GIOVANNI_DICANIO_WINREG_REGKEY_IMPL_HPP

//
// NOTE: This file is included by "RegKey.hpp" and contains the inline implementations
// of the RegKey class methods.
//


#include <Windows.h>        // Windows Platform SDK
#include <crtdbg.h>         // _ASSERTE
#include <wchar.h>          // wcsnlen

#include <memory>           // std::unique_ptr, std::make_unique
#include <string>           // std::string, std::wstring
#include <utility>          // std::swap, std::pair, std::move
#include <vector>           // std::vector


#ifndef GIOVANNI_DICANIO_WINREG_USE_CPP17

    //
    // <span> is not available in C++17 mode; it requires C++20.
    //

    #include <span>             // std::span

#endif // GIOVANNI_DICANIO_WINREG_USE_CPP17


#include "detail/Config.hpp"
#include "detail/ExpandStrings.hpp"
#include "detail/MultiString.hpp"
#include "detail/SafeCast.hpp"


namespace winreg
{

#ifdef GIOVANNI_DICANIO_WINREG_USE_CPP17

//
// C++17 code (use shortcut implementation in C++20)
//

//------------------------------------------------------------------------------
//          Overloads of relational comparison operators for RegKey
//------------------------------------------------------------------------------

[[nodiscard]] inline bool operator==(const RegKey& a, const RegKey& b) noexcept
{
    return a.Get() == b.Get();
}


[[nodiscard]] inline bool operator!=(const RegKey& a, const RegKey& b) noexcept
{
    return a.Get() != b.Get();
}


[[nodiscard]] inline bool operator<(const RegKey& a, const RegKey& b) noexcept
{
    return a.Get() < b.Get();
}


[[nodiscard]] inline bool operator<=(const RegKey& a, const RegKey& b) noexcept
{
    return a.Get() <= b.Get();
}


[[nodiscard]] inline bool operator>(const RegKey& a, const RegKey& b) noexcept
{
    return a.Get() > b.Get();
}


[[nodiscard]] inline bool operator>=(const RegKey& a, const RegKey& b) noexcept
{
    return a.Get() >= b.Get();
}

#endif // GIOVANNI_DICANIO_WINREG_USE_CPP17


//------------------------------------------------------------------------------
//                     RegKey Inline Method Implementation
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
    if (this != &other)
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
        if (!IsPredefined())
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

    if ((m_hKey == HKEY_CURRENT_USER)
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
    const std::wstring& subKey,
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


inline void RegKey::SetQwordValue(const std::wstring& valueName, const ULONGLONG data)
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
    const DWORD dataSize = detail::SafeCastSizeToDword((data.length() + 1) * sizeof(wchar_t));

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
    const DWORD dataSize = detail::SafeCastSizeToDword((data.length() + 1) * sizeof(wchar_t));

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
    const std::vector<wchar_t> multiString = detail::BuildMultiString(data);

    // Total size, in bytes, of the whole multi-string structure
    const DWORD dataSize = detail::SafeCastSizeToDword(multiString.size() * sizeof(wchar_t));

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


#ifndef GIOVANNI_DICANIO_WINREG_USE_CPP17

//
// C++20 code
//

inline void RegKey::SetBinaryValue(
    const std::wstring& valueName,
    const std::span<const BYTE> data
)
{
    _ASSERTE(IsValid());

    const DWORD dataSize = detail::SafeCastSizeToDword(data.size());

    SetBinaryValue(valueName, data.data(), dataSize);
}


inline void RegKey::SetBinaryValue(const std::wstring& valueName, const std::vector<BYTE>& data)
{
    _ASSERTE(IsValid());

    SetBinaryValue(valueName, std::span{ data });
}

#else

//
// C++17-compatible code
//

inline void RegKey::SetBinaryValue(const std::wstring& valueName, const std::vector<BYTE>& data)
{
    _ASSERTE(IsValid());

    // Total data size, in bytes
    const DWORD dataSize = detail::SafeCastSizeToDword(data.size());

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

#endif // GIOVANNI_DICANIO_WINREG_USE_CPP17


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
                                          const ULONGLONG data) noexcept
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
                                           const std::wstring& data)
{
    _ASSERTE(IsValid());

    // String size including the terminating NUL, in bytes
    const DWORD dataSize = detail::SafeCastSizeToDword((data.length() + 1) * sizeof(wchar_t));

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
                                                 const std::wstring& data)
{
    _ASSERTE(IsValid());

    // String size including the terminating NUL, in bytes
    const DWORD dataSize = detail::SafeCastSizeToDword((data.length() + 1) * sizeof(wchar_t));

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
    const std::vector<wchar_t> multiString = detail::BuildMultiString(data);

    // Total size, in bytes, of the whole multi-string structure
    const DWORD dataSize = detail::SafeCastSizeToDword(multiString.size() * sizeof(wchar_t));

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_MULTI_SZ,
        reinterpret_cast<const BYTE*>(multiString.data()),
        dataSize
    ) };
}


#ifndef GIOVANNI_DICANIO_WINREG_USE_CPP17

//
// C++20 code
//

inline RegResult RegKey::TrySetBinaryValue(const std::wstring& valueName,
                                           const std::span<const BYTE> data)
{
    _ASSERTE(IsValid());

    // Total data size, in bytes
    const DWORD dataSize = detail::SafeCastSizeToDword(data.size());

    return TrySetBinaryValue(valueName, data.data(), dataSize);
}


inline RegResult RegKey::TrySetBinaryValue(const std::wstring& valueName,
                                           const std::vector<BYTE>& data)
{
    _ASSERTE(IsValid());

    return TrySetBinaryValue(valueName, std::span{ data });
}

#else

//
// C++17-compatible code
//

inline RegResult RegKey::TrySetBinaryValue(const std::wstring& valueName,
                                           const std::vector<BYTE>& data)
{
    _ASSERTE(IsValid());

    // Total data size, in bytes
    const DWORD dataSize = detail::SafeCastSizeToDword(data.size());

    return RegResult{ ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_BINARY,
        data.data(),
        dataSize
    ) };
}

#endif // GIOVANNI_DICANIO_WINREG_USE_CPP17


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

    //
    // Read a DWORD (32-bit) value from the registry, invoking RegQueryValueEx
    //

    DWORD data = 0;                  // data to be read from the registry
    DWORD dataSize = sizeof(data);   // size of data, in bytes
    DWORD type = REG_NONE;           // type of the data stored in the registry

    const LSTATUS retCode = ::RegQueryValueExW(
        m_hKey,
        valueName.c_str(),
        nullptr,                        // reserved
        &type,
        reinterpret_cast<BYTE*>(&data),
        &dataSize
    );


    //
    // Handle error cases other than ERROR_MORE_DATA.
    //
    // ERROR_MORE_DATA means that the data stored in the registry is larger than a DWORD:
    // either a value of another type (e.g. REG_SZ) or a malformed REG_DWORD.
    // We'll handle that later.
    //
    if ((retCode != ERROR_SUCCESS) && (retCode != ERROR_MORE_DATA))
    {
        throw RegException{ retCode,
            "Cannot get DWORD value: RegQueryValueExW failed." };
    }

    //
    // Check type
    //
    // Note: We arrive at this check either on ERROR_SUCCESS or on ERROR_MORE_DATA.
    // RegQueryValueExW still fills 'type' on ERROR_MORE_DATA (verified, not documented);
    // anyway, since 'type' is initialized to REG_NONE, we still fail safely if that ever changes.
    //
    if (type != REG_DWORD)
    {
        throw RegException{ ERROR_UNSUPPORTED_TYPE,
            "Cannot get DWORD value: type is not REG_DWORD." };
    }

    //
    // Check size
    //
    // ERROR_MORE_DATA: REG_DWORD larger than 4 bytes.
    //
    // dataSize < 4: on short data, only the stored bytes are copied, with no zero-extension
    // (verified), so the upper bytes of 'data' would not come from the registry. And, anyway,
    // a REG_DWORD value shorter than 4 bytes is simply invalid.
    //
    if ((retCode == ERROR_MORE_DATA) || (dataSize != sizeof(data)))
    {
        throw RegException{ ERROR_INVALID_DATA,
            "Cannot get DWORD value: REG_DWORD data is not 4 bytes." };
    }

    return data;
}


inline ULONGLONG RegKey::GetQwordValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    //
    // Read a QWORD (64-bit) value from the registry, invoking RegQueryValueEx
    //

    ULONGLONG data = 0;              // data to be read from the registry
    DWORD dataSize = sizeof(data);   // size of data, in bytes
    DWORD type = REG_NONE;           // type of the data stored in the registry

    const LSTATUS retCode = ::RegQueryValueExW(
        m_hKey,
        valueName.c_str(),
        nullptr,                        // reserved
        &type,
        reinterpret_cast<BYTE*>(&data),
        &dataSize
    );


    //
    // Handle error cases other than ERROR_MORE_DATA.
    //
    // ERROR_MORE_DATA means that the data stored in the registry is larger than a QWORD:
    // either a value of another type (e.g. REG_SZ) or a malformed REG_QWORD.
    // We'll handle that later.
    //
    if ((retCode != ERROR_SUCCESS) && (retCode != ERROR_MORE_DATA))
    {
        throw RegException{ retCode,
            "Cannot get QWORD value: RegQueryValueExW failed." };
    }

    //
    // Check type
    //
    // Note: We arrive at this check either on ERROR_SUCCESS or on ERROR_MORE_DATA.
    // RegQueryValueExW still fills 'type' on ERROR_MORE_DATA (verified, not documented);
    // anyway, since 'type' is initialized to REG_NONE, we still fail safely if that ever changes.
    //
    if (type != REG_QWORD)
    {
        throw RegException{ ERROR_UNSUPPORTED_TYPE,
            "Cannot get QWORD value: type is not REG_QWORD." };
    }

    //
    // Check size
    //
    // ERROR_MORE_DATA: REG_QWORD larger than 8 bytes.
    //
    // dataSize < 8: on short data, only the stored bytes are copied, with no zero-extension
    // (verified), so the upper bytes of 'data' would not come from the registry. And, anyway,
    // a REG_QWORD value shorter than 8 bytes is simply invalid.
    //
    if ((retCode == ERROR_MORE_DATA) || (dataSize != sizeof(data)))
    {
        throw RegException{ ERROR_INVALID_DATA,
            "Cannot get QWORD value: REG_QWORD data is not 8 bytes." };
    }

    return data;
}


inline RegExpected<DWORD> RegKey::TryGetDwordValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());


    //
    // Read a DWORD (32-bit) value from the registry, invoking RegQueryValueEx
    //

    using RegValueType = DWORD;

    DWORD data = 0;                  // data to be read from the registry
    DWORD dataSize = sizeof(data);   // size of data, in bytes
    DWORD type = REG_NONE;           // type of the data stored in the registry

    const LSTATUS retCode = ::RegQueryValueExW(
        m_hKey,
        valueName.c_str(),
        nullptr,                        // reserved
        &type,
        reinterpret_cast<BYTE*>(&data),
        &dataSize
    );


    //
    // Handle error cases other than ERROR_MORE_DATA.
    //
    // ERROR_MORE_DATA means that the data stored in the registry is larger than a DWORD:
    // either a value of another type (e.g. REG_SZ) or a malformed REG_DWORD.
    // We'll handle that later.
    //
    if ((retCode != ERROR_SUCCESS) && (retCode != ERROR_MORE_DATA))
    {
        return RegExpected<RegValueType>::MakeRegExpectedWithError(retCode);
    }

    //
    // Check type
    //
    // Note: We arrive at this check either on ERROR_SUCCESS or on ERROR_MORE_DATA.
    // RegQueryValueExW still fills 'type' on ERROR_MORE_DATA (verified, not documented);
    // anyway, since 'type' is initialized to REG_NONE, we still fail safely if that ever changes.
    //
    if (type != REG_DWORD)
    {
        return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_UNSUPPORTED_TYPE);
    }

    //
    // Check size
    //
    // ERROR_MORE_DATA: REG_DWORD larger than 4 bytes.
    //
    // dataSize < 4: on short data, only the stored bytes are copied, with no zero-extension
    // (verified), so the upper bytes of 'data' would not come from the registry. And, anyway,
    // a REG_DWORD value shorter than 4 bytes is simply invalid.
    //
    if ((retCode == ERROR_MORE_DATA) || (dataSize != sizeof(data)))
    {
        return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_INVALID_DATA);
    }

    return RegExpected<RegValueType>::MakeSuccess(data);
}


inline RegExpected<ULONGLONG> RegKey::TryGetQwordValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    //
    // Read a QWORD (64-bit) value from the registry, invoking RegQueryValueEx
    //

    using RegValueType = ULONGLONG;

    ULONGLONG data = 0;              // data to be read from the registry
    DWORD dataSize = sizeof(data);   // size of data, in bytes
    DWORD type = REG_NONE;           // type of the data stored in the registry

    const LSTATUS retCode = ::RegQueryValueExW(
        m_hKey,
        valueName.c_str(),
        nullptr,                        // reserved
        &type,
        reinterpret_cast<BYTE*>(&data),
        &dataSize
    );


    //
    // Handle error cases other than ERROR_MORE_DATA.
    //
    // ERROR_MORE_DATA means that the data stored in the registry is larger than a QWORD:
    // either a value of another type (e.g. REG_SZ) or a malformed REG_QWORD.
    // We'll handle that later.
    //
    if ((retCode != ERROR_SUCCESS) && (retCode != ERROR_MORE_DATA))
    {
        return RegExpected<RegValueType>::MakeRegExpectedWithError(retCode);
    }

    //
    // Check type
    //
    // Note: We arrive at this check either on ERROR_SUCCESS or on ERROR_MORE_DATA.
    // RegQueryValueExW still fills 'type' on ERROR_MORE_DATA (verified, not documented);
    // anyway, since 'type' is initialized to REG_NONE, we still fail safely if that ever changes.
    //
    if (type != REG_QWORD)
    {
        return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_UNSUPPORTED_TYPE);
    }

    //
    // Check size
    //
    // ERROR_MORE_DATA: REG_QWORD larger than 8 bytes.
    //
    // dataSize < 8: on short data, only the stored bytes are copied, with no zero-extension
    // (verified), so the upper bytes of 'data' would not come from the registry. And, anyway,
    // a REG_QWORD value shorter than 8 bytes is simply invalid.
    //
    if ((retCode == ERROR_MORE_DATA) || (dataSize != sizeof(data)))
    {
        return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_INVALID_DATA);
    }

    return RegExpected<RegValueType>::MakeSuccess(data);
}


inline std::wstring RegKey::GetStringValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    std::wstring result;        // string to be read from the registry
    DWORD cbDataSize = 0;       // size of the string data, in bytes
    DWORD type = REG_NONE;      // type of the value stored in the registry


    LSTATUS retCode = ERROR_MORE_DATA;

    for (int attempt = 0; (attempt < detail::kMaxReadAttempts) && (retCode == ERROR_MORE_DATA); ++attempt)
    {
        //
        // Query the size of the result string
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,    // reserved
            &type,
            nullptr,    // data not required in this step; just need to get the data *size*
            &cbDataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            throw RegException{ retCode,
                "Cannot get the size of the string value: RegQueryValueExW failed." };
        }


        //
        // Check that type is REG_SZ
        //

        if (type != REG_SZ)
        {
            // Invalid type
            throw RegException{ ERROR_UNSUPPORTED_TYPE, "Cannot get string value: type is not REG_SZ." };
        }


        //
        // Special case of empty string
        //
        // If an empty value (size = 0 bytes) is found, return an empty string.
        //

        if (cbDataSize == 0)
        {
            return {};
        }

        //
        // Check the size: a UTF-16 string can't have an odd number of bytes
        //

        if (cbDataSize % sizeof(wchar_t) != 0)
        {
            throw RegException{ ERROR_INVALID_DATA,
                "Cannot get string value: odd byte length (not a valid UTF-16 string)." };
        }


        //
        // Resize the destination buffer
        //
        // Size is an even number of bytes (checked above); allocate exactly the stored data
        //

        result.resize(cbDataSize / sizeof(wchar_t));

        // Pass exactly the buffer size we own: result.size() characters (wchar_ts), expressed in bytes
        cbDataSize = detail::SafeCastSizeToDword(result.size() * sizeof(wchar_t));


        //
        // Read the string data in the destination buffer
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,                                    // reserved
            &type,
            reinterpret_cast<BYTE*>(result.data()),     // pointer to the data buffer
            &cbDataSize
        );


        // Safe re-checks: The value may have changed between the size query and the read.
        if (retCode == ERROR_SUCCESS)
        {
            if (type != REG_SZ)
            {
                // Invalid type
                throw RegException{ ERROR_UNSUPPORTED_TYPE, "Cannot get string value: type is not REG_SZ." };
            }

            if (cbDataSize % sizeof(wchar_t) != 0)
            {
                throw RegException{ ERROR_INVALID_DATA,
                    "Cannot get string value: odd byte length (not a valid UTF-16 string)." };
            }
        }
    }

    if (retCode == ERROR_MORE_DATA)
    {
        throw RegException{ retCode, "Cannot get the string value: it kept growing while being read." };
    }

    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot get the string value: RegQueryValueExW failed." };
    }

    // Keep only what RegQueryValueExW actually wrote
    // (an even number of bytes, checked above)
    result.resize(cbDataSize / sizeof(wchar_t));

    // Then stop at the first NUL, which drops the terminator and any additional padding NULs
    result.resize(wcsnlen(result.data(), result.size()));

    return result;
}


inline RegExpected<std::wstring> RegKey::TryGetStringValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    using RegValueType = std::wstring;

    std::wstring result;        // string to be read from the registry
    DWORD cbDataSize = 0;       // size of the string data, in bytes
    DWORD type = REG_NONE;      // type of the value stored in the registry


    LSTATUS retCode = ERROR_MORE_DATA;

    for (int attempt = 0; (attempt < detail::kMaxReadAttempts) && (retCode == ERROR_MORE_DATA); ++attempt)
    {
        //
        // Query the size of the result string
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,    // reserved
            &type,
            nullptr,    // data not required in this step; just need to get the data *size*
            &cbDataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            return RegExpected<RegValueType>::MakeRegExpectedWithError(retCode);
        }


        //
        // Check that type is REG_SZ
        //

        if (type != REG_SZ)
        {
            // Invalid type
            return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_UNSUPPORTED_TYPE);
        }


        //
        // Special case of empty string
        //
        // If an empty value (size = 0 bytes) is found, return an empty string.
        //

        if (cbDataSize == 0)
        {
            return RegExpected<RegValueType>::MakeSuccess(std::wstring{});
        }

        //
        // Check the size: a UTF-16 string can't have an odd number of bytes
        //

        if (cbDataSize % sizeof(wchar_t) != 0)
        {
            return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_INVALID_DATA);
        }


        //
        // Resize the destination buffer
        //
        // Size is an even number of bytes (checked above); allocate exactly the stored data
        //

        result.resize(cbDataSize / sizeof(wchar_t));

        // Pass exactly the buffer size we own: result.size() characters (wchar_ts), expressed in bytes
        // (Note: Cannot overflow this case, as it equals the even size returned by the API.)
        cbDataSize = detail::SafeCastSizeToDword(result.size() * sizeof(wchar_t));


        //
        // Read the string data in the destination buffer
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,                                    // reserved
            &type,
            reinterpret_cast<BYTE*>(result.data()),     // pointer to the data buffer
            &cbDataSize
        );


        // Safe re-checks: The value may have changed between the size query and the read.
        if (retCode == ERROR_SUCCESS)
        {
            if (type != REG_SZ)
            {
                // Invalid type
                return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_UNSUPPORTED_TYPE);
            }

            if (cbDataSize % sizeof(wchar_t) != 0)
            {
                return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_INVALID_DATA);
            }
        }
    }

    // Includes ERROR_MORE_DATA if the value kept growing for all attempts
    if (retCode != ERROR_SUCCESS)
    {
        return RegExpected<RegValueType>::MakeRegExpectedWithError(retCode);
    }

    // Keep only what RegQueryValueExW actually wrote
    // (an even number of bytes, checked above)
    result.resize(cbDataSize / sizeof(wchar_t));

    // Then stop at the first NUL, which drops the terminator and any additional padding NULs
    result.resize(wcsnlen(result.data(), result.size()));

    return RegExpected<RegValueType>::MakeSuccess(std::move(result));
}


inline std::wstring RegKey::GetExpandStringValue(const std::wstring& valueName,
                                                 const ExpandStringOption expandOption) const
{
    _ASSERTE(IsValid());

    std::wstring result;        // expand string to be read from the registry
    DWORD cbDataSize = 0;       // size of the string data, in bytes
    DWORD type = REG_NONE;      // type of the value stored in the registry


    LSTATUS retCode = ERROR_MORE_DATA;

    for (int attempt = 0; (attempt < detail::kMaxReadAttempts) && (retCode == ERROR_MORE_DATA); ++attempt)
    {
        //
        // Query the size of the result string
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,    // reserved
            &type,
            nullptr,    // data not required in this step; just need to get the data *size*
            &cbDataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            throw RegException{ retCode,
                "Cannot get the size of the expand string value: RegQueryValueExW failed." };
        }


        //
        // Check that type is REG_EXPAND_SZ
        //

        if (type != REG_EXPAND_SZ)
        {
            // Invalid type
            throw RegException{ ERROR_UNSUPPORTED_TYPE,
                "Cannot get expand string value: type is not REG_EXPAND_SZ." };
        }


        //
        // Special case of empty string
        //
        // If an empty value (size = 0 bytes) is found, return an empty string.
        //

        if (cbDataSize == 0)
        {
            return {};
        }

        //
        // Check the size: a UTF-16 string can't have an odd number of bytes
        //

        if (cbDataSize % sizeof(wchar_t) != 0)
        {
            throw RegException{ ERROR_INVALID_DATA,
                "Cannot get expand string value: odd byte length (not a valid UTF-16 string)." };
        }


        //
        // Resize the destination buffer
        //
        // Size is an even number of bytes (checked above); allocate exactly the stored data
        //

        result.resize(cbDataSize / sizeof(wchar_t));

        // Pass exactly the buffer size we own: result.size() characters (wchar_ts), expressed in bytes
        cbDataSize = detail::SafeCastSizeToDword(result.size() * sizeof(wchar_t));


        //
        // Read the expand string data in the destination buffer
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,                                    // reserved
            &type,
            reinterpret_cast<BYTE*>(result.data()),     // pointer to the data buffer
            &cbDataSize
        );


        // Safe re-checks: The value may have changed between the size query and the read.
        if (retCode == ERROR_SUCCESS)
        {
            if (type != REG_EXPAND_SZ)
            {
                // Invalid type
                throw RegException{ ERROR_UNSUPPORTED_TYPE,
                    "Cannot get expand string value: type is not REG_EXPAND_SZ." };
            }

            if (cbDataSize % sizeof(wchar_t) != 0)
            {
                throw RegException{ ERROR_INVALID_DATA,
                    "Cannot get expand string value: odd byte length (not a valid UTF-16 string)." };
            }
        }
    }

    if (retCode == ERROR_MORE_DATA)
    {
        throw RegException{ retCode, "Cannot get the expand string value: it kept growing while being read." };
    }

    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "Cannot get the expand string value: RegQueryValueExW failed." };
    }

    // Keep only what RegQueryValueExW actually wrote
    // (an even number of bytes, checked above)
    result.resize(cbDataSize / sizeof(wchar_t));

    // Then stop at the first NUL, which drops the terminator and any additional padding NULs
    result.resize(wcsnlen(result.data(), result.size()));

    // Expand the string, if required
    if (expandOption == ExpandStringOption::Expand)
    {
        result = detail::ExpandEnvironmentStringsToWstring(result);
    }

    return result;
}


inline RegExpected<std::wstring> RegKey::TryGetExpandStringValue(const std::wstring& valueName,
                                                                 const ExpandStringOption expandOption) const
{
    _ASSERTE(IsValid());

    using RegValueType = std::wstring;

    std::wstring result;        // expand string to be read from the registry
    DWORD cbDataSize = 0;       // size of the string data, in bytes
    DWORD type = REG_NONE;      // type of the value stored in the registry


    LSTATUS retCode = ERROR_MORE_DATA;

    for (int attempt = 0; (attempt < detail::kMaxReadAttempts) && (retCode == ERROR_MORE_DATA); ++attempt)
    {
        //
        // Query the size of the result string
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,    // reserved
            &type,
            nullptr,    // data not required in this step; just need to get the data *size*
            &cbDataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            return RegExpected<RegValueType>::MakeRegExpectedWithError(retCode);
        }


        //
        // Check that type is REG_EXPAND_SZ
        //

        if (type != REG_EXPAND_SZ)
        {
            // Invalid type
            return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_UNSUPPORTED_TYPE);
        }


        //
        // Special case of empty string
        //
        // If an empty value (size = 0 bytes) is found, return an empty string.
        //

        if (cbDataSize == 0)
        {
            return RegExpected<RegValueType>::MakeSuccess(std::wstring{});
        }

        //
        // Check the size: a UTF-16 string can't have an odd number of bytes
        //

        if (cbDataSize % sizeof(wchar_t) != 0)
        {
            return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_INVALID_DATA);
        }


        //
        // Resize the destination buffer
        //
        // Size is an even number of bytes (checked above); allocate exactly the stored data
        //

        result.resize(cbDataSize / sizeof(wchar_t));

        // Pass exactly the buffer size we own: result.size() characters (wchar_ts), expressed in bytes
        cbDataSize = detail::SafeCastSizeToDword(result.size() * sizeof(wchar_t));


        //
        // Read the expand string data in the destination buffer
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,                                    // reserved
            &type,
            reinterpret_cast<BYTE*>(result.data()),     // pointer to the data buffer
            &cbDataSize
        );


        // Safe re-checks: The value may have changed between the size query and the read.
        if (retCode == ERROR_SUCCESS)
        {
            if (type != REG_EXPAND_SZ)
            {
                // Invalid type
                return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_UNSUPPORTED_TYPE);
            }

            if (cbDataSize % sizeof(wchar_t) != 0)
            {
                return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_INVALID_DATA);
            }
        }
    }

    // Includes ERROR_MORE_DATA if the value kept growing for all attempts
    if (retCode != ERROR_SUCCESS)
    {
        return RegExpected<RegValueType>::MakeRegExpectedWithError(retCode);
    }

    // Keep only what RegQueryValueExW actually wrote
    // (an even number of bytes, checked above)
    result.resize(cbDataSize / sizeof(wchar_t));

    // Then stop at the first NUL, which drops the terminator and any additional padding NULs
    result.resize(wcsnlen(result.data(), result.size()));

    // Expand the string, if required
    if (expandOption == ExpandStringOption::Expand)
    {
        return detail::TryExpandEnvironmentStringsToWstring(result);
    }

    return RegExpected<RegValueType>::MakeSuccess(std::move(result));
}


inline std::vector<std::wstring> RegKey::GetMultiStringValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    // Room for the result multi-string, to be read from the registry
    std::vector<wchar_t> multiString;

    // Size of the multi-string, in bytes
    DWORD cbDataSize = 0;

    // Type of the value stored in the registry
    DWORD type = REG_NONE;


    LSTATUS retCode = ERROR_MORE_DATA;

    for (int attempt = 0; (attempt < detail::kMaxReadAttempts) && (retCode == ERROR_MORE_DATA); ++attempt)
    {
        //
        // Query the size of the multi-string
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,    // reserved
            &type,
            nullptr,    // data not required in this step; just need to get the data *size*
            &cbDataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            throw RegException{ retCode,
                "Cannot get the size of the multi-string value: RegQueryValueExW failed." };
        }


        //
        // Check that type is REG_MULTI_SZ
        //

        if (type != REG_MULTI_SZ)
        {
            // Invalid type
            throw RegException{ ERROR_UNSUPPORTED_TYPE,
                "Cannot get multi-string value: type is not REG_MULTI_SZ." };
        }


        //
        // Special case of empty multi-string
        //
        // If an empty value (size = 0 bytes) is found, return an empty vector.
        //

        if (cbDataSize == 0)
        {
            return {};
        }


        //
        // Check the size: a sequence of UTF-16 strings can't have an odd number of bytes,
        // as the code unit is wchar_t, which is 2 bytes.
        //

        if (cbDataSize % sizeof(wchar_t) != 0)
        {
            throw RegException{ ERROR_INVALID_DATA,
                "Cannot get multi-string value: odd byte length (not valid UTF-16 data)." };
        }


        //
        // Resize the destination buffer
        //
        // Size is an even number of bytes (checked above); make also room for up to two additional NULs
        // that may be missing from the registry data.
        //
        const size_t cchBuffer = (cbDataSize / sizeof(wchar_t)) + 2;
        multiString.resize(cchBuffer);

        // Pass exactly the buffer size we own: multiString.size() wchar_ts, expressed in bytes
        cbDataSize = detail::SafeCastSizeToDword(multiString.size() * sizeof(wchar_t));


        //
        // Read the multi-string data in the destination buffer
        //

        // Read the multi-string data
        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,            // reserved
            &type,
            reinterpret_cast<BYTE*>(multiString.data()),     // pointer to the data buffer
            &cbDataSize
        );

        // Safe re-checks: Check the value type and size are still valid between the two calls
        if (retCode == ERROR_SUCCESS)
        {
            if (type != REG_MULTI_SZ)
            {
                // Invalid type
                throw RegException{ ERROR_UNSUPPORTED_TYPE,
                    "Cannot get multi-string value: type is not REG_MULTI_SZ." };
            }

            if ((cbDataSize % sizeof(wchar_t)) != 0)
            {
                throw RegException{ ERROR_INVALID_DATA,
                    "Cannot get multi-string value: odd byte length (not valid UTF-16 data)." };
            }
        }
    }

    if (retCode == ERROR_MORE_DATA)
    {
        throw RegException{ retCode,
            "Cannot get the multi-string value: it kept growing while being read." };
    }

    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode,
            "Cannot get the multi-string value: RegQueryValueExW failed." };
    }

    // Keep only what RegQueryValueExW actually wrote
    // (an even number of bytes, checked above)
    multiString.resize(cbDataSize / sizeof(wchar_t));

    // Add any terminating NULs missing from the registry data
    if (!multiString.empty())
    {
        if (multiString.back() != L'\0')
        {
            multiString.push_back(L'\0');       // terminate the last string
        }

        if ((multiString.size() >= 2) && (multiString[multiString.size() - 2] != L'\0'))
        {
            multiString.push_back(L'\0');       // terminate the whole sequence
        }
    }

    // Convert the double-null-terminated string structure to a vector<wstring>,
    // and return that back to the caller
    return detail::ParseMultiString(multiString);
}


inline RegExpected<std::vector<std::wstring>>
RegKey::TryGetMultiStringValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    using RegValueType = std::vector<std::wstring>;

    // Room for the result multi-string, to be read from the registry
    std::vector<wchar_t> multiString;

    // Size of the multi-string, in bytes
    DWORD cbDataSize = 0;

    // Type of the value stored in the registry
    DWORD type = REG_NONE;


    LSTATUS retCode = ERROR_MORE_DATA;

    for (int attempt = 0; (attempt < detail::kMaxReadAttempts) && (retCode == ERROR_MORE_DATA); ++attempt)
    {
        //
        // Query the size of the multi-string
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,    // reserved
            &type,
            nullptr,    // data not required in this step; just need to get the data *size*
            &cbDataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            return RegExpected<RegValueType>::MakeRegExpectedWithError(retCode);
        }


        //
        // Check that type is REG_MULTI_SZ
        //

        if (type != REG_MULTI_SZ)
        {
            // Invalid type
            return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_UNSUPPORTED_TYPE);
        }


        //
        // Special case of empty multi-string
        //
        // If an empty value (size = 0 bytes) is found, return an empty vector.
        //

        if (cbDataSize == 0)
        {
            return RegExpected<RegValueType>::MakeSuccess(std::vector<std::wstring>{});
        }


        //
        // Check the size: a sequence of UTF-16 strings can't have an odd number of bytes,
        // as the code unit is wchar_t, which is 2 bytes.
        //

        if (cbDataSize % sizeof(wchar_t) != 0)
        {
            return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_INVALID_DATA);
        }


        //
        // Resize the destination buffer
        //
        // Size is an even number of bytes (checked above); make also room for up to two additional NULs
        // that may be missing from the registry data.
        //
        const size_t cchBuffer = (cbDataSize / sizeof(wchar_t)) + 2;
        multiString.resize(cchBuffer);

        // Pass exactly the buffer size we own: multiString.size() wchar_ts, expressed in bytes.
        // NOTE: Overflow is only theoretical here, as a registry value close to 4 GB isn't realistic.
        // Registry hives are nowhere near 4 GB, and Microsoft recommends keeping values small.
        cbDataSize = detail::SafeCastSizeToDword(multiString.size() * sizeof(wchar_t));


        //
        // Read the multi-string data in the destination buffer
        //

        // Read the multi-string data
        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,            // reserved
            &type,
            reinterpret_cast<BYTE*>(multiString.data()),     // pointer to the data buffer
            &cbDataSize
        );

        // Safe re-checks: Check the value type and size are still valid between the two calls
        if (retCode == ERROR_SUCCESS)
        {
            if (type != REG_MULTI_SZ)
            {
                // Invalid type
                return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_UNSUPPORTED_TYPE);
            }

            if ((cbDataSize % sizeof(wchar_t)) != 0)
            {
                return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_INVALID_DATA);
            }
        }
    }

    if (retCode != ERROR_SUCCESS)   // includes ERROR_MORE_DATA
    {
        return RegExpected<RegValueType>::MakeRegExpectedWithError(retCode);
    }

    // Keep only what RegQueryValueExW actually wrote
    // (an even number of bytes, checked above)
    multiString.resize(cbDataSize / sizeof(wchar_t));

    // Add any terminating NULs missing from the registry data
    if (!multiString.empty())
    {
        if (multiString.back() != L'\0')
        {
            multiString.push_back(L'\0');       // terminate the last string
        }

        if ((multiString.size() >= 2) && (multiString[multiString.size() - 2] != L'\0'))
        {
            multiString.push_back(L'\0');       // terminate the whole sequence
        }
    }

    // Convert the double-null-terminated string structure to a vector<wstring>,
    // and return that back to the caller
    return detail::TryParseMultiString(multiString);
}


inline std::vector<BYTE> RegKey::GetBinaryValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    // Room for the binary data, to be read from the registry
    std::vector<BYTE> binaryData;

    // Size of binary data, in bytes
    DWORD cbDataSize = 0;

    // Type of the value stored in the registry
    DWORD type = REG_NONE;

    LSTATUS retCode = ERROR_MORE_DATA;

    for (int attempt = 0; (attempt < detail::kMaxReadAttempts) && (retCode == ERROR_MORE_DATA); ++attempt)
    {
        //
        // Query the size of the binary data
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,    // reserved
            &type,
            nullptr,    // data not required in this step; just need to get the data *size*
            &cbDataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            throw RegException{ retCode,
                "Cannot get the size of the binary data: RegQueryValueExW failed." };
        }


        //
        // Check that type is REG_BINARY
        //

        if (type != REG_BINARY)
        {
            // Invalid type
            throw RegException{ ERROR_UNSUPPORTED_TYPE,
                "Cannot get binary data: type is not REG_BINARY." };
        }

        //
        // Special case of empty binary data (0 bytes): return an empty vector.
        //
        // This is important because it also avoids calling RegQueryValueExW()
        // with binaryData.data() == nullptr, which would turn the read into
        // just another size query (when the data pointer is NULL).
        //
        if (cbDataSize == 0)
        {
            return {};
        }


        //
        // Allocate a buffer of proper size to store the binary data
        //

        binaryData.resize(cbDataSize);


        //
        // Call RegQueryValueEx for the second time to read the binary data content into the vector
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,                    // reserved
            &type,
            binaryData.data(),          // pointer to the data buffer
            &cbDataSize
        );

        // Safe type re-check: Check the value type is still valid between the two calls
        if ((retCode == ERROR_SUCCESS) && (type != REG_BINARY))
        {
            // Invalid type
            throw RegException{ ERROR_UNSUPPORTED_TYPE,
                "Cannot get binary data: type is not REG_BINARY." };
        }
    }

    if (retCode == ERROR_MORE_DATA)
    {
        throw RegException{ retCode,
            "Cannot get the binary data: it kept growing while being read." };
    }

    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode,
            "Cannot get the binary data: RegQueryValueExW failed." };
    }

    // Resize vector to the actual size returned by the last call to RegQueryValueEx;
    // this keeps only what the successful read wrote, so leftover data from earlier
    // attempts never leaks through.
    binaryData.resize(cbDataSize);

    return binaryData;
}


inline RegExpected<std::vector<BYTE>> RegKey::TryGetBinaryValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    using RegValueType = std::vector<BYTE>;

    // Room for the binary data, to be read from the registry
    std::vector<BYTE> binaryData;

    // Size of binary data, in bytes
    DWORD cbDataSize = 0;

    // Type of the value stored in the registry
    DWORD type = REG_NONE;

    LSTATUS retCode = ERROR_MORE_DATA;

    for (int attempt = 0; (attempt < detail::kMaxReadAttempts) && (retCode == ERROR_MORE_DATA); ++attempt)
    {
        //
        // Query the size of the binary data
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,    // reserved
            &type,
            nullptr,    // data not required in this step; just need to get the data *size*
            &cbDataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            return RegExpected<RegValueType>::MakeRegExpectedWithError(retCode);
        }


        //
        // Check that type is REG_BINARY
        //

        if (type != REG_BINARY)
        {
            // Invalid type
            return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_UNSUPPORTED_TYPE);
        }

        //
        // Special case of empty binary data (0 bytes): return an empty vector.
        //
        // This is important because it also avoids calling RegQueryValueExW()
        // with binaryData.data() == nullptr, which would turn the read into
        // just another size query (when the data pointer is NULL).
        //
        if (cbDataSize == 0)
        {
            return RegExpected<RegValueType>::MakeSuccess(std::vector<BYTE>{});
        }


        //
        // Allocate a buffer of proper size to store the binary data
        //

        binaryData.resize(cbDataSize);


        //
        // Call RegQueryValueEx for the second time to read the binary data content into the vector
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,                    // reserved
            &type,
            binaryData.data(),          // pointer to the data buffer
            &cbDataSize
        );

        // Safe type re-check: Check the value type is still valid between the two calls
        if ((retCode == ERROR_SUCCESS) && (type != REG_BINARY))
        {
            // Invalid type
            return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_UNSUPPORTED_TYPE);
        }
    }

    if (retCode != ERROR_SUCCESS)       // also includes ERROR_MORE_DATA
    {
        return RegExpected<RegValueType>::MakeRegExpectedWithError(retCode);
    }

    // Resize vector to the actual size returned by the last call to RegQueryValueEx;
    // this keeps only what the successful read wrote, so leftover data from earlier
    // attempts never leaks through.
    binaryData.resize(cbDataSize);

    return RegExpected<RegValueType>::MakeSuccess(std::move(binaryData));
}


inline RegKey::RawValue RegKey::GetRawValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    //
    // Note:
    // The algorithm used here doesn't work for HKEY_PERFORMANCE_DATA,
    // as in this case the size reported on ERROR_MORE_DATA is undefined,
    // and the correct approach would be growing the buffer continuously
    // until the call succeeds.
    //
    // TODO: I could implement it in some future release of the library...
    //
    if (m_hKey == HKEY_PERFORMANCE_DATA)
    {
        throw RegException{ ERROR_NOT_SUPPORTED,
            "RegKey::GetRawValue does not support HKEY_PERFORMANCE_DATA." };
    }

    // Room for the raw binary data, to be read from the registry
    std::vector<BYTE> binaryData;

    // Size of raw binary data, in bytes
    DWORD cbDataSize = 0;

    // Type of the value stored in the registry
    DWORD type = REG_NONE;

    LSTATUS retCode = ERROR_MORE_DATA;

    for (int attempt = 0; (attempt < detail::kMaxReadAttempts) && (retCode == ERROR_MORE_DATA); ++attempt)
    {
        //
        // Query the size of the raw binary data
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,    // reserved
            &type,
            nullptr,    // data not required in this step; just need to get the data *size*
            &cbDataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            throw RegException{ retCode,
                "Cannot get the size of the raw binary data: RegQueryValueExW failed." };
        }


        //
        // Note: No type check here!
        //
        // No type restriction, so this method can be used to read raw binary data
        // for unsupported or less documented types like REG_RESOURCE_LIST,
        // REG_FULL_RESOURCE_DESCRIPTOR, REG_LINK, etc.
        //

        //
        // Special case of empty binary data (0 bytes): return the value's type with empty data.
        //
        // This is important because it also avoids calling RegQueryValueExW()
        // with binaryData.data() == nullptr, which would turn the read into
        // just another size query (when the data pointer is NULL).
        //
        if (cbDataSize == 0)
        {
            return RawValue{ type, {} };
        }

        //
        // Allocate a buffer of proper size to store the binary data
        //

        binaryData.resize(cbDataSize);


        //
        // Call RegQueryValueEx for the second time to read the binary data content into the vector
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,                    // reserved
            &type,
            binaryData.data(),          // pointer to the data buffer
            &cbDataSize
        );

        //
        // No type check needed here, as we are just reading raw binary data
        //
    }

    if (retCode == ERROR_MORE_DATA)
    {
        throw RegException{ retCode,
            "Cannot get the raw binary data: it kept growing while being read." };
    }

    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode,
            "Cannot get the raw binary data: RegQueryValueExW failed." };
    }

    // Resize vector to the actual size returned by the last call to RegQueryValueEx;
    // this keeps only what the successful read wrote, so leftover data from earlier
    // attempts never leaks through.
    binaryData.resize(cbDataSize);

    return RawValue{ type, std::move(binaryData) };
}


inline RegExpected<RegKey::RawValue> RegKey::TryGetRawValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    using RegValueType = RegKey::RawValue;


    //
    // Note:
    // The algorithm used here doesn't work for HKEY_PERFORMANCE_DATA,
    // as in this case the size reported on ERROR_MORE_DATA is undefined,
    // and the correct approach would be growing the buffer continuously
    // until the call succeeds.
    //
    // TODO: I could implement it in some future release of the library...
    //
    if (m_hKey == HKEY_PERFORMANCE_DATA)
    {
        return RegExpected<RegValueType>::MakeRegExpectedWithError(ERROR_NOT_SUPPORTED);
    }

    // Room for the raw binary data, to be read from the registry
    std::vector<BYTE> binaryData;

    // Size of raw binary data, in bytes
    DWORD cbDataSize = 0;

    // Type of the value stored in the registry
    DWORD type = REG_NONE;

    LSTATUS retCode = ERROR_MORE_DATA;

    for (int attempt = 0; (attempt < detail::kMaxReadAttempts) && (retCode == ERROR_MORE_DATA); ++attempt)
    {
        //
        // Query the size of the raw binary data
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,    // reserved
            &type,
            nullptr,    // data not required in this step; just need to get the data *size*
            &cbDataSize
        );
        if (retCode != ERROR_SUCCESS)
        {
            return RegExpected<RegValueType>::MakeRegExpectedWithError(retCode);
        }


        //
        // Note: No type check here!
        //
        // No type restriction, so this method can be used to read raw binary data
        // for unsupported or less documented types like REG_RESOURCE_LIST,
        // REG_FULL_RESOURCE_DESCRIPTOR, REG_LINK, etc.
        //

        //
        // Special case of empty binary data (0 bytes): return the value's type with empty data.
        //
        // This is important because it also avoids calling RegQueryValueExW()
        // with binaryData.data() == nullptr, which would turn the read into
        // just another size query (when the data pointer is NULL).
        //
        if (cbDataSize == 0)
        {
            return RegExpected<RegValueType>::MakeSuccess(RawValue{ type, {} });
        }

        //
        // Allocate a buffer of proper size to store the binary data
        //

        binaryData.resize(cbDataSize);


        //
        // Call RegQueryValueEx for the second time to read the binary data content into the vector
        //

        retCode = ::RegQueryValueExW(
            m_hKey,
            valueName.c_str(),
            nullptr,                    // reserved
            &type,
            binaryData.data(),          // pointer to the data buffer
            &cbDataSize
        );

        //
        // No type check needed here, as we are just reading raw binary data
        //
    }

    if (retCode != ERROR_SUCCESS)       // includes ERROR_MORE_DATA
    {
        return RegExpected<RegValueType>::MakeRegExpectedWithError(retCode);
    }

    // Resize vector to the actual size returned by the last call to RegQueryValueEx;
    // this keeps only what the successful read wrote, so leftover data from earlier
    // attempts never leaks through.
    binaryData.resize(cbDataSize);

    return RegExpected<RegValueType>::MakeSuccess(RawValue{ type, std::move(binaryData) });
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
        return RegExpected<ReturnType>::MakeRegExpectedWithError(retCode);
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
            return RegExpected<ReturnType>::MakeRegExpectedWithError(retCode);
        }

        // On success, the ::RegEnumKeyEx API writes the length of the
        // subkey name in the subKeyNameLen output parameter
        // (not including the terminating NUL).
        // So I can build a wstring based on that length.
        subkeyNames.emplace_back(nameBuffer.get(), subKeyNameLen);
    }

    return RegExpected<ReturnType>::MakeSuccess( subkeyNames );
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
        return RegExpected<ReturnType>::MakeRegExpectedWithError(retCode);
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
            return RegExpected<ReturnType>::MakeRegExpectedWithError(retCode);
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

    return RegExpected<ReturnType>::MakeSuccess( valueInfo );
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
        return RegExpected<ReturnType>::MakeRegExpectedWithError(retCode);
    }

    return RegExpected<ReturnType>::MakeSuccess( typeId );
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
        return RegExpected<ReturnType>::MakeRegExpectedWithError(retCode);
    }

    return RegExpected<ReturnType>::MakeSuccess( infoKey );
}


inline RegKey::KeyReflection RegKey::QueryReflectionKey() const
{
    _ASSERTE(IsValid());

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
    _ASSERTE(IsValid());

    using ReturnType = RegKey::KeyReflection;

    BOOL isReflectionDisabled = FALSE;
    LSTATUS retCode = ::RegQueryReflectionKey(m_hKey, &isReflectionDisabled);
    if (retCode != ERROR_SUCCESS)
    {
        return RegExpected<ReturnType>::MakeRegExpectedWithError(retCode);
    }

    KeyReflection keyReflection = isReflectionDisabled ? KeyReflection::ReflectionDisabled
                                                       : KeyReflection::ReflectionEnabled;
    return RegExpected<ReturnType>::MakeSuccess( keyReflection );
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
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegLoadKeyW(m_hKey, subKey.c_str(), filename.c_str());
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegLoadKeyW failed." };
    }
}


inline RegResult RegKey::TryLoadKey(const std::wstring& subKey,
                                    const std::wstring& filename) noexcept
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegLoadKeyW(m_hKey, subKey.c_str(), filename.c_str()) };
}


inline void RegKey::SaveKey(const std::wstring& filename,
                            SECURITY_ATTRIBUTES* const securityAttributes) const
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegSaveKeyW(m_hKey, filename.c_str(), securityAttributes);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegSaveKeyW failed." };
    }
}


inline RegResult RegKey::TrySaveKey(const std::wstring& filename,
                                    SECURITY_ATTRIBUTES* const securityAttributes) const noexcept
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegSaveKeyW(m_hKey, filename.c_str(), securityAttributes) };
}


inline void RegKey::EnableReflectionKey()
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegEnableReflectionKey(m_hKey);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegEnableReflectionKey failed." };
    }
}


inline RegResult RegKey::TryEnableReflectionKey() noexcept
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegEnableReflectionKey(m_hKey) };
}


inline void RegKey::DisableReflectionKey()
{
    _ASSERTE(IsValid());

    LSTATUS retCode = ::RegDisableReflectionKey(m_hKey);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegDisableReflectionKey failed." };
    }
}


inline RegResult RegKey::TryDisableReflectionKey() noexcept
{
    _ASSERTE(IsValid());

    return RegResult{ ::RegDisableReflectionKey(m_hKey) };
}


inline bool RegKey::ContainsValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    //
    // Invoke RegQueryValueEx to just check if the input value exists under the current key;
    // no type, no data, no size needed.
    //

    const LSTATUS retCode = ::RegQueryValueExW(
        m_hKey,             // key
        valueName.c_str(),  // value name
        nullptr,            // reserved
        nullptr,            // type not required
        nullptr,            // data not required
        nullptr             // data size not required (allowed, since 'data' is nullptr)
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
        throw RegException{ retCode,
            "RegQueryValueExW failed when checking if the current key contains the specified value." };
    }
}


inline RegExpected<bool> RegKey::TryContainsValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    //
    // Invoke RegQueryValueEx to just check if the input value exists under the current key;
    // no type, no data, no size needed.
    //

    const LSTATUS retCode = ::RegQueryValueExW(
        m_hKey,             // key
        valueName.c_str(),  // value name
        nullptr,            // reserved
        nullptr,            // type not required
        nullptr,            // data not required
        nullptr             // data size not required (allowed, since 'data' is nullptr)
    );

    if (retCode == ERROR_SUCCESS)
    {
        // The value exists under the current key
        return RegExpected<bool>::MakeSuccess(true);
    }
    else if (retCode == ERROR_FILE_NOT_FOUND)
    {
        // The value does *not* exist under the current key
        return RegExpected<bool>::MakeSuccess(false);
    }
    else
    {
        // Some other error occurred - pass the returned error code back to the caller
        return RegExpected<bool>::MakeRegExpectedWithError(retCode);
    }
}


inline bool RegKey::ContainsSubKey(const std::wstring& subKey,
                                   const REGSAM wow64Flag /* = KEY_WOW64_64KEY */) const
{
    _ASSERTE(IsValid());

    // Only the WOW64 view bits are meaningful here
    _ASSERTE((wow64Flag & ~(KEY_WOW64_64KEY | KEY_WOW64_32KEY)) == 0);


    //
    // Let's try and open the specified subKey, then check the return code
    // of RegOpenKeyExW to figure out if the subKey exists or not.
    //

    HKEY hSubKey = nullptr;
    const REGSAM desiredAccess = KEY_QUERY_VALUE | wow64Flag;
    LSTATUS retCode = ::RegOpenKeyExW(
        m_hKey,
        subKey.c_str(),
        0,
        desiredAccess,
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
    //
    // NOTE: ERROR_ACCESS_DENIED
    //
    // There could be a special case for ERROR_ACCESS_DENIED:
    // The sub-key could exist, but the caller can't open it.
    //
    // That's because name resolution happens before the access check (ACL) on the final key,
    // so a non-existent sub-key would report ERROR_FILE_NOT_FOUND instead.
    // So, in this case I could still return true (= the sub-key exists) back to the caller.
    //
    // I read that an exception to this could be:
    // a token without SeChangeNotifyPrivilege could be denied on an intermediate key.
    //
    // So, even with ERROR_ACCESS_DENIED the sub-key almost certainly exists.
    //
    // Anyway, that requires more testing.
    //
    // TODO: More test on this special case of ERROR_ACCESS_DENIED.
    //
    else if (retCode == ERROR_ACCESS_DENIED)
    {
        return true;
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


inline RegExpected<bool> RegKey::TryContainsSubKey(const std::wstring& subKey,
                                                   REGSAM wow64Flag /* = KEY_WOW64_64KEY */) const
{
    _ASSERTE(IsValid());

    // Only the WOW64 view bits are meaningful here
    _ASSERTE((wow64Flag & ~(KEY_WOW64_64KEY | KEY_WOW64_32KEY)) == 0);


    //
    // Let's try and open the specified subKey, then check the return code
    // of RegOpenKeyExW to figure out if the subKey exists or not.
    //

    HKEY hSubKey = nullptr;

    const REGSAM desiredAccess = KEY_QUERY_VALUE | wow64Flag;

    LSTATUS retCode = ::RegOpenKeyExW(
        m_hKey,
        subKey.c_str(),
        0,
        desiredAccess,
        &hSubKey
    );
    if (retCode == ERROR_SUCCESS)
    {
        // We were able to open the specified sub-key, so the sub-key does exist.
        //
        // Don't forget to close the sub-key opened for this testing purpose!
        ::RegCloseKey(hSubKey);
        hSubKey = nullptr;

        return RegExpected<bool>::MakeSuccess(true);
    }
    else if ((retCode == ERROR_FILE_NOT_FOUND) || (retCode == ERROR_PATH_NOT_FOUND))
    {
        // The specified sub-key does not exist
        return RegExpected<bool>::MakeSuccess(false);
    }
    //
    // NOTE: ERROR_ACCESS_DENIED
    //
    // There could be a special case for ERROR_ACCESS_DENIED:
    // The sub-key could exist, but the caller can't open it.
    //
    // That's because name resolution happens before the access check (ACL) on the final key,
    // so a non-existent sub-key would report ERROR_FILE_NOT_FOUND instead.
    // So, in this case I could still return true (= the sub-key exists) back to the caller.
    //
    // I read that an exception to this could be:
    // a token without SeChangeNotifyPrivilege could be denied on an intermediate key.
    //
    // So, even with ERROR_ACCESS_DENIED the sub-key almost certainly exists.
    //
    // Anyway, that requires more testing.
    //
    // TODO: More test on this special case of ERROR_ACCESS_DENIED.
    //
    else if (retCode == ERROR_ACCESS_DENIED)
    {
        return RegExpected<bool>::MakeSuccess(true);
    }
    else
    {
        // Some other error occurred
        return RegExpected<bool>::MakeRegExpectedWithError(retCode);
    }
}


inline void RegKey::ConnectRegistry(const std::wstring& machineName, const HKEY hKeyPredefined)
{
    // Try the connection
    HKEY hKeyResult = nullptr;
    LSTATUS retCode = ::RegConnectRegistryW(machineName.c_str(), hKeyPredefined, &hKeyResult);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ retCode, "RegConnectRegistryW failed." };
    }

    // Safely close any previously opened key
    Close();

    // Take ownership of the result key
    m_hKey = hKeyResult;
}


inline RegResult RegKey::TryConnectRegistry(const std::wstring& machineName,
    const HKEY hKeyPredefined) noexcept
{
    // Try the connection
    HKEY hKeyResult = nullptr;
    RegResult retCode{ ::RegConnectRegistryW(machineName.c_str(), hKeyPredefined, &hKeyResult) };
    if (retCode.Failed())
    {
        return retCode;
    }

    // Safely close any previously opened key
    Close();

    // Take ownership of the result key
    m_hKey = hKeyResult;

    _ASSERTE(retCode.IsOk());
    return retCode;
}


inline const wchar_t* RegKey::RegTypeToString(const DWORD regType)
{
    switch (regType)
    {
        case REG_SZ:                return L"REG_SZ";
        case REG_EXPAND_SZ:         return L"REG_EXPAND_SZ";
        case REG_MULTI_SZ:          return L"REG_MULTI_SZ";
        case REG_DWORD:             return L"REG_DWORD";
        case REG_QWORD:             return L"REG_QWORD";
        case REG_BINARY:            return L"REG_BINARY";

        case REG_NONE:              return L"REG_NONE";
        case REG_LINK:              return L"REG_LINK";
        case REG_RESOURCE_LIST:     return L"REG_RESOURCE_LIST";
        case REG_DWORD_BIG_ENDIAN:  return L"REG_DWORD_BIG_ENDIAN";

        default:                    return L"Unknown/unsupported registry type";
    }
}


} // namespace winreg


#endif // GIOVANNI_DICANIO_WINREG_REGKEY_IMPL_HPP

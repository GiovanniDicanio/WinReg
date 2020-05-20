////////////////////////////////////////////////////////////////////////////////
//
// FILE: WinReg.cpp
// DESC: Implementation code of non-inline methods, and private helpers.
//
// By Giovanni Dicanio
//
// See WinReg.hpp and the LICENSE file for license information.
//
////////////////////////////////////////////////////////////////////////////////


//==============================================================================
//                              Includes
//==============================================================================

#include "WinReg.hpp"       // The library's public header

#include <memory>           // std::unique_ptr, std::make_unique



//==============================================================================
//              File-Private Helper Classes and Functions
//==============================================================================

namespace
{

//------------------------------------------------------------------------------
// Simple scoped-based RAII wrapper that *automatically* invokes ::LocalFree()
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
    T* m_ptr = nullptr;
};

} // namespace


//------------------------------------------------------------------------------
// Helper function to build a multi-string from a vector<wstring>.
//
// A multi-string is a sequence of contiguous NUL-terminated strings,
// that terminates with an additional NUL.
// Basically, considered as a whole, the sequence is terminated by two NULs.
// E.g.:
//          Hello\0World\0\0
//------------------------------------------------------------------------------
[[nodiscard]] static std::vector<wchar_t> BuildMultiString(
    const std::vector<std::wstring>& data
)
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
        multiString.insert(multiString.end(), s.begin(), s.end());

        // Don't forget to NUL-terminate the current string
        multiString.emplace_back(L'\0');
    }

    // Add the last NUL-terminator
    multiString.emplace_back(L'\0');

    return multiString;
}


//==============================================================================
//                      Non-inline Implementations
//==============================================================================

namespace winreg
{


void RegKey::SetMultiStringValue(
    const std::wstring& valueName,
    const std::vector<std::wstring>& data
)
{
    _ASSERTE(IsValid());

    // First, we have to build a double-NUL-terminated multi-string from the input data
    const std::vector<wchar_t> multiString = BuildMultiString(data);

    // Total size, in bytes, of the whole multi-string structure
    const DWORD dataSize = static_cast<DWORD>(multiString.size() * sizeof(wchar_t));

    LONG retCode = ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_MULTI_SZ,
        reinterpret_cast<const BYTE*>(&multiString[0]),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot write multi-string value: RegSetValueExW failed.");
    }
}


RegResult RegKey::TrySetMultiStringValue(
    const std::wstring& valueName,
    const std::vector<std::wstring>& data
) // Can't be noexcept, because the BuildMultiString() helper can throw,
  // e.g. on memory allocation failure!
{
    //
    // NOTE:
    // This method can't be marked noexcept, because it calls BuildMultiString(),
    // which allocates dynamic memory for the resulting std::vector<wchar_t>.
    //

    _ASSERTE(IsValid());

    // First, we have to build a double-NUL-terminated multi-string from the input data
    const std::vector<wchar_t> multiString = BuildMultiString(data);

    // Total size, in bytes, of the whole multi-string structure
    const DWORD dataSize = static_cast<DWORD>(multiString.size() * sizeof(wchar_t));

    return ::RegSetValueExW(
        m_hKey,
        valueName.c_str(),
        0, // reserved
        REG_MULTI_SZ,
        reinterpret_cast<const BYTE*>(&multiString[0]),
        dataSize
    );
}


std::wstring RegKey::GetStringValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    // Get the size of the result string
    DWORD dataSize = 0; // size of data, in bytes
    const DWORD flags = RRF_RT_REG_SZ;
    LONG retCode = ::RegGetValueW(
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
        throw RegException(retCode, "Cannot get size of string value: RegGetValueW failed.");
    }

    // Allocate a string of proper size.
    // Note that dataSize is in bytes and includes the terminating NUL;
    // we have to convert the size from bytes to wchar_ts for wstring::resize.
    std::wstring result(dataSize / sizeof(wchar_t), L' ');

    // Call RegGetValue for the second time to read the string's content
    retCode = ::RegGetValueW(
        m_hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        &result[0], // output buffer
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot get string value: RegGetValueW failed.");
    }

    // Remove the NUL terminator scribbled by RegGetValue from the wstring
    result.resize((dataSize / sizeof(wchar_t)) - 1);

    return result;
}


std::wstring RegKey::GetExpandStringValue(
    const std::wstring& valueName,
    const ExpandStringOption expandOption
) const
{
    _ASSERTE(IsValid());

    DWORD flags = RRF_RT_REG_EXPAND_SZ;

    // Adjust the flag for RegGetValue considering the expand string option specified by the caller
    if (expandOption == ExpandStringOption::DontExpand)
    {
        flags |= RRF_NOEXPAND;
    }

    // Get the size of the result string
    DWORD dataSize = 0; // size of data, in bytes
    LONG retCode = ::RegGetValueW(
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
        throw RegException(retCode, "Cannot get size of expand string value: RegGetValueW failed.");
    }

    // Allocate a string of proper size.
    // Note that dataSize is in bytes and includes the terminating NUL.
    // We must convert from bytes to wchar_ts for wstring::resize.
    std::wstring result(dataSize / sizeof(wchar_t), L' ');

    // Call RegGetValue for the second time to read the string's content
    retCode = ::RegGetValueW(
        m_hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        &result[0], // output buffer
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot get expand string value: RegGetValueW failed.");
    }

    // Remove the NUL terminator scribbled by RegGetValue from the wstring
    result.resize((dataSize / sizeof(wchar_t)) - 1);

    return result;
}


std::vector<std::wstring> RegKey::GetMultiStringValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    // Request the size of the multi-string, in bytes
    DWORD dataSize = 0;
    const DWORD flags = RRF_RT_REG_MULTI_SZ;
    LONG retCode = ::RegGetValueW(
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
        throw RegException(retCode, "Cannot get size of multi-string value: RegGetValueW failed.");
    }

    // Allocate room for the result multi-string.
    // Note that dataSize is in bytes, but our vector<wchar_t>::resize method requires size
    // to be expressed in wchar_ts.
    std::vector<wchar_t> data(dataSize / sizeof(wchar_t), L' ');

    // Read the multi-string from the registry into the vector object
    retCode = ::RegGetValueW(
        m_hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // no type required
        &data[0],   // output buffer
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot get multi-string value: RegGetValueW failed.");
    }

    // Resize vector to the actual size returned by GetRegValue.
    // Note that the vector is a vector of wchar_ts, instead the size returned by GetRegValue
    // is in bytes, so we have to scale from bytes to wchar_t count.
    data.resize(dataSize / sizeof(wchar_t));

    // Parse the double-NUL-terminated string into a vector<wstring>,
    // which will be returned to the caller
    std::vector<std::wstring> result;
    const wchar_t* currStringPtr = &data[0];
    while (*currStringPtr != L'\0')
    {
        // Current string is NUL-terminated, so get its length calling wcslen
        const size_t currStringLength = wcslen(currStringPtr);

        // Add current string to the result vector
        result.emplace_back(currStringPtr, currStringLength);

        // Move to the next string
        currStringPtr += currStringLength + 1;
    }

    return result;
}


std::vector<BYTE> RegKey::GetBinaryValue(const std::wstring& valueName) const
{
    _ASSERTE(IsValid());

    // Get the size of the binary data
    DWORD dataSize = 0; // size of data, in bytes
    const DWORD flags = RRF_RT_REG_BINARY;
    LONG retCode = ::RegGetValueW(
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
        throw RegException(retCode, "Cannot get size of binary data: RegGetValueW failed.");
    }

    // Allocate a buffer of proper size to store the binary data
    std::vector<BYTE> data(dataSize);

    // Call RegGetValue for the second time to read the data content
    retCode = ::RegGetValueW(
        m_hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        &data[0],   // output buffer
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot get binary data: RegGetValueW failed.");
    }

    return data;
}


RegResult RegKey::TryGetStringValue(
    const std::wstring& valueName,
    std::wstring& result
) const
{
    _ASSERTE(IsValid());

    result.clear();

    // Get the size of the result string
    DWORD dataSize = 0; // size of data, in bytes
    const DWORD flags = RRF_RT_REG_SZ;
    RegResult retCode = ::RegGetValueW(
        m_hKey,
        nullptr, // no subkey
        valueName.c_str(),
        flags,
        nullptr, // type not required
        nullptr, // output buffer not needed now
        &dataSize
    );
    if (retCode.Failed())
    {
        _ASSERTE(result.empty());
        return retCode;
    }

    // Allocate a string of proper size.
    // Note that dataSize is in bytes and includes the terminating NUL;
    // we have to convert the size from bytes to wchar_ts for wstring::resize.
    result.resize(dataSize / sizeof(wchar_t));

    // Call RegGetValue for the second time to read the string's content
    retCode = ::RegGetValueW(
        m_hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        &result[0], // output buffer
        &dataSize
    );
    if (retCode.Failed())
    {
        result.clear();
        return retCode;
    }

    // Remove the NUL terminator scribbled by RegGetValue from the wstring
    result.resize((dataSize / sizeof(wchar_t)) - 1);

    _ASSERTE(retCode.IsOk());
    return retCode;
}


RegResult RegKey::TryGetExpandStringValue(
    const std::wstring& valueName,
    std::wstring& result,
    const ExpandStringOption expandOption
) const
{
    _ASSERTE(IsValid());

    result.clear();

    DWORD flags = RRF_RT_REG_EXPAND_SZ;

    // Adjust the flag for RegGetValue considering the expand string option specified by the caller
    if (expandOption == ExpandStringOption::DontExpand)
    {
        flags |= RRF_NOEXPAND;
    }

    // Get the size of the result string
    DWORD dataSize = 0; // size of data, in bytes
    RegResult retCode = ::RegGetValueW(
        m_hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        nullptr,    // output buffer not needed now
        &dataSize
    );
    if (retCode.Failed())
    {
        _ASSERTE(result.empty());
        return retCode;
    }

    // Allocate a string of proper size.
    // Note that dataSize is in bytes and includes the terminating NUL.
    // We must convert from bytes to wchar_ts for wstring::resize.
    result.resize(dataSize / sizeof(wchar_t));

    // Call RegGetValue for the second time to read the string's content
    retCode = ::RegGetValueW(
        m_hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        &result[0], // output buffer
        &dataSize
    );
    if (retCode.Failed())
    {
        result.clear();
        return retCode;
    }

    // Remove the NUL terminator scribbled by RegGetValue from the wstring
    result.resize((dataSize / sizeof(wchar_t)) - 1);

    _ASSERTE(retCode.IsOk());
    return retCode;
}


RegResult RegKey::TryGetMultiStringValue(
    const std::wstring& valueName,
    std::vector<std::wstring>& result
) const
{
    _ASSERTE(IsValid());

    result.clear();

    // Request the size of the multi-string, in bytes
    DWORD dataSize = 0;
    const DWORD flags = RRF_RT_REG_MULTI_SZ;
    RegResult retCode = ::RegGetValueW(
        m_hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        nullptr,    // output buffer not needed now
        &dataSize
    );
    if (retCode.Failed())
    {
        _ASSERTE(result.empty());
        return retCode;
    }

    // Allocate room for the result multi-string.
    // Note that dataSize is in bytes, but our vector<wchar_t>::resize method requires size
    // to be expressed in wchar_ts.
    std::vector<wchar_t> data(dataSize / sizeof(wchar_t));

    // Read the multi-string from the registry into the vector object
    retCode = ::RegGetValueW(
        m_hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // no type required
        &data[0],   // output buffer
        &dataSize
    );
    if (retCode.Failed())
    {
        _ASSERTE(result.empty());
        return retCode;
    }

    // Resize vector to the actual size returned by GetRegValue.
    // Note that the vector is a vector of wchar_ts, instead the size returned by GetRegValue
    // is in bytes, so we have to scale from bytes to wchar_t count.
    data.resize(dataSize / sizeof(wchar_t));


    // Parse the double-NUL-terminated string into a vector<wstring>,
    // which will be returned to the caller
    _ASSERTE(result.empty());
    const wchar_t* currStringPtr = &data[0];
    while (*currStringPtr != L'\0')
    {
        // Current string is NUL-terminated, so get its length calling wcslen
        const size_t currStringLength = wcslen(currStringPtr);

        // Add current string to the result vector
        result.emplace_back(currStringPtr, currStringLength);

        // Move to the next string
        currStringPtr += currStringLength + 1;
    }

    _ASSERTE(retCode.IsOk());
    return retCode;
}


RegResult RegKey::TryGetBinaryValue(
    const std::wstring& valueName,
    std::vector<BYTE>& result
) const
{
    _ASSERTE(IsValid());

    result.clear();

    // Get the size of the binary data
    DWORD dataSize = 0; // size of data, in bytes
    const DWORD flags = RRF_RT_REG_BINARY;
    RegResult retCode = ::RegGetValueW(
        m_hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        nullptr,    // output buffer not needed now
        &dataSize
    );
    if (retCode.Failed())
    {
        _ASSERTE(result.empty());
        return retCode;
    }

    // Allocate a buffer of proper size to store the binary data
    result.resize(dataSize);

    // Call RegGetValue for the second time to read the data content
    retCode = ::RegGetValueW(
        m_hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        &result[0], // output buffer
        &dataSize
    );
    if (retCode.Failed())
    {
        result.clear();
        return retCode;
    }

    _ASSERTE(retCode.IsOk());
    return retCode;
}


std::vector<std::wstring> RegKey::EnumSubKeys() const
{
    _ASSERTE(IsValid());

    // Get some useful enumeration info, like the total number of subkeys
    // and the maximum length of the subkey names
    DWORD subKeyCount = 0;
    DWORD maxSubKeyNameLen = 0;
    LONG retCode = ::RegQueryInfoKeyW(
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
        throw RegException(retCode,
            "RegQueryInfoKeyW failed while preparing for subkey enumeration."
        );
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
            throw RegException(retCode, "Cannot enumerate subkeys: RegEnumKeyExW failed.");
        }

        // On success, the ::RegEnumKeyEx API writes the length of the
        // subkey name in the subKeyNameLen output parameter
        // (not including the terminating NUL).
        // So I can build a wstring based on that length.
        subkeyNames.emplace_back(nameBuffer.get(), subKeyNameLen);
    }

    return subkeyNames;
}


std::vector<std::pair<std::wstring, DWORD>> RegKey::EnumValues() const
{
    _ASSERTE(IsValid());

    // Get useful enumeration info, like the total number of values
    // and the maximum length of the value names
    DWORD valueCount = 0;
    DWORD maxValueNameLen = 0;
    LONG retCode = ::RegQueryInfoKeyW(
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
        throw RegException(
            retCode,
            "RegQueryInfoKeyW failed while preparing for value enumeration."
        );
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
            throw RegException(retCode, "Cannot enumerate values: RegEnumValueW failed.");
        }

        // On success, the RegEnumValue API writes the length of the
        // value name in the valueNameLen output parameter
        // (not including the terminating NUL).
        // So we can build a wstring based on that.
        valueInfo.emplace_back(
            std::wstring(nameBuffer.get(), valueNameLen),
            valueType
        );
    }

    return valueInfo;
}


RegResult RegKey::TryEnumSubKeys(std::vector<std::wstring>& subKeys) const
{
    subKeys.clear();

    _ASSERTE(IsValid());

    // Get some useful enumeration info, like the total number of subkeys
    // and the maximum length of the subkey names
    DWORD subKeyCount = 0;
    DWORD maxSubKeyNameLen = 0;
    RegResult retCode = ::RegQueryInfoKeyW(
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
    if (retCode.Failed())
    {
        return retCode;
    }

    // NOTE: According to the MSDN documentation, the size returned for subkey name max length
    // does *not* include the terminating NUL, so let's add +1 to take it into account
    // when I allocate the buffer for reading subkey names.
    maxSubKeyNameLen++;

    // Preallocate a buffer for the subkey names
    auto nameBuffer = std::make_unique<wchar_t[]>(maxSubKeyNameLen);


    // Should have been cleared at the beginning of the method
    _ASSERTE(subKeys.empty());

    // Reserve room in the vector to speed up the following insertion loop
    subKeys.reserve(subKeyCount);

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
        if (retCode.Failed())
        {
            subKeys.clear();
            return retCode;
        }

        // On success, the ::RegEnumKeyEx API writes the length of the
        // subkey name in the subKeyNameLen output parameter
        // (not including the terminating NUL).
        // So I can build a wstring based on that length.
        subKeys.emplace_back(nameBuffer.get(), subKeyNameLen);
    }

    return ERROR_SUCCESS;
}


RegResult RegKey::TryEnumValues(std::vector<std::pair<std::wstring, DWORD>>& values) const
{
    values.clear();

    _ASSERTE(IsValid());

    // Get useful enumeration info, like the total number of values
    // and the maximum length of the value names
    DWORD valueCount = 0;
    DWORD maxValueNameLen = 0;
    RegResult retCode = ::RegQueryInfoKeyW(
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
    if (retCode.Failed())
    {
        return retCode;
    }

    // NOTE: According to the MSDN documentation, the size returned for value name max length
    // does *not* include the terminating NUL, so let's add +1 to take it into account
    // when I allocate the buffer for reading value names.
    maxValueNameLen++;

    // Preallocate a buffer for the value names
    auto nameBuffer = std::make_unique<wchar_t[]>(maxValueNameLen);

    // Should have been cleared at the beginning of the method
    _ASSERTE(values.empty());

    // Reserve room in the vector to speed up the following insertion loop
    values.reserve(valueCount);

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
        if (retCode.Failed())
        {
            values.clear();
            return retCode;
        }

        // On success, the RegEnumValue API writes the length of the
        // value name in the valueNameLen output parameter
        // (not including the terminating NUL).
        // So we can build a wstring based on that.
        values.emplace_back(
            std::wstring(nameBuffer.get(), valueNameLen),
            valueType
        );
    }

    return ERROR_SUCCESS;
}


std::wstring RegResult::ErrorMessage(const DWORD languageId) const
{
    // Invoke FormatMessage() to retrieve the error message from Windows
    ScopedLocalFree<wchar_t> messagePtr;
    DWORD retCode = ::FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER |
        FORMAT_MESSAGE_FROM_SYSTEM |
        FORMAT_MESSAGE_IGNORE_INSERTS,
        nullptr,
        m_result,
        languageId,
        reinterpret_cast<LPWSTR>(messagePtr.AddressOf()),
        0,
        nullptr);
    if (retCode == 0)
    {
        // FormatMessage failed: return an empty string
        return std::wstring();
    }

    // Safely copy the C-string returned by FormatMessage() into a std::wstring object,
    // and return it back to the caller.
    return std::wstring(messagePtr.Get());
}


} // namespace winreg


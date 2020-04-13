////////////////////////////////////////////////////////////////////////////////
//
//      *** Modern C++ Wrappers Around Windows Registry C API ***
// 
//               Copyright (C) by Giovanni Dicanio 
//
// ===========================================================================
// FILE: WinReg.cpp
// DESC: Non-inline method implementations for the WinReg wrapper.
// ===========================================================================
//
// The MIT License(MIT)
//
// Copyright(c) 2020 by Giovanni Dicanio
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


#include "WinReg.hpp"       // Public header


//------------------------------------------------------------------------------
//                      Module-Private Helper Functions
//------------------------------------------------------------------------------

namespace 
{

// Helper function to build a multi-string from a vector<wstring>
static std::vector<wchar_t> BuildMultiString(const std::vector<std::wstring>& data)
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
    multiString.reserve(totalLen);

    // Copy the single strings into the multi-string
    for (const auto& s : data)
    {      
        multiString.insert(multiString.end(), s.begin(), s.end());
        
        // Don't forget to NUL-terminate the current string
        multiString.push_back(L'\0');
    }

    // Add the last NUL-terminator
    multiString.push_back(L'\0');

    return multiString;
}

} // namespace 


//------------------------------------------------------------------------------
//                  RegKey Non-Inline Method Implementations
//------------------------------------------------------------------------------

namespace winreg
{

void RegKey::Create(
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
        throw RegException(retCode, "RegCreateKeyEx failed." );
    }

    // Safely close any previously opened key
    Close();

    // Take ownership of the newly created key
    m_hKey = hKey;
}


RegResult RegKey::TryCreate(
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


void RegKey::Open(
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
        throw RegException(retCode, "RegOpenKeyEx failed.");
    }

    // Safely close any previously opened key
    Close();

    // Take ownership of the newly created key
    m_hKey = hKey;
}


RegResult RegKey::TryOpen(
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
        throw RegException(retCode, "Cannot write multi-string value: RegSetValueEx failed.");
    }
}


RegResult RegKey::TrySetMultiStringValue(
    const std::wstring& valueName,
    const std::vector<std::wstring>& data
) // Can't be noexcept!
{
    //
    // NOTE:
    // This method can't be marked noexcept, because it calls details::BuildMultiString(),
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
        nullptr, // no subkey
        valueName.c_str(),
        flags,
        nullptr, // type not required
        nullptr, // output buffer not needed now
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot get size of string value: RegGetValue failed.");
    }

    // Allocate a string of proper size.
    // Note that dataSize is in bytes and includes the terminating NUL;
    // we have to convert the size from bytes to wchar_ts for wstring::resize.
    std::wstring result;
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
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot get string value: RegGetValue failed.");
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
        throw RegException(retCode, "Cannot get size of expand string value: RegGetValue failed.");
    }

    // Allocate a string of proper size.
    // Note that dataSize is in bytes and includes the terminating NUL.
    // We must convert from bytes to wchar_ts for wstring::resize.
    std::wstring result;
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
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException(retCode, "Cannot get expand string value: RegGetValue failed.");
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
        throw RegException(retCode, "Cannot get size of multi-string value: RegGetValue failed.");
    }

    // Allocate room for the result multi-string.
    // Note that dataSize is in bytes, but our vector<wchar_t>::resize method requires size 
    // to be expressed in wchar_ts.
    std::vector<wchar_t> data;
    data.resize(dataSize / sizeof(wchar_t));

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
        throw RegException(retCode, "Cannot get multi-string value: RegGetValue failed.");
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
        result.push_back(std::wstring(currStringPtr, currStringLength));

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
        throw RegException(retCode, "Cannot get size of binary data: RegGetValue failed.");
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
        throw RegException(retCode, "Cannot get binary data: RegGetValue failed.");
    }

    return data;
}


RegResult RegKey::TryGetStringValue(const std::wstring& valueName,
    std::wstring& result) const
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


RegResult RegKey::TryGetMultiStringValue(const std::wstring& valueName,
    std::vector<std::wstring>& result) const
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
    std::vector<wchar_t> data;
    data.resize(dataSize / sizeof(wchar_t));

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
        result.push_back(std::wstring(currStringPtr, currStringLength));

        // Move to the next string
        currStringPtr += currStringLength + 1;
    }

    _ASSERTE(retCode.IsOk());
    return retCode;
}


RegResult RegKey::TryGetBinaryValue(const std::wstring& valueName,
    std::vector<BYTE>& result) const
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
    DWORD subKeyCount{};
    DWORD maxSubKeyNameLen{};
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
            "RegQueryInfoKey failed while preparing for subkey enumeration."
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
            throw RegException(retCode, "Cannot enumerate subkeys: RegEnumKeyEx failed.");
        }

        // On success, the ::RegEnumKeyEx API writes the length of the
        // subkey name in the subKeyNameLen output parameter 
        // (not including the terminating NUL).
        // So I can build a wstring based on that length.
        subkeyNames.push_back(std::wstring(nameBuffer.get(), subKeyNameLen));
    }

    return subkeyNames;
}


std::vector<std::pair<std::wstring, DWORD>> RegKey::EnumValues() const
{
    _ASSERTE(IsValid());

    // Get useful enumeration info, like the total number of values
    // and the maximum length of the value names
    DWORD valueCount{};
    DWORD maxValueNameLen{};
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
            "RegQueryInfoKey failed while preparing for value enumeration."
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

    // Enumerate all the values
    for (DWORD index = 0; index < valueCount; index++)
    {
        // Get the name and the type of the current value
        DWORD valueNameLen = maxValueNameLen;
        DWORD valueType{};
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
            throw RegException(retCode, "Cannot enumerate values: RegEnumValue failed.");
        }

        // On success, the RegEnumValue API writes the length of the
        // value name in the valueNameLen output parameter 
        // (not including the terminating NUL).
        // So we can build a wstring based on that.
        valueInfo.push_back(
            std::make_pair(std::wstring(nameBuffer.get(), valueNameLen), valueType)
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
    DWORD subKeyCount{};
    DWORD maxSubKeyNameLen{};
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
        subKeys.push_back(std::wstring(nameBuffer.get(), subKeyNameLen));
    }

    return ERROR_SUCCESS;
}


RegResult RegKey::TryEnumValues(std::vector<std::pair<std::wstring, DWORD>>& values) const
{
    values.clear();

    _ASSERTE(IsValid());

    // Get useful enumeration info, like the total number of values
    // and the maximum length of the value names
    DWORD valueCount{};
    DWORD maxValueNameLen{};
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

    // The value names and types will be stored here
    std::vector<std::pair<std::wstring, DWORD>> valueInfo;

    // Should have been cleared at the beginning of the method
    _ASSERTE(values.empty());

    // Enumerate all the values
    for (DWORD index = 0; index < valueCount; index++)
    {
        // Get the name and the type of the current value
        DWORD valueNameLen = maxValueNameLen;
        DWORD valueType{};
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
        valueInfo.push_back(
            std::make_pair(std::wstring(nameBuffer.get(), valueNameLen), valueType)
        );
    }

    return ERROR_SUCCESS;
}


RegResult RegKey::TryConnectRegistry(const std::wstring& machineName,
    HKEY hKeyPredefined) noexcept
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


} // namespace winreg


#ifndef INCLUDE_GIOVANNI_DICANIO_WINREG_HPP
#define INCLUDE_GIOVANNI_DICANIO_WINREG_HPP


////////////////////////////////////////////////////////////////////////////////
//
//      *** Modern C++ Wrappers Around Windows Registry C API ***
// 
//               Copyright (C) by Giovanni Dicanio 
//  
// First version: 2017, January 22nd
// Last update: 2017, April 21th
// 
// E-mail: <giovanni.dicanio AT gmail.com>
// 
// Registry key handles are safely and conveniently wrapped 
// in the RegKey resource manager C++ class.
// 
// Errors are signaled throwing exceptions of class RegException.
// 
// There are free functions that wrap Win32 registry APIs as well.
// These functions convert Win32 error codes to C++ exceptions.
// 
// Unicode UTF-16 strings are represented using the std::wstring class; 
// ATL's CString is not used, to avoid dependencies from ATL or MFC.
// 
// This is a header-only self-contained reusable module.
//
// Compiler: Visual Studio 2015
// Code compiles clean at /W4 on both 32-bit and 64-bit builds.
// 
// ===========================================================================
//
// The MIT License(MIT)
//
// Copyright(c) 2017 Giovanni Dicanio
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

#include <memory>           // std::unique_ptr
#include <stdexcept>        // std::runtime_error
#include <string>           // std::wstring
#include <utility>          // std::swap, std::pair
#include <vector>           // std::vector


namespace winreg
{


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
    // Uses default KEY_READ|KEY_WRITE access; for finer grained control, 
    // call the Create() method overloads.
    // Throw RegException on failure.
    RegKey(HKEY hKeyParent, const std::wstring& subKey);

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
    HKEY Get() const noexcept;

    // Is the wrapped HKEY handle valid?
    bool IsValid() const noexcept;

    // Same as IsValid(), but allow a short "if (regKey)" syntax
    explicit operator bool() const noexcept;


    //
    // Operations
    // 

    // Close current HKEY handle.
    // If there's no valid handle, do nothing.
    void Close() noexcept;

    // Transfer ownership of current HKEY to the caller.
    // Note that the caller is responsible for closing the key handle!
    HKEY Detach() noexcept;

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
        REGSAM desiredAccess = KEY_READ | KEY_WRITE
    );


    //
    // Relational comparison operators are overloaded as non-members
    // ==, !=, <, <=, >, >=
    // 


private:
    // The wrapped registry key handle
    HKEY m_hKey{ nullptr };
};


//------------------------------------------------------------------------------
// An exception representing an error with the registry operations
//------------------------------------------------------------------------------
class RegException
    : public std::runtime_error
{
public:
    RegException(const char* message, LONG errorCode)
        : std::runtime_error{ message }
        , m_errorCode{ errorCode }
    {}

    RegException(const std::string& message, LONG errorCode)
        : std::runtime_error{ message }
        , m_errorCode{ errorCode }
    {}

    // Get the error code returned by Windows registry APIs
    LONG ErrorCode() const noexcept
    {
        return m_errorCode;
    }

private:
    // Error code, as returned by Windows registry APIs
    LONG m_errorCode;
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
{}


inline RegKey::RegKey(const HKEY hKeyParent, const std::wstring& subKey)
{
    Create(hKeyParent, subKey);
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
        ::RegCloseKey(m_hKey);
        
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


inline HKEY RegKey::Detach() noexcept
{
    HKEY hKey{ m_hKey };

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
    HKEY hKey{ nullptr };
    LONG retCode = ::RegCreateKeyEx(
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
        throw RegException{ "RegCreateKeyEx failed.", retCode };
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
    HKEY hKey{ nullptr };
    LONG retCode = ::RegOpenKeyEx(
        hKeyParent,
        subKey.c_str(),
        REG_NONE,           // default options
        desiredAccess,
        &hKey
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "RegOpenKeyEx failed.", retCode };
    }

    // Safely close any previously opened key
    Close();

    // Take ownership of the newly created key
    m_hKey = hKey;
}


//------------------------------------------------------------------------------
//                  Win32 Registry API Wrapper Functions
//------------------------------------------------------------------------------

inline void SetDwordValue(HKEY hKey, const std::wstring& valueName, DWORD data)
{
    LONG retCode = ::RegSetValueEx(
        hKey, 
        valueName.c_str(), 
        0, // reserved
        REG_DWORD, 
        reinterpret_cast<const BYTE*>(&data), 
        sizeof(data)
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot write DWORD value: RegSetValueEx failed.", retCode };
    }
}


inline void SetQwordValue(HKEY hKey, const std::wstring& valueName, const ULONGLONG& data)
{
    LONG retCode = ::RegSetValueEx(
        hKey,
        valueName.c_str(),
        0, // reserved
        REG_QWORD,
        reinterpret_cast<const BYTE*>(&data),
        sizeof(data)
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot write QWORD value: RegSetValueEx failed.", retCode };
    }
}


inline void SetStringValue(HKEY hKey, const std::wstring& valueName, const std::wstring& data)
{
    // String size including the terminating NUL, in bytes
    const DWORD dataSize = static_cast<DWORD>((data.length() + 1) * sizeof(wchar_t));
    
    LONG retCode = ::RegSetValueEx(
        hKey,
        valueName.c_str(),
        0, // reserved
        REG_SZ,
        reinterpret_cast<const BYTE*>(data.c_str()),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot write string value: RegSetValueEx failed.", retCode };
    }
}


inline void SetExpandStringValue(HKEY hKey, const std::wstring& valueName, const std::wstring& data)
{
    // String size including the terminating NUL, in bytes
    const DWORD dataSize = static_cast<DWORD>((data.length() + 1) * sizeof(wchar_t));

    LONG retCode = ::RegSetValueEx(
        hKey,
        valueName.c_str(),
        0, // reserved
        REG_EXPAND_SZ,
        reinterpret_cast<const BYTE*>(data.c_str()),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot write expand string value: RegSetValueEx failed.", retCode };
    }
}


namespace details
{

// Helper function to build a multi-string from a vector<wstring>
inline std::vector<wchar_t> BuildMultiString(const std::vector<std::wstring>& data)
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

} // namespace details


inline void SetMultiStringValue(
    HKEY hKey, 
    const std::wstring& valueName, 
    const std::vector<std::wstring>& data
)
{
    // First, we have to build a double-NUL-terminated multi-string from the input data
    const std::vector<wchar_t> multiString = details::BuildMultiString(data);

    // Total size, in bytes, of the whole multi-string structure
    const DWORD dataSize = static_cast<DWORD>(multiString.size() * sizeof(wchar_t));

    LONG retCode = ::RegSetValueEx(
        hKey,
        valueName.c_str(),
        0, // reserved
        REG_MULTI_SZ,
        reinterpret_cast<const BYTE*>(&multiString[0]),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot write multi-string value: RegSetValueEx failed.", retCode };
    }
}


inline void SetBinaryValue(HKEY hKey, const std::wstring& valueName, const std::vector<BYTE>& data)
{
    // Total data size, in bytes
    const DWORD dataSize = static_cast<DWORD>(data.size());

    LONG retCode = ::RegSetValueEx(
        hKey,
        valueName.c_str(),
        0, // reserved
        REG_BINARY,
        &data[0],
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot write binary data value: RegSetValueEx failed.", retCode };
    }
}


inline void SetBinaryValue(
    HKEY hKey,
    const std::wstring& valueName, 
    const void* data, 
    DWORD dataSize
)
{
    LONG retCode = ::RegSetValueEx(
        hKey,
        valueName.c_str(),
        0, // reserved
        REG_BINARY,
        static_cast<const BYTE*>(data),
        dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot write binary data value: RegSetValueEx failed.", retCode };
    }
}


inline DWORD GetDwordValue(HKEY hKey, const std::wstring& valueName)
{
    DWORD data{};                   // to be read from the registry
    DWORD dataSize = sizeof(data);  // size of data, in bytes

    const DWORD flags = RRF_RT_REG_DWORD;
    LONG retCode = ::RegGetValue(
        hKey,
        nullptr, // no subkey
        valueName.c_str(),
        flags,
        nullptr, // type not required
        &data,
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot get DWORD value: RegGetValue failed.", retCode };
    }

    return data;
}


inline ULONGLONG GetQwordValue(HKEY hKey, const std::wstring& valueName)
{
    ULONGLONG data{};               // to be read from the registry
    DWORD dataSize = sizeof(data);  // size of data, in bytes

    const DWORD flags = RRF_RT_REG_QWORD;
    LONG retCode = ::RegGetValue(
        hKey,
        nullptr, // no subkey
        valueName.c_str(),
        flags,
        nullptr, // type not required
        &data,
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot get QWORD value: RegGetValue failed.", retCode };
    }

    return data;
}


inline std::wstring GetStringValue(HKEY hKey, const std::wstring& valueName)
{
    // Get the size of the result string
    DWORD dataSize = 0; // size of data, in bytes
    const DWORD flags = RRF_RT_REG_SZ;
    LONG retCode = ::RegGetValue(
        hKey,
        nullptr, // no subkey
        valueName.c_str(),
        flags,
        nullptr, // type not required
        nullptr, // output buffer not needed now
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot get size of string value: RegGetValue failed.", retCode };
    }

    // Allocate a string of proper size.
    // Note that dataSize is in bytes and includes the terminating NUL;
    // we have to convert the size from bytes to wchar_ts for wstring::resize.
    std::wstring result;
    result.resize(dataSize / sizeof(wchar_t));

    // Call RegGetValue for the second time to read the string's content
    retCode = ::RegGetValue(
        hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        &result[0], // output buffer
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot get string value: RegGetValue failed.", retCode };
    }

    // Remove the NUL terminator scribbled by RegGetValue from the wstring
    result.resize((dataSize / sizeof(wchar_t)) - 1);

    return result;
}


enum class ExpandStringOption
{
    DontExpand,
    Expand
};

inline std::wstring GetExpandStringValue(
    HKEY hKey,
    const std::wstring& valueName, 
    ExpandStringOption expandOption = ExpandStringOption::DontExpand
)
{
    DWORD flags = RRF_RT_REG_EXPAND_SZ;

    // Adjust the flag for RegGetValue considering the expand string option specified by the caller   
    if (expandOption == ExpandStringOption::DontExpand)
    {
        flags |= RRF_NOEXPAND;
    }

    // Get the size of the result string
    DWORD dataSize = 0; // size of data, in bytes
    LONG retCode = ::RegGetValue(
        hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        nullptr,    // output buffer not needed now
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot get size of expand string value: RegGetValue failed.", retCode };
    }

    // Allocate a string of proper size.
    // Note that dataSize is in bytes and includes the terminating NUL.
    // We must convert from bytes to wchar_ts for wstring::resize.
    std::wstring result;
    result.resize(dataSize / sizeof(wchar_t));

    // Call RegGetValue for the second time to read the string's content
    retCode = ::RegGetValue(
        hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        &result[0], // output buffer
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot get expand string value: RegGetValue failed.", retCode };
    }

    // Remove the NUL terminator scribbled by RegGetValue from the wstring
    result.resize((dataSize / sizeof(wchar_t)) - 1);

    return result;
}


inline std::vector<std::wstring> GetMultiStringValue(HKEY hKey, const std::wstring& valueName)
{
    // Request the size of the multi-string, in bytes
    DWORD dataSize = 0;
    const DWORD flags = RRF_RT_REG_MULTI_SZ;
    LONG retCode = ::RegGetValue(
        hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        nullptr,    // output buffer not needed now
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot get size of multi-string value: RegGetValue failed.", retCode };
    }

    // Allocate room for the result multi-string.
    // Note that dataSize is in bytes, but our vector<wchar_t>::resize method requires size 
    // to be expressed in wchar_ts.
    std::vector<wchar_t> data;
    data.resize(dataSize / sizeof(wchar_t));

    // Read the multi-string from the registry into the vector object
    retCode = ::RegGetValue(
        hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // no type required
        &data[0],   // output buffer
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot get multi-string value: RegGetValue failed.", retCode };
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
        result.push_back(std::wstring{ currStringPtr, currStringLength });

        // Move to the next string
        currStringPtr += currStringLength + 1;
    }

    return result;
}


inline std::vector<BYTE> GetBinaryValue(HKEY hKey, const std::wstring& valueName)
{
    // Get the size of the binary data
    DWORD dataSize = 0; // size of data, in bytes
    const DWORD flags = RRF_RT_REG_BINARY;
    LONG retCode = ::RegGetValue(
        hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        nullptr,    // output buffer not needed now
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot get size of binary data: RegGetValue failed.", retCode };
    }

    // Allocate a buffer of proper size to store the binary data
    std::vector<BYTE> data(dataSize);

    // Call RegGetValue for the second time to read the data content
    retCode = ::RegGetValue(
        hKey,
        nullptr,    // no subkey
        valueName.c_str(),
        flags,
        nullptr,    // type not required
        &data[0],   // output buffer
        &dataSize
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot get binary data: RegGetValue failed.", retCode };
    }

    return data;
}


// Return the DWORD type ID for the input registry value
inline DWORD QueryValueType(HKEY hKey, const std::wstring& valueName)
{
    DWORD typeId{};     // will be returned by RegGetValue

    const DWORD flags = RRF_RT_ANY;     // no type restriction
    LONG retCode = ::RegGetValue(
        hKey,
        nullptr, // no subkey
        valueName.c_str(),
        flags,
        &typeId,
        nullptr, // not interested
        nullptr  // not interested
    );
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "Cannot get the value type: RegGetValue failed.", retCode };
    }

    return typeId;
}


inline void QueryInfoKey(HKEY hKey, DWORD& subKeys, DWORD &values, FILETIME& lastWriteTime)
{
    LONG retCode = ::RegQueryInfoKey(
        hKey,
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
        throw RegException{ "RegQueryInfoKey failed.", retCode };
    }
}


// Enumerate the subkeys of a specified registry key, using RegEnumKeyEx
inline std::vector<std::wstring> EnumSubKeys(HKEY hKey)
{
    // Get some useful enumeration info, like the total number of subkeys
    // and the maximum length of the subkey names
    DWORD subKeyCount{};
    DWORD maxSubKeyNameLen{};
    LONG retCode = ::RegQueryInfoKey(
        hKey,
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
        throw RegException{ "RegQueryInfoKey failed while preparing for subkey enumeration.", retCode };
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
        retCode = ::RegEnumKeyEx(
            hKey,
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
            throw RegException{ "Cannot enumerate subkeys: RegEnumKeyEx failed.", retCode };
        }

        // On success, the ::RegEnumKeyEx API writes the length of the
        // subkey name in the subKeyNameLen output parameter 
        // (not including the terminating NUL).
        // So I can build a wstring based on that length.
        subkeyNames.push_back(std::wstring{ nameBuffer.get(), subKeyNameLen });
    }

    return subkeyNames;
}


// Enumerate the values under the input registry key, using RegEnumValue.
// Returns a vector of pairs: In each pair, the wstring is the value name, 
// the DWORD is the value type.
inline std::vector<std::pair<std::wstring, DWORD>> EnumValues(HKEY hKey)
{
    // Get useful enumeration info, like the total number of values
    // and the maximum length of the value names
    DWORD valueCount{};
    DWORD maxValueNameLen{};
    LONG retCode = ::RegQueryInfoKey(
        hKey,
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
            "RegQueryInfoKey failed while preparing for value enumeration.",
            retCode
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

    // Enumerate all the values
    for (DWORD index = 0; index < valueCount; index++)
    {
        // Get the name and the type of the current value
        DWORD valueNameLen = maxValueNameLen;
        DWORD valueType{};
        retCode = ::RegEnumValue(
            hKey,
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
            throw RegException{ "Cannot enumerate values: RegEnumValue failed.", retCode };
        }

        // On success, the RegEnumValue API writes the length of the
        // value name in the valueNameLen output parameter 
        // (not including the terminating NUL).
        // So we can build a wstring based on that.
        valueInfo.push_back(
            std::make_pair(std::wstring{ nameBuffer.get(), valueNameLen }, valueType)
        );
    }

    return valueInfo;
}


inline void DeleteValue(HKEY hKey, const std::wstring& valueName)
{
    LONG retCode = ::RegDeleteValue(hKey, valueName.c_str());
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "RegDeleteValue failed.", retCode };
    }
}


inline void DeleteKey(HKEY hKey, const std::wstring& subKey, REGSAM desiredAccess)
{
    LONG retCode = ::RegDeleteKeyEx(hKey, subKey.c_str(), desiredAccess, 0);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "RegDeleteKeyEx failed.", retCode };
    }
}


inline void FlushKey(HKEY hKey)
{
    LONG retCode = ::RegFlushKey(hKey);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "RegFlushKey failed.", retCode };
    }
}


inline void LoadKey(HKEY hKey, const std::wstring& subKey, const std::wstring& filename)
{
    LONG retCode = ::RegLoadKey(hKey, subKey.c_str(), filename.c_str());
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "RegLoadKey failed.", retCode };
    }
}


inline void SaveKey(
    HKEY hKey,
    const std::wstring& filename, 
    SECURITY_ATTRIBUTES* const securityAttributes
)
{
    LONG retCode = ::RegSaveKey(hKey, filename.c_str(), securityAttributes);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "RegSaveKey failed.", retCode };
    }
}


inline void EnableReflectionKey(HKEY hKey)
{
    LONG retCode = ::RegEnableReflectionKey(hKey);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "RegEnableReflectionKey failed.", retCode };
    }
}


inline void DisableReflectionKey(HKEY hKey)
{
    LONG retCode = ::RegDisableReflectionKey(hKey);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "RegDisableReflectionKey failed.", retCode };
    }
}


inline bool QueryReflectionKey(HKEY hKey)
{
    BOOL isReflectionDisabled = FALSE;
    LONG retCode = ::RegQueryReflectionKey(hKey, &isReflectionDisabled);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "RegQueryReflectionKey failed.", retCode };
    }

    return (isReflectionDisabled ? true : false);
}


inline void ConnectRegistry(const std::wstring& machineName, HKEY hKey, HKEY& hKeyResult)
{
    LONG retCode = ::RegConnectRegistry(machineName.c_str(), hKey, &hKeyResult);
    if (retCode != ERROR_SUCCESS)
    {
        throw RegException{ "RegConnectRegistry failed.", retCode };
    }
}


// Return a string representation of Windows registry types
inline std::wstring RegTypeToString(DWORD regType)
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


} // namespace winreg


#endif // INCLUDE_GIOVANNI_DICANIO_WINREG_HPP

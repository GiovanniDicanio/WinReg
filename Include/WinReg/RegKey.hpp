////////////////////////////////////////////////////////////////////////////////
//
//         *** Modern C++ Wrappers Around Windows Registry C API ***
//
//       Copyright (C) by Giovanni Dicanio (giovanni.dicanio@gmail.com)
//
//----------------------------------------------------------------------------
// SPDX-License-Identifier: MIT
////////////////////////////////////////////////////////////////////////////////

#ifndef GIOVANNI_DICANIO_WINREG_REGKEY_HPP
#define GIOVANNI_DICANIO_WINREG_REGKEY_HPP


#include <Windows.h>        // Windows Platform SDK


#ifndef GIOVANNI_DICANIO_WINREG_USE_CPP17

    //
    // C++20 features
    //

    #include <compare>          // std::compare_three_way, std::strong_ordering
    #include <span>             // std::span

#endif // GIOVANNI_DICANIO_WINREG_USE_CPP17

#include <string>           // std::string, std::wstring
#include <utility>          // std::swap, std::pair, std::move
#include <vector>           // std::vector

#include "RegException.hpp"
#include "RegExpected.hpp"
#include "RegResult.hpp"


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
    void SetQwordValue(const std::wstring& valueName, ULONGLONG data);
    void SetStringValue(const std::wstring& valueName, const std::wstring& data);
    void SetExpandStringValue(const std::wstring& valueName, const std::wstring& data);
    void SetMultiStringValue(const std::wstring& valueName, const std::vector<std::wstring>& data);
    void SetBinaryValue(const std::wstring& valueName, const std::vector<BYTE>& data);
    void SetBinaryValue(const std::wstring& valueName, const void* data, DWORD dataSize);

    // std::span is not available in C++17; must use at least C++20.
#ifndef GIOVANNI_DICANIO_WINREG_USE_CPP17
    void SetBinaryValue(const std::wstring& valueName, std::span<const BYTE> data);
#endif // GIOVANNI_DICANIO_WINREG_USE_CPP17


    //
    // Registry Value Setters Returning RegResult
    // (instead of throwing RegException on error)
    //

    [[nodiscard]] RegResult TrySetDwordValue(const std::wstring& valueName, DWORD data) noexcept;
    [[nodiscard]] RegResult TrySetQwordValue(const std::wstring& valueName, ULONGLONG data) noexcept;

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

    // std::span is not available in C++17; must use at least C++20.
#ifndef GIOVANNI_DICANIO_WINREG_USE_CPP17
    [[nodiscard]] RegResult TrySetBinaryValue(const std::wstring& valueName,
        std::span<const BYTE> data);
#endif // GIOVANNI_DICANIO_WINREG_USE_CPP17


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


    // Raw value returned e.g. from the GetRawValue() method
    struct RawValue
    {
        DWORD Type{ REG_NONE };
        std::vector<BYTE> Data;
    };

    //
    // GetRawValue is very similar to GetBinaryValue, except that here in the "raw" version
    // I do *not* check the type of the value in the registry to be REG_BINARY.
    // In this way, the type is *not* restricted to REG_BINARY, and this GetRawValue method
    // can be invoked to retrieve binary data for unsupported or less documented types like
    // REG_RESOURCE_LIST, REG_FULL_RESOURCE_DESCRIPTOR, REG_LINK, etc.
    //
    // This addition came from an issue opened on this code on GitHub:
    //
    // No way to get the value of a REG_RESOURCE_LIST key and some others
    // https://github.com/GiovanniDicanio/WinReg/issues/77
    //
    [[nodiscard]] RawValue GetRawValue(const std::wstring& valueName) const;


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

    // Similar to TryGetBinaryValue, but without type restriction to REG_BINARY.
    // Can be used to retrieve raw binary data for other types like REG_RESOURCE_LIST, etc.
    // Please read the comment to the GetRawValue method to learn more.
    [[nodiscard]] RegExpected<RawValue> TryGetRawValue(const std::wstring& valueName) const;


    //
    // Query Operations
    //

    // Information about a registry key (retrieved by QueryInfoKey)
    struct InfoKey
    {
        DWORD    NumberOfSubKeys{ 0 };
        DWORD    NumberOfValues{ 0 };
        FILETIME LastWriteTime{};
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


    //
    // NOTE on Enumeration Methods: EnumSubKeys/EnumValues/TryEnumSubKeys/TryEnumValues
    // --------------------------------------------------------------------------------
    //
    // In the current implementation, the subkey/value count and max-name-length come
    // from one RegQueryInfoKeyW snapshot, then the loop enumerates by index afterward.
    // If the registry is mutated concurrently (e.g. subkeys/values added or removed
    // by another process) between that snapshot and the loop, RegEnumKeyExW/RegEnumValueW
    // can return ERROR_NO_MORE_ITEMS or ERROR_MORE_DATA mid-loop, which gets thrown
    // as a RegException (or returned as a RegExpected containing an error code).
    // This is a TOCTOU race inherent to the underlying Win32 enumeration API, and
    // probably worth noting here.
    //
    // As a side note, the MSDN documentation about RegEnumKeyEx reads:
    //
    // https://learn.microsoft.com/en-us/windows/win32/api/winreg/nf-winreg-regenumkeyexw
    //
    //      While an application is using the RegEnumKeyEx function,
    //      it should not make calls to any registration functions
    //      that might change the key being enumerated.
    //
    //

    // Enumerate the subkeys of the registry key, using RegEnumKeyEx
    [[nodiscard]] std::vector<std::wstring> EnumSubKeys() const;

    // Enumerate the values under the registry key, using RegEnumValue.
    // Returns a vector of pairs: In each pair, the wstring is the value name,
    // the DWORD is the value type.
    [[nodiscard]] std::vector<std::pair<std::wstring, DWORD>> EnumValues() const;

    // Check if the current key contains the specified value
    [[nodiscard]] bool ContainsValue(const std::wstring& valueName) const;

    // Check if the current key contains the specified sub-key.
    //
    // NOTE:
    // wow64Flag must match the registry view used to open this key.
    // The default mode is opening the sub-key with the KEY_WOW64_64KEY flag,
    // as done in other places in the library, like the constructors.
    // Other options are KEY_WOW64_32KEY, or 0 for the default view.
    //
    [[nodiscard]] bool ContainsSubKey(const std::wstring& subKey,
                                      REGSAM wow64Flag = KEY_WOW64_64KEY) const;


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

    // Check if the current key contains the specified sub-key.
    // NOTE:
    // wow64Flag must match the registry view used to open this key.
    // The default mode is opening the sub-key with the KEY_WOW64_64KEY flag,
    // as done in other places in the library, like the constructors.
    // Other options are KEY_WOW64_32KEY, or 0 for the default view.
    //
    [[nodiscard]] RegExpected<bool> TryContainsSubKey(const std::wstring& subKey,
                                                      REGSAM wow64Flag = KEY_WOW64_64KEY) const;


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
    [[nodiscard]] static const wchar_t* RegTypeToString(DWORD regType);


#ifndef GIOVANNI_DICANIO_WINREG_USE_CPP17

    //
    // C++20 code
    //

    //
    // Relational comparison operators are overloaded as friend non-members.
    //
    // In C++20 we can simply use the default implementation of operator==
    // (which does a member-wise equality comparison), and implement operator<=>
    // using a simple invocation to std::compare_three_way{}.
    // The compiler will automatically synthesize all the remaining
    // relational operators (e.g. !=, <, <=, >, >=).
    //

    [[nodiscard]] friend bool operator==(const RegKey&, const RegKey&) noexcept = default;

    [[nodiscard]] friend std::strong_ordering operator<=>(const RegKey& lhs,
                                                          const RegKey& rhs) noexcept
    {
        return std::compare_three_way{}(lhs.m_hKey, rhs.m_hKey);
    }

#else

    //
    // C++17 code
    //

    //
    // Relational comparison operators are overloaded as non-members
    // ==, !=, <, <=, >, >=
    //
    // (see RegKey_impl.hpp)
    //

#endif // GIOVANNI_DICANIO_WINREG_USE_CPP17


    //
    // Private Implementation
    //

private:
    // The wrapped registry key handle
    HKEY m_hKey{ nullptr };
};


} // namespace winreg


#include "RegKey_impl.hpp"      // Inline implementations


#endif // GIOVANNI_DICANIO_WINREG_REGKEY_HPP

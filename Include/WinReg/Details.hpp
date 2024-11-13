////////////////////////////////////////////////////////////////////////////////
// FILE: Details.hpp
// DESC: Internal helper functions and classes for the WinReg library.
// AUTHOR: Giovanni Dicanio
////////////////////////////////////////////////////////////////////////////////

#ifndef GIOVANNI_DICANIO_WINREG_DETAILS_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_DETAILS_HPP_INCLUDED


#include <Windows.h>        // Windows Platform SDK
#include <crtdbg.h>         // _ASSERTE

#include <limits>           // std::numeric_limits
#include <stdexcept>        // std::runtime_error, std::overflow_error
#include <string>           // std::string, std::wstring
#include <vector>           // std::vector

#include "WinReg/RegException.hpp"
#include "WinReg/RegExpected.hpp"
#include "WinReg/Utf8Conv.hpp"      // For UTF-8 <-> UTF-16 conversions


namespace winreg
{

//------------------------------------------------------------------------------
//                  Private Helper Classes and Functions
//------------------------------------------------------------------------------

namespace details
{


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


//------------------------------------------------------------------------------
// Return true if casting a size_t value to a DWORD is safe
// (e.g. there is no overflow); false otherwise.
//------------------------------------------------------------------------------
[[nodiscard]] inline bool SizeToDwordCastIsSafe([[maybe_unused]] const size_t size) noexcept
{
#ifdef _WIN64

    //
    // In 64-bit builds, DWORD is an unsigned 32-bit integer,
    // while size_t is an unsigned *64-bit* integer.
    // So we need to pay attention to the conversion from size_t --> to DWORD.
    //

    using DestinationType = DWORD;

    // Pre-compute at compile-time the maximum value that can be stored by a DWORD.
    // Note that this value is stored in a size_t for proper comparison with the 'size' parameter.
    constexpr size_t kMaxDwordValue = static_cast<size_t>((std::numeric_limits<DestinationType>::max)());

    // Check against overflow
    if (size > kMaxDwordValue)
    {
        // Overflow from size_t to DWORD
        return false;
    }

    // The conversion is safe
    return true;

#else
    //
    // In 32-bit builds with Microsoft Visual C++, a size_t is an unsigned 32-bit value,
    // just like a DWORD. So, we can optimized this case out for 32-bit builds.
    //

    static_assert(sizeof(size_t) == sizeof(DWORD)); // Both 32-bit unsigned integers on 32-bit x86
    //UNREFERENCED_PARAMETER(size); // Replaced with [[maybe_unused]] for compatibility with MinGW 32-bit
    return true;

#endif // _WIN64
}


//------------------------------------------------------------------------------
// Safely cast a size_t value (usually from the STL)
// to a DWORD (usually for Win32 API calls).
// In case of overflow, throws an exception of type std::overflow_error.
//------------------------------------------------------------------------------
[[nodiscard]] inline DWORD SafeCastSizeToDword(const size_t size)
{

#ifdef _WIN64

    //
    // In 64-bit builds, DWORD is an unsigned 32-bit integer,
    // while size_t is an unsigned *64-bit* integer.
    // So we need to pay attention to the conversion from size_t --> to DWORD.
    //

    using DestinationType = DWORD;

    // Check against overflow
    if (!SizeToDwordCastIsSafe(size))
    {
        throw std::overflow_error(
            "Input size_t value is too big: size_t value doesn't fit into a DWORD.");
    }

    return static_cast<DestinationType>(size);

#else
    //
    // In 32-bit builds with Microsoft Visual C++, a size_t is an unsigned 32-bit value,
    // just like a DWORD. So, we can optimize this case out for 32-bit builds.
    //

    _ASSERTE(SizeToDwordCastIsSafe(size)); // double-check just in debug builds

    static_assert(sizeof(size_t) == sizeof(DWORD)); // Both 32-bit unsigned integers on 32-bit x86
    return static_cast<DWORD>(size);

#endif // _WIN64
}


//------------------------------------------------------------------------------
// String trait definition for having UTF-8 strings at the library interface
//------------------------------------------------------------------------------
struct StringTraitUtf8
{
    // Use std::string to store UTF-8 strings
    typedef std::string StringType;

    static StringType ConstructFromUtf16(const std::wstring& sourceUtf16)
    {
        return details::Utf16ToUtf8(sourceUtf16);
    }

    static std::wstring ToUtf16(const StringType& sourceUtf8)
    {
        return details::Utf8ToUtf16(sourceUtf8);
    }
};


//------------------------------------------------------------------------------
// String trait definition for having UTF-16 strings at the library interface
//------------------------------------------------------------------------------
struct StringTraitUtf16
{
    // Use std::wstring to store UTF-16 strings
    typedef std::wstring StringType;

    //
    // Since this trait represents a UTF-16 string, the conversions with UTF-16
    // and this trait strings are trivial (just pass-through).
    //

    static StringType ConstructFromUtf16(const std::wstring& sourceUtf16)
    {
        return sourceUtf16;
    }

    static std::wstring ToUtf16(const StringType& sourceUtf16)
    {
        return sourceUtf16;
    }
};



//------------------------------------------------------------------------------
// Convert a UTF-16 string vector into a UTF-8 string vector.
//------------------------------------------------------------------------------
[[nodiscard]] inline std::vector<std::string>
Utf8StringVectorFromUtf16(const std::vector<std::wstring>& utf16Strings)
{
    // Result vector
    std::vector<std::string> utf8Strings;

    // Make room in the result vector, as we already know how many input strings we have
    utf8Strings.reserve(utf16Strings.size());

    // Convert from UTF-16 to UTF-8 for each input string
    for (const auto& utf16 : utf16Strings)
    {
        utf8Strings.push_back(details::Utf16ToUtf8(utf16));
    }

    return utf8Strings;
}


//------------------------------------------------------------------------------
// Convert a UTF-8 string vector into a UTF-16 string vector.
//------------------------------------------------------------------------------
[[nodiscard]] inline std::vector<std::wstring>
Utf16StringVectorFromUtf8(const std::vector<std::string>& utf8Strings)
{
    // Result vector
    std::vector<std::wstring> utf16Strings;

    // Make room in the result vector, as we already know how many input strings we have
    utf16Strings.reserve(utf8Strings.size());

    // Convert from UTF-8 to UTF-16 for each input string
    for (const auto& utf8 : utf8Strings)
    {
        utf16Strings.push_back(details::Utf8ToUtf16(utf8));
    }

    return utf16Strings;
}


} // namespace details
} // namespace winreg

#endif // GIOVANNI_DICANIO_WINREG_DETAILS_HPP_INCLUDED

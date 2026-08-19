////////////////////////////////////////////////////////////////////////////////
//
//         *** Modern C++ Wrappers Around Windows Registry C API ***
//
//       Copyright (C) by Giovanni Dicanio (giovanni.dicanio@gmail.com)
//
//----------------------------------------------------------------------------
// SPDX-License-Identifier: MIT
////////////////////////////////////////////////////////////////////////////////


#ifndef GIOVANNI_DICANIO_WINREG_INTERNAL_MULTISTRING_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_INTERNAL_MULTISTRING_HPP_INCLUDED


#include <Windows.h>        // Windows Platform SDK

#include <string.h>         // wcslen

#include <string>           // std::wstring
#include <vector>           // std::vector

#include "WinReg/RegException.hpp"



////////////////////////////////////////////////////////////////////////////////
// Multi-string helper functions
////////////////////////////////////////////////////////////////////////////////


namespace winreg
{

namespace winreg_internal
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
[[nodiscard]] std::vector<wchar_t> BuildMultiString(const std::vector<std::wstring>& data);


//------------------------------------------------------------------------------
// Return true if the wchar_t sequence stored in 'data' terminates
// with two null (L'\0') wchar_t's
//------------------------------------------------------------------------------
[[nodiscard]] bool IsDoubleNullTerminated(const std::vector<wchar_t>& data);


//------------------------------------------------------------------------------
// Given a sequence of wchar_ts representing a double-null-terminated string,
// returns a vector of wstrings that represent the single strings.
//
// Also supports embedded empty strings in the sequence.
//------------------------------------------------------------------------------
[[nodiscard]] std::vector<std::wstring> ParseMultiString(const std::vector<wchar_t>& data);



//------------------------------------------------------------------------------
//                      Inline Function Implementation
//------------------------------------------------------------------------------


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


inline bool IsDoubleNullTerminated(const std::vector<wchar_t>& data)
{
    // First check that there's enough room for at least two nulls
    if (data.size() < 2)
    {
        return false;
    }

    // Check that the sequence terminates with two nulls (L'\0', L'\0')
    const size_t lastPosition = data.size() - 1;
    return (data[lastPosition] == L'\0') &&
           (data[lastPosition - 1] == L'\0');
}


inline std::vector<std::wstring> ParseMultiString(const std::vector<wchar_t>& data)
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


} // namespace winreg_internal

} // namespace winreg


#endif // GIOVANNI_DICANIO_WINREG_INTERNAL_MULTISTRING_HPP_INCLUDED

////////////////////////////////////////////////////////////////////////////////
//
//         *** Modern C++ Wrappers Around Windows Registry C API ***
//
//       Copyright (C) by Giovanni Dicanio (giovanni.dicanio@gmail.com)
//
//----------------------------------------------------------------------------
// SPDX-License-Identifier: MIT
////////////////////////////////////////////////////////////////////////////////

#pragma once


#include <Windows.h>            // Windows Platform SDK

#include <expected>             // std::unexpected
#include <string>               // std::wstring

#include "../RegException.hpp"
#include "../RegExpected.hpp"
#include "../detail/Config.hpp"


////////////////////////////////////////////////////////////////////////////////
// Expand String Helper Functions
////////////////////////////////////////////////////////////////////////////////

namespace winreg
{

namespace detail
{


//------------------------------------------------------------------------------
// Expand environment strings to a std::wstring, invoking the
// ExpandEnvironmentStringsW API.
//------------------------------------------------------------------------------
[[nodiscard]] std::wstring ExpandEnvironmentStringsToWstring(const std::wstring& source);


//------------------------------------------------------------------------------
// Expand environment strings to a std::wstring, invoking the
// ExpandEnvironmentStringsW API.
// Return errors via RegExpected.
//------------------------------------------------------------------------------
[[nodiscard]] RegExpected<std::wstring> TryExpandEnvironmentStringsToWstring(const std::wstring& source);



//------------------------------------------------------------------------------
//                      Inline Function Implementations
//------------------------------------------------------------------------------


//------------------------------------------------------------------------------
// Pass a fallback error code (e.g. ERROR_INVALID_DATA) if GetLastError
// returns 0, but you still have an error.
//------------------------------------------------------------------------------
[[nodiscard]] inline LSTATUS ErrorOr(const DWORD error, const LSTATUS fallback) noexcept
{
    return (error != ERROR_SUCCESS) ? static_cast<LSTATUS>(error) : fallback;
}


inline std::wstring ExpandEnvironmentStringsToWstring(const std::wstring& source)
{
    RegExpected<std::wstring> result = TryExpandEnvironmentStringsToWstring(source);
    if (!result.has_value())
    {
        throw ::winreg::RegException{ result.error().Code(),
            "Cannot expand environment strings: ExpandEnvironmentStringsW failed." };
    }

    return std::move(result).value();
}


inline RegExpected<std::wstring> TryExpandEnvironmentStringsToWstring(const std::wstring& source)
{
    if (source.empty())
    {
        return std::wstring{};
    }

    // Query the destination string size, in wchar_ts, including the terminating NUL
    DWORD cchRequired = ::ExpandEnvironmentStringsW(source.c_str(), nullptr, 0);
    if (cchRequired == 0)
    {
        const DWORD lastError = ::GetLastError();
        return std::unexpected{ RegResult{ ErrorOr(lastError, ERROR_INVALID_DATA) } };
    }

    std::wstring result;
    for (int attempt = 0; attempt < detail::kMaxReadAttempts; ++attempt)
    {
        result.resize(cchRequired);

        // Read the string into the buffer we own
        const DWORD cchWritten = ::ExpandEnvironmentStringsW(source.c_str(), result.data(), cchRequired);
        if (cchWritten == 0)
        {
            const DWORD lastError = ::GetLastError();
            return std::unexpected{ RegResult{ ErrorOr(lastError, ERROR_INVALID_DATA) } };
        }

        if (cchWritten <= cchRequired)
        {
            // Keep what was written, then cut at the first NUL (drops the terminator)
            result.resize(cchWritten);                              // what the API says it wrote
            result.resize(wcsnlen(result.data(), result.size()));   // cut at the first NUL
            return result;
        }

        // Buffer too small: the environment changed between calls; retry with new size
        cchRequired = cchWritten;
    }

    // The string kept growing while expanding
    return std::unexpected{ RegResult{ ERROR_MORE_DATA } };
}


} // namespace detail

} // namespace winreg


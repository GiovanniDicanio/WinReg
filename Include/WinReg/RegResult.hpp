////////////////////////////////////////////////////////////////////////////////
// FILE: RegResult.hpp
// DESC: Define class to wrap LSTATUS return codes used by Windows Registry APIs
// AUTHOR: Giovanni Dicanio
////////////////////////////////////////////////////////////////////////////////

#ifndef GIOVANNI_DICANIO_WINREG_REGRESULT_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_REGRESULT_HPP_INCLUDED

#include <Windows.h>            // Windows Platform SDK

#include <string>               // std::wstring

#include "WinReg/ScopedLocalFree.hpp"


namespace winreg
{

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


inline std::wstring RegResult::ErrorMessage(const DWORD languageId) const
{
    // Invoke FormatMessage() to retrieve the error message from Windows
    details::ScopedLocalFree<wchar_t> messagePtr;
    DWORD retCode = ::FormatMessageW(
        FORMAT_MESSAGE_ALLOCATE_BUFFER |
        FORMAT_MESSAGE_FROM_SYSTEM |
        FORMAT_MESSAGE_IGNORE_INSERTS,
        nullptr,
        m_result,
        languageId,
        reinterpret_cast<LPWSTR>(messagePtr.AddressOf()),
        0,
        nullptr
    );
    if (retCode == 0)
    {
        // FormatMessage failed: return an empty string
        return std::wstring{};
    }

    // Safely copy the C-string returned by FormatMessage() into a std::wstring object,
    // and return it back to the caller.
    return std::wstring{ messagePtr.Get() };
}


} // namespace winreg

#endif // GIOVANNI_DICANIO_WINREG_REGRESULT_HPP_INCLUDED

////////////////////////////////////////////////////////////////////////////////
// FILE: Utf8Conv.hpp
// DESC: UTF-8 <-> UTF-16 helper conversion functions.
// AUTHOR: Giovanni Dicanio
////////////////////////////////////////////////////////////////////////////////

#ifndef GIOVANNI_DICANIO_WINREG_UTF8CONV_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_UTF8CONV_HPP_INCLUDED

////////////////////////////////////////////////////////////////////////////////
//
// UTF-8 <-> UTF-16 conversion functions
//
// Code based from my GitHub repo:
// https://github.com/GiovanniDicanio/UnicodeConvStd/tree/using-string-views
//
// Header-only implementation:
// https://github.com/GiovanniDicanio/UnicodeConvStd/blob/using-string-views/UnicodeConvStd/UnicodeConvStd.hpp
//
////////////////////////////////////////////////////////////////////////////////


//==============================================================================
//                              Includes
//==============================================================================

#include <windows.h>    // Win32 Platform SDK

#include <crtdbg.h>     // _ASSERTE

#include <limits>       // std::numeric_limits
#include <stdexcept>    // std::runtime_error, std::overflow_error
#include <string>       // std::string, std::wstring
#include <string_view>  // std::string_view, std::wstring_view


//==============================================================================
//                              Implementation
//==============================================================================


namespace winreg::details
{

//------------------------------------------------------------------------------
// Represents an error during Unicode conversions
//------------------------------------------------------------------------------
class UnicodeConversionException
    : public std::runtime_error
{
public:

    enum class ConversionType
    {
        FromUtf16ToUtf8,
        FromUtf8ToUtf16
    };

    UnicodeConversionException(DWORD errorCode, ConversionType conversionType, const char* message)
        : std::runtime_error(message),
        m_errorCode(errorCode),
        m_conversionType(conversionType)
    {
    }

    UnicodeConversionException(DWORD errorCode, ConversionType conversionType, const std::string& message)
        : std::runtime_error(message),
        m_errorCode(errorCode),
        m_conversionType(conversionType)
    {
    }

    [[nodiscard]] DWORD GetErrorCode() const noexcept
    {
        return m_errorCode;
    }

    [[nodiscard]] ConversionType GetConversionType() const noexcept
    {
        return m_conversionType;
    }

private:
    DWORD m_errorCode;
    ConversionType m_conversionType;
};


//------------------------------------------------------------------------------
// Helper function to safely convert a size_t value to int.
// If size_t is too large, throws a std::overflow_error.
//------------------------------------------------------------------------------
inline [[nodiscard]] int SafeCastSizeToInt(size_t s)
{
    using DestinationType = int;

    if (s > static_cast<size_t>((std::numeric_limits<DestinationType>::max)()))
    {
        throw std::overflow_error(
            "Input size is too long: size_t-length doesn't fit into int.");
    }

    return static_cast<DestinationType>(s);
}


//------------------------------------------------------------------------------
// Convert from UTF-16 std::wstring_view to UTF-8 std::string.
// Signal errors throwing UnicodeConversionException.
//------------------------------------------------------------------------------
inline [[nodiscard]] std::string Utf16ToUtf8(std::wstring_view utf16)
{
    // Special case of empty input string
    if (utf16.empty())
    {
        // Empty input --> return empty output string
        return std::string{};
    }

    // Safely fail if an invalid UTF-16 character sequence is encountered
    constexpr DWORD kFlags = WC_ERR_INVALID_CHARS;

    const int utf16Length = SafeCastSizeToInt(utf16.length());

    // Get the length, in chars, of the resulting UTF-8 string
    const int utf8Length = ::WideCharToMultiByte(
        CP_UTF8,            // convert to UTF-8
        kFlags,             // conversion flags
        utf16.data(),       // source UTF-16 string
        utf16Length,        // length of source UTF-16 string, in wchar_ts
        nullptr,            // unused - no conversion required in this step
        0,                  // request size of destination buffer, in chars
        nullptr, nullptr    // unused
    );
    if (utf8Length == 0)
    {
        // Conversion error: capture error code and throw
        const DWORD errorCode = ::GetLastError();
        throw UnicodeConversionException(
            errorCode,
            UnicodeConversionException::ConversionType::FromUtf16ToUtf8,
            "Can't get result UTF-8 string length (WideCharToMultiByte failed).");
    }

    // Make room in the destination string for the converted bits
    std::string utf8(utf8Length, ' ');
    char* utf8Buffer = utf8.data();
    _ASSERTE(utf8Buffer != nullptr);

    // Do the actual conversion from UTF-16 to UTF-8
    int result = ::WideCharToMultiByte(
        CP_UTF8,            // convert to UTF-8
        kFlags,             // conversion flags
        utf16.data(),       // source UTF-16 string
        utf16Length,        // length of source UTF-16 string, in wchar_ts
        utf8Buffer,         // pointer to destination buffer
        utf8Length,         // size of destination buffer, in chars
        nullptr, nullptr    // unused
    );
    if (result == 0)
    {
        // Conversion error: capture error code and throw
        const DWORD errorCode = ::GetLastError();
        throw UnicodeConversionException(
            errorCode,
            UnicodeConversionException::ConversionType::FromUtf16ToUtf8,
            "Can't convert from UTF-16 to UTF-8 string (WideCharToMultiByte failed).");
    }

    return utf8;
}


//------------------------------------------------------------------------------
// Convert from UTF-8 std::string_view to UTF-16 std::wstring.
// Signal errors throwing UnicodeConversionException.
//------------------------------------------------------------------------------
inline [[nodiscard]] std::wstring Utf8ToUtf16(std::string_view utf8)
{
    // Special case of empty input string
    if (utf8.empty())
    {
        // Empty input --> return empty output string
        return std::wstring{};
    }

    // Safely fail if an invalid UTF-8 character sequence is encountered
    constexpr DWORD kFlags = MB_ERR_INVALID_CHARS;

    const int utf8Length = SafeCastSizeToInt(utf8.length());

    // Get the size of the destination UTF-16 string
    const int utf16Length = ::MultiByteToWideChar(
        CP_UTF8,       // source string is in UTF-8
        kFlags,        // conversion flags
        utf8.data(),   // source UTF-8 string pointer
        utf8Length,    // length of the source UTF-8 string, in chars
        nullptr,       // unused - no conversion done in this step
        0              // request size of destination buffer, in wchar_ts
    );
    if (utf16Length == 0)
    {
        // Conversion error: capture error code and throw
        const DWORD errorCode = ::GetLastError();
        throw UnicodeConversionException(
            errorCode,
            UnicodeConversionException::ConversionType::FromUtf8ToUtf16,
            "Can't get result UTF-16 string length (MultiByteToWideChar failed).");
    }

    // Make room in the destination string for the converted bits
    std::wstring utf16(utf16Length, L' ');
    wchar_t* utf16Buffer = utf16.data();
    _ASSERTE(utf16Buffer != nullptr);

    // Do the actual conversion from UTF-8 to UTF-16
    int result = ::MultiByteToWideChar(
        CP_UTF8,       // source string is in UTF-8
        kFlags,        // conversion flags
        utf8.data(),   // source UTF-8 string pointer
        utf8Length,    // length of source UTF-8 string, in chars
        utf16Buffer,   // pointer to destination buffer
        utf16Length    // size of destination buffer, in wchar_ts
    );
    if (result == 0)
    {
        // Conversion error: capture error code and throw
        const DWORD errorCode = ::GetLastError();
        throw UnicodeConversionException(
            errorCode,
            UnicodeConversionException::ConversionType::FromUtf8ToUtf16,
            "Can't convert from UTF-8 to UTF-16 string (MultiByteToWideChar failed).");
    }

    return utf16;
}


} // namespace winreg::details

#endif // GIOVANNI_DICANIO_WINREG_UTF8CONV_HPP_INCLUDED

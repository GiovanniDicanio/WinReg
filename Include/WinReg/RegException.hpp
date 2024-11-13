////////////////////////////////////////////////////////////////////////////////
// FILE: RegException.hpp
// DESC: Exception class indicating errors with registry operations.
// AUTHOR: Giovanni Dicanio
////////////////////////////////////////////////////////////////////////////////

#ifndef GIOVANNI_DICANIO_WINREG_REGEXCEPTION_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_REGEXCEPTION_HPP_INCLUDED


#include <Windows.h>            // Windows API

#include <string>               // std::string
#include <system_error>         // std::system_error


namespace winreg
{

//------------------------------------------------------------------------------
// An exception representing an error with the registry operations
//------------------------------------------------------------------------------
class RegException
    : public std::system_error
{
public:
    RegException(LSTATUS errorCode, const char* message);
    RegException(LSTATUS errorCode, const std::string& message);
};



//------------------------------------------------------------------------------
//                       RegException Inline Methods
//------------------------------------------------------------------------------

inline RegException::RegException(const LSTATUS errorCode, const char* const message)
    : std::system_error{ errorCode, std::system_category(), message }
{}


inline RegException::RegException(const LSTATUS errorCode, const std::string& message)
    : std::system_error{ errorCode, std::system_category(), message }
{}


} // namespace winreg

#endif // GIOVANNI_DICANIO_WINREG_REGEXCEPTION_HPP_INCLUDED

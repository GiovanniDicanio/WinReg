////////////////////////////////////////////////////////////////////////////////
//
//         *** Modern C++ Wrappers Around Windows Registry C API ***
//
//       Copyright (C) by Giovanni Dicanio (giovanni.dicanio@gmail.com)
//
//----------------------------------------------------------------------------
// SPDX-License-Identifier: MIT
////////////////////////////////////////////////////////////////////////////////


#ifndef GIOVANNI_DICANIO_WINREG_REGEXCEPTION_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_REGEXCEPTION_HPP_INCLUDED


#include <Windows.h>        // Windows Platform SDK

#include <string>           // std::string
#include <system_error>     // std::system_error


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
//                  RegException Inline Method Implementation
//------------------------------------------------------------------------------

inline RegException::RegException(const LSTATUS errorCode, const char* const message)
    : std::system_error{ errorCode, std::system_category(), message }
{}


inline RegException::RegException(const LSTATUS errorCode, const std::string& message)
    : std::system_error{ errorCode, std::system_category(), message }
{}


} // namespace winreg


#endif // GIOVANNI_DICANIO_WINREG_REGEXCEPTION_HPP_INCLUDED

////////////////////////////////////////////////////////////////////////////////
//
//         *** Modern C++ Wrappers Around Windows Registry C API ***
//
//       Copyright (C) by Giovanni Dicanio (giovanni.dicanio@gmail.com)
//
//----------------------------------------------------------------------------
// SPDX-License-Identifier: MIT
////////////////////////////////////////////////////////////////////////////////


#ifndef GIOVANNI_DICANIO_WINREG_INTERNAL_SAFECAST_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_INTERNAL_SAFECAST_HPP_INCLUDED


#include <Windows.h>        // Windows Platform SDK

#include <stdexcept>        // std::overflow_error
#include <utility>          // std::in_range


namespace winreg
{

namespace winreg_internal
{


//------------------------------------------------------------------------------
// Safely cast a size_t value (usually from the STL)
// to a DWORD (usually for Win32 API calls).
// In case of overflow, throws an exception of type std::overflow_error.
//------------------------------------------------------------------------------
[[nodiscard]] inline DWORD SafeCastSizeToDword(const size_t size)
{
    // Check against overflow
    if (!std::in_range<DWORD>(size))
    {
        throw std::overflow_error(
            "Input size_t value is too big: size_t value doesn't fit into a DWORD.");
    }

    return static_cast<DWORD>(size);
}


} // namespace winreg_internal

} // namespace winreg


#endif // GIOVANNI_DICANIO_WINREG_INTERNAL_SAFECAST_HPP_INCLUDED

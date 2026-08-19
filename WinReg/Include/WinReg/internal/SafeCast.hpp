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
#include <crtdbg.h>         // _ASSERTE

#include <limits>           // std::numeric_limits
#include <stdexcept>        // std::overflow_error


namespace winreg
{

namespace winreg_internal
{


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


} // namespace winreg_internal

} // namespace winreg


#endif // GIOVANNI_DICANIO_WINREG_INTERNAL_SAFECAST_HPP_INCLUDED

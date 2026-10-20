////////////////////////////////////////////////////////////////////////////////
//
//         *** Modern C++ Wrappers Around Windows Registry C API ***
//
//       Copyright (C) by Giovanni Dicanio (giovanni.dicanio@gmail.com)
//
//----------------------------------------------------------------------------
// SPDX-License-Identifier: MIT
////////////////////////////////////////////////////////////////////////////////

#ifndef GIOVANNI_DICANIO_WINREG_DETAIL_CONFIG_HPP
#define GIOVANNI_DICANIO_WINREG_DETAIL_CONFIG_HPP


namespace winreg
{

namespace detail
{

    //
    // Maximum number of retries when attempting to read data from the registry.
    //
    // Each retry means the value grew between the size query and the read.
    // A few retries cover normal concurrent writes; more would only delay
    // the failure under continuous contention.
    //
    inline constexpr int kMaxReadAttempts = 10;

} // namespace detail

} // namespace winreg


#endif // GIOVANNI_DICANIO_WINREG_DETAIL_CONFIG_HPP

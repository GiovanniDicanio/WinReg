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


#include <expected>         // std::expected

#include "RegResult.hpp"


namespace winreg
{

    //------------------------------------------------------------------------------
    // A class template that stores a value of type T (e.g. DWORD, std::wstring)
    // on success, or a RegResult on error.
    //
    // Used as the return value of some Registry RegKey::TryGetXxxValue() methods
    // as an alternative to exception-throwing methods.
    //------------------------------------------------------------------------------

    template <typename T>
    using RegExpected = std::expected<T, RegResult>;

    //
    // NOTE: This was a custom class template in previous versions of WinReg;
    // now thanks to std::expected the code was simplified a lot!!
    //

} // namespace winreg

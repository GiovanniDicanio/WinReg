////////////////////////////////////////////////////////////////////////////////
//
//         *** Modern C++ Wrappers Around Windows Registry C API ***
//
//       Copyright (C) by Giovanni Dicanio (giovanni.dicanio@gmail.com)
//
// First version: 2017, January 22nd
// Last update:   2026, August 13th
//
//----------------------------------------------------------------------------
// SPDX-License-Identifier: MIT
////////////////////////////////////////////////////////////////////////////////


#ifndef GIOVANNI_DICANIO_WINREG_WINREG_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_WINREG_HPP_INCLUDED

////////////////////////////////////////////////////////////////////////////////
// Main public header - Include this in client code
////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////
//
// Registry key handles are safely and conveniently wrapped
// in the RegKey resource manager C++ class.
//
// Many methods are available in two forms:
//
// - One form that signals errors throwing exceptions
//   of class RegException (e.g. RegKey::Open)
//
// - Another form that returns RegResult objects (e.g. RegKey::TryOpen)
//
// In addition, there are also some methods named like TryGet...Value
// (e.g. TryGetDwordValue), that _try_ to perform the given query,
// and return a RegExpected object. On success, that object contains
// the value read from the registry. On failure, the returned RegExpected object
// contains a RegResult storing the return code from the Windows Registry API call.
//
// Unicode UTF-16 strings are represented using the std::wstring class;
// ATL's CString is not used, to avoid dependencies from ATL or MFC.
//
// IDE/Compiler: Visual Studio 2022
// C++ Language Standard: C++17 (/std:c++17)
// Code compiles cleanly at warning level 4 (/W4) on both 32-bit and 64-bit builds,
// also in C++20 mode (/std:c++20).
//
// Requires building in Unicode mode (which has been the default since VS2005).
//
// ===========================================================================
//
// The MIT License(MIT)
//
// Copyright(c) 2017-2026 by Giovanni Dicanio
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files(the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and / or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions :
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
//
////////////////////////////////////////////////////////////////////////////////


#include "WinReg/RegKey.hpp"            // The main class for the library

#include "WinReg/RegException.hpp"
#include "WinReg/RegExpected.hpp"
#include "WinReg/RegResult.hpp"


#endif // GIOVANNI_DICANIO_WINREG_WINREG_HPP_INCLUDED

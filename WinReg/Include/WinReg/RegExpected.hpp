////////////////////////////////////////////////////////////////////////////////
//
//         *** Modern C++ Wrappers Around Windows Registry C API ***
//
//       Copyright (C) by Giovanni Dicanio (giovanni.dicanio@gmail.com)
//
//----------------------------------------------------------------------------
// SPDX-License-Identifier: MIT
////////////////////////////////////////////////////////////////////////////////


#ifndef GIOVANNI_DICANIO_WINREG_REGEXPECTED_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_REGEXPECTED_HPP_INCLUDED


#include <Windows.h>        // Windows Platform SDK
#include <crtdbg.h>         // _ASSERTE

#include <variant>          // std::variant

#include "WinReg/RegResult.hpp"


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
class RegExpected
{
public:
    // Initialize the object with an error code
    explicit RegExpected(const RegResult& errorCode) noexcept;

    // Initialize the object with a value (the success case)
    explicit RegExpected(const T& value);

    // Initialize the object with a value (the success case),
    // optimized for move semantics
    explicit RegExpected(T&& value);

    // Does this object contain a valid value?
    [[nodiscard]] explicit operator bool() const noexcept;

    // Does this object contain a valid value?
    [[nodiscard]] bool IsValid() const noexcept;

    // Access the value (if the object contains a valid value).
    // Throws an exception if the object is in invalid state.
    [[nodiscard]] const T& GetValue() const;

    // Access the error code (if the object contains an error status)
    // Throws an exception if the object is in valid state.
    [[nodiscard]] RegResult GetError() const;



    // Builds a RegExpected object that stores an error code
    static [[nodiscard]] RegExpected<T> MakeRegExpectedWithError(const LSTATUS retCode);


private:
    // Stores a value of type T on success,
    // or RegResult on error
    std::variant<RegResult, T> m_var;
};




//------------------------------------------------------------------------------
//                   RegExpected Inline Method Implementation
//------------------------------------------------------------------------------

template <typename T>
inline RegExpected<T>::RegExpected(const RegResult& errorCode) noexcept
    : m_var{ errorCode }
{
}


template <typename T>
inline RegExpected<T>::RegExpected(const T& value)
    : m_var{ value }
{
}


template <typename T>
inline RegExpected<T>::RegExpected(T&& value)
    : m_var{ std::move(value) }
{
}


template <typename T>
inline RegExpected<T>::operator bool() const noexcept
{
    return IsValid();
}


template <typename T>
inline bool RegExpected<T>::IsValid() const noexcept
{
    return std::holds_alternative<T>(m_var);
}


template <typename T>
inline const T& RegExpected<T>::GetValue() const
{
    // Check that the object stores a valid value
    _ASSERTE(IsValid());

    // If the object is in a valid state, the variant stores an instance of T
    return std::get<T>(m_var);
}


template <typename T>
inline RegResult RegExpected<T>::GetError() const
{
    // Check that the object is in an invalid state
    _ASSERTE(!IsValid());

    // If the object is in an invalid state, the variant stores a RegResult
    // that represents an error code from the Windows Registry API
    return std::get<RegResult>(m_var);
}


template <typename T>
inline RegExpected<T> RegExpected<T>::MakeRegExpectedWithError(const LSTATUS retCode)
{
    return RegExpected<T>{ RegResult{ retCode } };
}


} // namespace winreg


#endif // GIOVANNI_DICANIO_WINREG_REGEXPECTED_HPP_INCLUDED

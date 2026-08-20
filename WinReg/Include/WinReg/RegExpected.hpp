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

#include <type_traits>      // std::is_same_v
#include <utility>          // std::move
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

    static_assert(
        !std::is_same_v<T, RegResult>,
        "RegExpected<T>: T must not be RegResult.");


    //
    // Factory functions for RegExpected
    //

    // Build a RegExpected storing a valid value
    static [[nodiscard]] RegExpected MakeSuccess(const T& value);

    // Build a RegExpected storing a valid value
    // (optimized for move semantics)
    static [[nodiscard]] RegExpected MakeSuccess(T&& value);

    // Build a RegExpected storing an error code
    static [[nodiscard]] RegExpected MakeError(RegResult error);



    // Does this object contain a valid value?
    [[nodiscard]] explicit operator bool() const noexcept;

    // Does this object contain a valid value?
    [[nodiscard]] bool IsValid() const noexcept;

    // Access the value (if the object contains a valid value).
    // Throws an exception if the object is in invalid state.
    [[nodiscard]] const T& GetValue() const &;

    // Access the value (if the object contains a valid value).
    // Throws an exception if the object is in invalid state.
    [[nodiscard]] T& GetValue() &;

    // Access the value (if the object contains a valid value).
    // Throws an exception if the object is in invalid state.
    [[nodiscard]] T&& GetValue() &&;

    // Access the value (if the object contains a valid value).
    // Throws an exception if the object is in invalid state.
    [[nodiscard]] const T&& GetValue() const &&;

    // Access the error code (if the object contains an error status)
    // Throws an exception if the object is in valid state.
    [[nodiscard]] RegResult GetError() const;


    // Helper function: Builds a RegExpected object that stores a RegResult
    // containing a Windows Registry API error code expressed as a "raw" LSTATUS.
    static [[nodiscard]] RegExpected<T> MakeRegExpectedWithError(LSTATUS retCode);


private:
    // Stores a value of type T on success,
    // or RegResult on error
    std::variant<T, RegResult> m_var;


    // Use tags to distinguish between the success and error cases
    struct ValueTag {};
    struct ErrorTag {};


    // Initialize the object with a value (the success case)
    RegExpected(ValueTag, const T& value);

    // Initialize the object with a value (the success case),
    // optimized for move semantics
    RegExpected(ValueTag, T&& value);

    // Initialize the object with an error code
    //
    // Note: RegResult is just a tiny wrapper around an LSTATUS (== LONG),
    // so there is no need to provide two overloads const T& and T&&
    // like for the "valid value" (ValueTag) success case.
    RegExpected(ErrorTag, RegResult errorCode);
};



//------------------------------------------------------------------------------
//                   RegExpected Inline Method Implementation
//------------------------------------------------------------------------------

template <typename T>
inline RegExpected<T>::RegExpected(ValueTag, const T& value)
    : m_var{ std::in_place_type<T>, value }
{
}


template <typename T>
inline RegExpected<T>::RegExpected(ValueTag, T&& value)
    : m_var{ std::in_place_type<T>, std::move(value) }
{
}


template <typename T>
inline RegExpected<T>::RegExpected(ErrorTag, RegResult errorCode)
    : m_var{ std::in_place_type<RegResult>, std::move(errorCode) }
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
inline const T& RegExpected<T>::GetValue() const &
{
    // Check that the object stores a valid value
    _ASSERTE(IsValid());

    // If the object is in a valid state, the variant stores an instance of T
    return std::get<T>(m_var);
}


template <typename T>
inline T& RegExpected<T>::GetValue() &
{
    // Check that the object stores a valid value
    _ASSERTE(IsValid());

    // If the object is in a valid state, the variant stores an instance of T
    return std::get<T>(m_var);
}


template <typename T>
inline T&& RegExpected<T>::GetValue() &&
{
    // Check that the object stores a valid value
    _ASSERTE(IsValid());

    // If the object is in a valid state, the variant stores an instance of T
    return std::get<T>(std::move(m_var));
}


template <typename T>
inline const T&& RegExpected<T>::GetValue() const &&
{
    // Check that the object stores a valid value
    _ASSERTE(IsValid());

    // If the object is in a valid state, the variant stores an instance of T
    return std::get<T>(std::move(m_var));
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
inline RegExpected<T> RegExpected<T>::MakeSuccess(const T& value)
{
    return RegExpected{ ValueTag{}, value };
}


template <typename T>
inline RegExpected<T> RegExpected<T>::MakeSuccess(T&& value)
{
    return RegExpected{ ValueTag{}, std::move(value) };
}


template <typename T>
inline RegExpected<T> RegExpected<T>::MakeError(RegResult error)
{
    return RegExpected{ ErrorTag{}, std::move(error) };

}


template <typename T>
inline RegExpected<T> RegExpected<T>::MakeRegExpectedWithError(const LSTATUS retCode)
{
    return RegExpected<T>::MakeError( RegResult{ retCode } );
}


} // namespace winreg


#endif // GIOVANNI_DICANIO_WINREG_REGEXPECTED_HPP_INCLUDED

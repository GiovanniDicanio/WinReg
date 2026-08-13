////////////////////////////////////////////////////////////////////////////////
//
//         *** Modern C++ Wrappers Around Windows Registry C API ***
//
//       Copyright (C) by Giovanni Dicanio (giovanni.dicanio@gmail.com)
//
//----------------------------------------------------------------------------
// SPDX-License-Identifier: MIT
////////////////////////////////////////////////////////////////////////////////


#ifndef GIOVANNI_DICANIO_WINREG_INTERNAL_SCOPEDLOCALFREE_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_INTERNAL_SCOPEDLOCALFREE_HPP_INCLUDED


#include <Windows.h>        // Windows Platform SDK


namespace winreg
{

namespace winreg_internal
{


//------------------------------------------------------------------------------
// Simple scoped-based RAII wrapper that *automatically* invokes LocalFree()
// in its destructor.
//------------------------------------------------------------------------------
template <typename T>
class ScopedLocalFree
{
public:

    using Type = T;
    using TypePtr = T*;


    // Init wrapped pointer to nullptr
    ScopedLocalFree() noexcept = default;

    // Automatically and safely invoke ::LocalFree()
    ~ScopedLocalFree() noexcept
    {
        Free();
    }

    //
    // Ban copy and move operations
    //
    ScopedLocalFree(const ScopedLocalFree&) = delete;
    ScopedLocalFree(ScopedLocalFree&&) = delete;
    ScopedLocalFree& operator=(const ScopedLocalFree&) = delete;
    ScopedLocalFree& operator=(ScopedLocalFree&&) = delete;


    // Read-only access to the wrapped pointer
    [[nodiscard]] T* Get() const noexcept
    {
        return m_ptr;
    }

    // Writable access to the wrapped pointer
    [[nodiscard]] T** AddressOf() noexcept
    {
        return &m_ptr;
    }

    // Explicit pointer conversion to bool
    explicit operator bool() const noexcept
    {
        return (m_ptr != nullptr);
    }

    // Safely invoke ::LocalFree() on the wrapped pointer
    void Free() noexcept
    {
        if (m_ptr != nullptr)
        {
            ::LocalFree(m_ptr);
            m_ptr = nullptr;
        }
    }


    //
    // IMPLEMENTATION
    //
private:
    T* m_ptr{ nullptr };
};


} // namespace winreg_internal

} // namespace winreg


#endif // GIOVANNI_DICANIO_WINREG_INTERNAL_SCOPEDLOCALFREE_HPP_INCLUDED

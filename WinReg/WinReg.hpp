#ifndef GIOVANNI_DICANIO_WINREG_HPP_INCLUDED
#define GIOVANNI_DICANIO_WINREG_HPP_INCLUDED

////////////////////////////////////////////////////////////////////////////////
//
//      *** Modern C++ Wrappers Around Windows Registry C API ***
//
//               Copyright (C) by Giovanni Dicanio
//
// First version: 2017, January 22nd
// Last update:   2020, June 11th
//
// E-mail: <first name>.<last name> AT REMOVE_THIS gmail.com
//
// Registry key handles are safely and conveniently wrapped
// in the RegKey resource manager C++ class.
//
// Errors are signaled throwing exceptions of class RegException.
// In addition, there are also some methods named like TryGet...
// (e.g. TryGetDwordValue), that _try_ to perform the given query,
// and return a std::optional value.
// (In particular, on failure, the returned std::optional object
// doesn't contain any value).
//
// Unicode UTF-16 strings are represented using the std::wstring class;
// ATL's CString is not used, to avoid dependencies from ATL or MFC.
//
// Compiler: Visual Studio 2019
// Code compiles cleanly at /W4 on both 32-bit and 64-bit builds.
//
// Requires building in Unicode mode (which is the default since VS2005).
//
// ===========================================================================
//
// The MIT License(MIT)
//
// Copyright(c) 2017-2020 by Giovanni Dicanio
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files(the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and / or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions :
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
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

#include <Windows.h>
#include <memory>
#include <optional>
#include <string>
#include <system_error>
#include <utility>
#include <vector>

#define ASSERT(...)
namespace winreg
{
    // Forward class declarations
    class RegException;
    template<typename>
    class RegResult;

    //------------------------------------------------------------------------------
    //                  Private Helper Classes and Functions
    //------------------------------------------------------------------------------

    namespace detail
    {
        //------------------------------------------------------------------------------
        // Simple scoped-based RAII wrapper that *automatically* invokes
        // LocalFree() in its destructor.
        //------------------------------------------------------------------------------
        template <typename T> class ScopedLocalFree
        {
                //
                // IMPLEMENTATION
                //
            private:
                T* m_ptr{ nullptr };

            public:
                typedef T Type;
                typedef T* TypePtr;

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
                T* Get() const noexcept
                {
                    return m_ptr;
                }

                // Writable access to the wrapped pointer
                T** AddressOf() noexcept
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
                    if(m_ptr != nullptr)
                    {
                        LocalFree(m_ptr);
                        m_ptr = nullptr;
                    }
                }
        };

        inline size_t stringlength(const char* s)
        {
            return std::strlen(s);
        }

        inline size_t stringlength(const wchar_t* s)
        {
            return wcslen(s);
        }

        //------------------------------------------------------------------------------
        // Helper function to build a multi-string from a vector<string>.
        //
        // A multi-string is a sequence of contiguous NUL-terminated strings,
        // that terminates with an additional NUL.
        // Basically, considered as a whole, the sequence is terminated by two
        // NULs. E.g.:
        //          Hello\0World\0\0
        //------------------------------------------------------------------------------
        template<typename CharT>
        inline std::vector<CharT> BuildMultiString(const std::vector<std::basic_string<CharT>>& data)
        {
            // Special case of the empty multi-string
            if(data.empty())
            {
                // Build a vector containing just two NULs
                return std::vector<CharT>(2, CharT('\0'));
            }

            // Get the total length in wchar_ts of the multi-string
            size_t totalLen = 0;
            for(const auto& s : data)
            {
                // Add one to current string's length for the terminating NUL
                totalLen += (s.length() + 1);
            }

            // Add one for the last NUL terminator (making the whole structure double-NUL terminated)
            totalLen++;

            // Allocate a buffer to store the multi-string
            std::vector<CharT> multiString;

            // Reserve room in the vector to speed up the following insertion loop
            multiString.reserve(totalLen);

            // Copy the single strings into the multi-string
            for(const auto& s : data)
            {
                multiString.insert(multiString.end(), s.begin(), s.end());

                // Don't forget to NUL-terminate the current string
                multiString.emplace_back(CharT('\0'));
            }

            // Add the last NUL-terminator
            multiString.emplace_back(CharT('\0'));

            return multiString;
        }

    }// namespace detail

    //------------------------------------------------------------------------------
    // Safe, efficient and convenient C++ wrapper around HKEY registry key handles.
    //
    // This class is movable but not copyable.
    //
    // This class is designed to be very *efficient* and low-overhead, for example:
    // non-throwing operations are carefully marked as noexcept, so the C++ compiler
    // can emit optimized code.
    //
    // Moreover, this class just wraps a raw HKEY handle, without any
    // shared-ownership overhead like in std::shared_ptr; you can think of this
    // class kind of like a std::unique_ptr for HKEYs.
    //
    // The class is also swappable (defines a custom non-member swap);
    // relational operators are properly overloaded as well.
    //------------------------------------------------------------------------------

    //------------------------------------------------------------------------------
    // An exception representing an error with the registry operations
    //------------------------------------------------------------------------------
    class RegException : public std::system_error
    {
        public:
            RegException(LONG errorCode, const char* message);
            RegException(LONG errorCode, const std::string& message);
    };

    //------------------------------------------------------------------------------
    // A tiny wrapper around LONG return codes used by the Windows Registry API.
    //------------------------------------------------------------------------------
    template<typename CharT>
    class RegResult
    {
        private:
            // Error code returned by Windows Registry C API;
            // default initialized to success code.
            LONG m_result{ ERROR_SUCCESS };

        public:
            // Initialize to success code (ERROR_SUCCESS)
            RegResult() noexcept = default;

            // Conversion constructor, *not* marked "explicit" on purpose,
            // allows easy and convenient conversion from Win32 API return code type
            // to this C++ wrapper.
            inline RegResult(const LONG result) noexcept : m_result{ result }
            {
            }

            // Is the wrapped code a success code?
            inline bool IsOk() const noexcept
            {
                return m_result == ERROR_SUCCESS;
            }

            // Is the wrapped error code a failure code?
            inline bool Failed() const noexcept
            {
                return m_result != ERROR_SUCCESS;
            }

            // Is the wrapped code a success code?
            inline operator bool() const noexcept
            {
                return IsOk();
            }

            // Get the wrapped Win32 code
            inline LONG Code() const noexcept
            {
                return m_result;
            }

            // Return the system error message associated to the current error code
            std::basic_string<CharT> ErrorMessage() const
            {
                return ErrorMessage(MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT));
            }

            // Return the system error message associated to the current error code,
            // using the given input language identifier
            std::basic_string<CharT> ErrorMessage(DWORD languageId) const
            {
                // Invoke FormatMessage() to retrieve the error message from Windows
                detail::ScopedLocalFree<CharT> messagePtr;
                DWORD retCode = FormatMessage(
                FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                nullptr, m_result, languageId,
                reinterpret_cast<const CharT*>(messagePtr.AddressOf()), 0, nullptr);
                if(retCode == 0)
                {
                    // FormatMessage failed: return an empty string
                    return std::basis_string<CharT>{};
                }
                // Safely copy the C-string returned by FormatMessage() into a
                // std::string object, and return it back to the caller.
                return std::basic_string<CharT>{ messagePtr.Get() };
            }
    };

    template<typename CharT>
    class RegKey
    {
        private:
            // The wrapped registry key handle
            HKEY m_hKey{ nullptr };

        public:
            enum class ExpandStringOption
            {
                DontExpand,
                Expand
            };


        public:
            //
            // Construction/Destruction
            //

            // Initialize as an empty key handle
            RegKey() noexcept = default;

            // Take ownership of the input key handle
            explicit RegKey(HKEY hKey) noexcept : m_hKey{ hKey }
            {
            }

            // Open the given registry key if it exists, else create a new key.
            // Uses default KEY_READ|KEY_WRITE access.
            // For finer grained control, call the Create() method overloads.
            // Throw RegException on failure.
            RegKey(HKEY hKeyParent, const std::basic_string<CharT>& subKey)
            {
                Create(hKeyParent, subKey);
            }

            // Open the given registry key if it exists, else create a new key.
            // Allow the caller to specify the desired access to the key (e.g.
            // KEY_READ for read-only access). For finer grained control, call the
            // Create() method overloads. Throw RegException on failure.
            RegKey(HKEY hKeyParent, const std::basic_string<CharT>& subKey, REGSAM desiredAccess)
            {
                Create(hKeyParent, subKey, desiredAccess);
            }

            // Take ownership of the input key handle.
            // The input key handle wrapper is reset to an empty state.
            RegKey(RegKey&& other) noexcept : m_hKey{ other.m_hKey }
            {
                // Other doesn't own the handle anymore
                other.m_hKey = nullptr;
            }

            // Move-assign from the input key handle.
            // Properly check against self-move-assign (which is safe and does nothing).
            RegKey& operator=(RegKey&& other) noexcept
            {
                // Prevent self-move-assign
                if((this != &other) && (m_hKey != other.m_hKey))
                {
                    // Close current
                    Close();

                    // Move from other (i.e. take ownership of other's raw handle)
                    m_hKey = other.m_hKey;
                    other.m_hKey = nullptr;
                }
                return *this;
            }

            // Ban copy
            RegKey(const RegKey&) = delete;
            RegKey& operator=(const RegKey&) = delete;

            // Safely close the wrapped key handle (if any)
            ~RegKey() noexcept
            {
                // Release the owned handle (if any)
                Close();
            }

            //
            // Properties
            //

            // Access the wrapped raw HKEY handle
            HKEY Get() const noexcept
            {
                return m_hKey;
            }

            // Is the wrapped HKEY handle valid?
            bool IsValid() const noexcept
            {
                return m_hKey != nullptr;
            }

            // Same as IsValid(), but allow a short "if (regKey)" syntax
            explicit operator bool() const noexcept
            {
                return IsValid();
            }

            // Is the wrapped handle a predefined handle (e.g.HKEY_CURRENT_USER) ?
            bool IsPredefined() const noexcept
            {
                // Predefined keys
                // https://msdn.microsoft.com/en-us/library/windows/desktop/ms724836(v=vs.85).aspx

                if((m_hKey == HKEY_CURRENT_USER) || (m_hKey == HKEY_LOCAL_MACHINE)
                   || (m_hKey == HKEY_CLASSES_ROOT) || (m_hKey == HKEY_CURRENT_CONFIG)
                   || (m_hKey == HKEY_CURRENT_USER_LOCAL_SETTINGS)
                   || (m_hKey == HKEY_PERFORMANCE_DATA) || (m_hKey == HKEY_PERFORMANCE_NLSTEXT)
                   || (m_hKey == HKEY_PERFORMANCE_TEXT) || (m_hKey == HKEY_USERS))
                {
                    return true;
                }

                return false;
            }

            //
            // Operations
            //

            // Close current HKEY handle.
            // If there's no valid handle, do nothing.
            // This method doesn't close predefined HKEY handles (e.g. HKEY_CURRENT_USER).
            void Close() noexcept
            {
                if(IsValid())
                {
                    // Do not call RegCloseKey on predefined keys
                    if(!IsPredefined())
                    {
                        RegCloseKey(m_hKey);
                    }

                    // Avoid dangling references
                    m_hKey = nullptr;
                }
            }

            // Transfer ownership of current HKEY to the caller.
            // Note that the caller is responsible for closing the key handle!
            HKEY Detach() noexcept
            {
                HKEY hKey = m_hKey;

                // We don't own the HKEY handle anymore
                m_hKey = nullptr;

                // Transfer ownership to the caller
                return hKey;
            }

            // Take ownership of the input HKEY handle.
            // Safely close any previously open handle.
            // Input key handle can be nullptr.
            void Attach(HKEY hKey) noexcept
            {
                // Prevent self-attach
                if(m_hKey != hKey)
                {
                    // Close any open registry handle
                    Close();

                    // Take ownership of the input hKey
                    m_hKey = hKey;
                }
            }

            // Non-throwing swap;
            // Note: There's also a non-member swap overload
            void SwapWith(RegKey& other) noexcept
            {
                // Enable ADL (not necessary in this case, but good practice)
                using std::swap;

                // Swap the raw handle members
                swap(m_hKey, other.m_hKey);
            }

            //
            // Wrappers around Windows Registry APIs.
            // See the official MSDN documentation for these APIs for detailed
            // explanations of the wrapper method parameters.
            //

            // Wrapper around RegCreateKeyEx, that allows you to specify desired access
            void Create(HKEY hKeyParent, const std::basic_string<CharT>& subKey, REGSAM desiredAccess = KEY_READ | KEY_WRITE)
            {
                constexpr DWORD kDefaultOptions = REG_OPTION_NON_VOLATILE;

                Create(hKeyParent, subKey, desiredAccess, kDefaultOptions,
                       nullptr,// no security attributes,
                       nullptr// no disposition
                );
            }

            // Wrapper around RegCreateKeyEx
            void Create(HKEY hKeyParent,
                        const std::basic_string<CharT>& subKey,
                        REGSAM desiredAccess,
                        DWORD options,
                        SECURITY_ATTRIBUTES* securityAttributes,
                        DWORD* disposition)
            {
                HKEY hKey = nullptr;
                LONG retCode = RegCreateKeyEx(
                hKeyParent, subKey.c_str(),
                0,// reserved
                REG_NONE,// user-defined class type parameter not supported
                options, desiredAccess, securityAttributes, &hKey, disposition);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegCreateKeyEx failed." };
                }

                // Safely close any previously opened key
                Close();

                // Take ownership of the newly created key
                m_hKey = hKey;
            }

            // Wrapper around RegOpenKeyEx
            void Open(HKEY hKeyParent, const std::basic_string<CharT>& subKey, REGSAM desiredAccess = KEY_READ | KEY_WRITE)
            {
                HKEY hKey = nullptr;
                LONG retCode = RegOpenKeyEx(hKeyParent, subKey.c_str(),
                                            REG_NONE,// default options
                                            desiredAccess, &hKey);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegOpenKeyEx failed." };
                }

                // Safely close any previously opened key
                Close();

                // Take ownership of the newly created key
                m_hKey = hKey;
            }

            // Wrapper around RegCreateKeyEx, that allows you to specify desired access
            RegResult<CharT> TryCreate(HKEY hKeyParent,
                                const std::basic_string<CharT>& subKey,
                                REGSAM desiredAccess = KEY_READ | KEY_WRITE) noexcept
            {
                constexpr DWORD kDefaultOptions = REG_OPTION_NON_VOLATILE;

                return TryCreate(hKeyParent, subKey, desiredAccess, kDefaultOptions,
                                 nullptr,// no security attributes,
                                 nullptr// no disposition
                );
            }

            // Wrapper around RegCreateKeyEx
            RegResult<CharT> TryCreate(HKEY hKeyParent,
                                const std::basic_string<CharT>& subKey,
                                REGSAM desiredAccess,
                                DWORD options,
                                SECURITY_ATTRIBUTES* securityAttributes,
                                DWORD* disposition) noexcept
            {
                HKEY hKey = nullptr;
                RegResult<CharT> retCode = RegCreateKeyEx(
                    hKeyParent, subKey.c_str(),
                    0,// reserved
                    REG_NONE,// user-defined class type parameter not supported
                    options, desiredAccess, securityAttributes, &hKey, disposition
                );
                if(retCode.Failed())
                {
                    return retCode;
                }
                // Safely close any previously opened key
                Close();
                // Take ownership of the newly created key
                m_hKey = hKey;
                ASSERT(retCode.IsOk());
                return retCode;
            }

            // Wrapper around RegOpenKeyEx
            RegResult<CharT> TryOpen(HKEY hKeyParent,
                              const std::basic_string<CharT>& subKey,
                              REGSAM desiredAccess = KEY_READ | KEY_WRITE) noexcept
            {
                HKEY hKey = nullptr;
                RegResult<CharT> retCode = RegOpenKeyEx(
                    hKeyParent, subKey.c_str(),
                    REG_NONE,// default options
                    desiredAccess, &hKey
                );
                if(retCode.Failed())
                {
                    return retCode;
                }
                // Safely close any previously opened key
                Close();
                // Take ownership of the newly created key
                m_hKey = hKey;
                ASSERT(retCode.IsOk());
                return retCode;
            }

            //
            // Registry Value Setters
            //

            void SetDwordValue(const std::basic_string<CharT>& valueName, DWORD data)
            {
                ASSERT(IsValid());
                LONG retCode = RegSetValueEx(
                    m_hKey, valueName.c_str(),
                    0,// reserved
                    REG_DWORD, reinterpret_cast<const BYTE*>(&data),
                    sizeof(data));
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot write DWORD value: RegSetValueEx failed." };
                }
            }

            void SetQwordValue(const std::basic_string<CharT>& valueName, const ULONGLONG& data)
            {
                ASSERT(IsValid());

                LONG retCode = RegSetValueEx(m_hKey, valueName.c_str(),
                                             0,// reserved
                                             REG_QWORD, reinterpret_cast<const BYTE*>(&data),
                                             sizeof(data));
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot write QWORD value: RegSetValueEx failed." };
                }
            }

            void SetStringValue(const std::basic_string<CharT>& valueName, const std::basic_string<CharT>& data)
            {
                ASSERT(IsValid());

                // String size including the terminating NUL, in bytes
                const DWORD dataSize = static_cast<DWORD>((data.length() + 1) * sizeof(CharT));

                LONG retCode = RegSetValueEx(
                    m_hKey, valueName.c_str(),
                    0,// reserved
                    REG_SZ, reinterpret_cast<const BYTE*>(data.c_str()), dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot write string value: RegSetValueEx failed." };
                }
            }
            void SetExpandStringValue(const std::basic_string<CharT>& valueName, const std::basic_string<CharT>& data)
            {
                ASSERT(IsValid());

                // String size including the terminating NUL, in bytes
                const DWORD dataSize
                = static_cast<DWORD>((data.length() + 1) * sizeof(CharT));

                LONG retCode = RegSetValueEx(
                m_hKey, valueName.c_str(),
                0,// reserved
                REG_EXPAND_SZ, reinterpret_cast<const BYTE*>(data.c_str()), dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot write expand string value: RegSetValueEx failed." };
                }
            }

            void SetMultiStringValue(const std::basic_string<CharT>& valueName,
                                     const std::vector<std::basic_string<CharT>>& data)
            {
                ASSERT(IsValid());

                // First, we have to build a double-NUL-terminated multi-string from the input data
                const std::vector<CharT> multiString = detail::BuildMultiString(data);

                // Total size, in bytes, of the whole multi-string structure
                const DWORD dataSize = static_cast<DWORD>(multiString.size() * sizeof(CharT));

                LONG retCode = RegSetValueEx(
                    m_hKey, valueName.c_str(),
                    0,// reserved
                    REG_MULTI_SZ, reinterpret_cast<const BYTE*>(&multiString[0]), dataSize
                );
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot write multi-string value: RegSetValueEx failed." };
                }
            }

            void SetBinaryValue(const std::basic_string<CharT>& valueName, const std::vector<BYTE>& data)
            {
                ASSERT(IsValid());

                // Total data size, in bytes
                const DWORD dataSize = static_cast<DWORD>(data.size());

                LONG retCode = RegSetValueEx(m_hKey, valueName.c_str(),
                                             0,// reserved
                                             REG_BINARY, &data[0], dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot write binary data value: RegSetValueEx failed." };
                }
            }

            void SetBinaryValue(const std::basic_string<CharT>& valueName, const void* data, DWORD dataSize)
            {
                ASSERT(IsValid());

                LONG retCode
                = RegSetValueEx(m_hKey, valueName.c_str(),
                                0,// reserved
                                REG_BINARY, static_cast<const BYTE*>(data), dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot write binary data value: RegSetValueEx failed." };
                }
            }

            //
            // Registry Value Getters
            //

            DWORD GetDwordValue(const std::basic_string<CharT>& valueName) const
            {
                ASSERT(IsValid());

                DWORD data = 0;// to be read from the registry
                DWORD dataSize = sizeof(data);// size of data, in bytes

                constexpr DWORD flags = RRF_RT_REG_DWORD;
                LONG retCode = RegGetValue(m_hKey,
                                           nullptr,// no subkey
                                           valueName.c_str(), flags,
                                           nullptr,// type not required
                                           &data, &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot get DWORD value: RegGetValue failed." };
                }

                return data;
            }

            ULONGLONG GetQwordValue(const std::basic_string<CharT>& valueName) const
            {
                ASSERT(IsValid());

                ULONGLONG data = 0;// to be read from the registry
                DWORD dataSize = sizeof(data);// size of data, in bytes

                constexpr DWORD flags = RRF_RT_REG_QWORD;
                LONG retCode = RegGetValue(m_hKey,
                                           nullptr,// no subkey
                                           valueName.c_str(), flags,
                                           nullptr,// type not required
                                           &data, &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot get QWORD value: RegGetValue failed." };
                }

                return data;
            }

            std::basic_string<CharT> GetStringValue(const std::basic_string<CharT>& valueName) const
            {
                ASSERT(IsValid());

                // Get the size of the result string
                DWORD dataSize = 0;// size of data, in bytes
                constexpr DWORD flags = RRF_RT_REG_SZ;
                LONG retCode = RegGetValue(m_hKey,
                                           nullptr,// no subkey
                                           valueName.c_str(), flags,
                                           nullptr,// type not required
                                           nullptr,// output buffer not needed now
                                           &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot get size of string value: RegGetValue failed." };
                }

                // Allocate a string of proper size.
                // Note that dataSize is in bytes and includes the terminating NUL;
                // we have to convert the size from bytes to wchar_ts for string::resize.
                std::basic_string<CharT> result(dataSize / sizeof(CharT), CharT(' '));

                // Call RegGetValue for the second time to read the string's content
                retCode = RegGetValue(m_hKey,
                                      nullptr,// no subkey
                                      valueName.c_str(), flags,
                                      nullptr,// type not required
                                      &result[0],// output buffer
                                      &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot get string value: RegGetValue failed." };
                }

                // Remove the NUL terminator scribbled by RegGetValue from the string
                result.resize((dataSize / sizeof(CharT)) - 1);

                return result;
            }

            ExpandStringOption x;
            std::basic_string<CharT> GetExpandStringValue(
                const std::basic_string<CharT>& valueName,
                ExpandStringOption expandOption= ExpandStringOption::DontExpand) const
            {
                ASSERT(IsValid());

                DWORD flags = RRF_RT_REG_EXPAND_SZ;

                // Adjust the flag for RegGetValue considering the expand string option specified by the caller
                if(expandOption == ExpandStringOption::DontExpand)
                {
                    flags |= RRF_NOEXPAND;
                }

                // Get the size of the result string
                DWORD dataSize = 0;// size of data, in bytes
                LONG retCode = RegGetValue(m_hKey,
                                           nullptr,// no subkey
                                           valueName.c_str(), flags,
                                           nullptr,// type not required
                                           nullptr,// output buffer not needed now
                                           &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot get size of expand string value: RegGetValue failed." };
                }

                // Allocate a string of proper size.
                // Note that dataSize is in bytes and includes the terminating NUL.
                // We must convert from bytes to wchar_ts for string::resize.
                std::basic_string<CharT> result(dataSize / sizeof(CharT), CharT(' '));

                // Call RegGetValue for the second time to read the string's content
                retCode = RegGetValue(m_hKey,
                                      nullptr,// no subkey
                                      valueName.c_str(), flags,
                                      nullptr,// type not required
                                      &result[0],// output buffer
                                      &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot get expand string value: RegGetValue failed." };
                }

                // Remove the NUL terminator scribbled by RegGetValue from the string
                result.resize((dataSize / sizeof(CharT)) - 1);

                return result;
            }

            std::vector<std::basic_string<CharT>> GetMultiStringValue(const std::basic_string<CharT>& valueName) const
            {
                ASSERT(IsValid());

                // Request the size of the multi-string, in bytes
                DWORD dataSize = 0;
                constexpr DWORD flags = RRF_RT_REG_MULTI_SZ;
                LONG retCode = RegGetValue(m_hKey,
                                           nullptr,// no subkey
                                           valueName.c_str(), flags,
                                           nullptr,// type not required
                                           nullptr,// output buffer not needed now
                                           &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot get size of multi-string value: RegGetValue failed." };
                }

                // Allocate room for the result multi-string.
                // Note that dataSize is in bytes, but our vector<wchar_t>::resize
                // method requires size to be expressed in wchar_ts.
                std::vector<CharT> data(dataSize / sizeof(CharT), CharT(' '));

                // Read the multi-string from the registry into the vector object
                retCode = RegGetValue(m_hKey,
                                      nullptr,// no subkey
                                      valueName.c_str(), flags,
                                      nullptr,// no type required
                                      &data[0],// output buffer
                                      &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot get multi-string value: RegGetValue failed." };
                }

                // Resize vector to the actual size returned by GetRegValue.
                // Note that the vector is a vector of wchar_ts, instead the size returned by GetRegValue
                // is in bytes, so we have to scale from bytes to wchar_t count.
                data.resize(dataSize / sizeof(CharT));

                // Parse the double-NUL-terminated string into a vector<string>,
                // which will be returned to the caller
                std::vector<std::basic_string<CharT>> result;
                const CharT* currStringPtr = &data[0];
                while(*currStringPtr != CharT('\0'))
                {
                    // Current string is NUL-terminated, so get its length calling wcslen
                    const size_t currStringLength = detail::stringlength(currStringPtr);

                    // Add current string to the result vector
                    result.emplace_back(currStringPtr, currStringLength);

                    // Move to the next string
                    currStringPtr += currStringLength + 1;
                }

                return result;
            }

            std::vector<BYTE> GetBinaryValue(const std::basic_string<CharT>& valueName) const
            {
                ASSERT(IsValid());

                // Get the size of the binary data
                DWORD dataSize = 0;// size of data, in bytes
                constexpr DWORD flags = RRF_RT_REG_BINARY;
                LONG retCode = RegGetValue(m_hKey,
                                           nullptr,// no subkey
                                           valueName.c_str(), flags,
                                           nullptr,// type not required
                                           nullptr,// output buffer not needed now
                                           &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot get size of binary data: RegGetValue failed." };
                }

                // Allocate a buffer of proper size to store the binary data
                std::vector<BYTE> data(dataSize);

                // Call RegGetValue for the second time to read the data content
                retCode = RegGetValue(m_hKey,
                                      nullptr,// no subkey
                                      valueName.c_str(), flags,
                                      nullptr,// type not required
                                      &data[0],// output buffer
                                      &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot get binary data: RegGetValue failed." };
                }

                return data;
            }

            //
            // Registry Value Getters Returning std::optional
            // (instead of throwing RegException on error)
            //

            std::optional<DWORD> TryGetDwordValue(const std::basic_string<CharT>& valueName) const
            {
                ASSERT(IsValid());

                DWORD data = 0;// to be read from the registry
                DWORD dataSize = sizeof(data);// size of data, in bytes

                constexpr DWORD flags = RRF_RT_REG_DWORD;
                LONG retCode = RegGetValue(m_hKey,
                                           nullptr,// no subkey
                                           valueName.c_str(), flags,
                                           nullptr,// type not required
                                           &data, &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    return {};
                }

                return data;
            }

            std::optional<ULONGLONG> TryGetQwordValue(const std::basic_string<CharT>& valueName) const
            {
                ASSERT(IsValid());

                ULONGLONG data = 0;// to be read from the registry
                DWORD dataSize = sizeof(data);// size of data, in bytes

                constexpr DWORD flags = RRF_RT_REG_QWORD;
                LONG retCode = RegGetValue(m_hKey,
                                           nullptr,// no subkey
                                           valueName.c_str(), flags,
                                           nullptr,// type not required
                                           &data, &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    return {};
                }

                return data;
            }

            std::optional<std::basic_string<CharT>> TryGetStringValue(const std::basic_string<CharT>& valueName) const
            {
                ASSERT(IsValid());

                // Get the size of the result string
                DWORD dataSize = 0;// size of data, in bytes
                constexpr DWORD flags = RRF_RT_REG_SZ;
                LONG retCode = RegGetValue(m_hKey,
                                           nullptr,// no subkey
                                           valueName.c_str(), flags,
                                           nullptr,// type not required
                                           nullptr,// output buffer not needed now
                                           &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    return {};
                }

                // Allocate a string of proper size.
                // Note that dataSize is in bytes and includes the terminating NUL;
                // we have to convert the size from bytes to wchar_ts for string::resize.
                std::basic_string<CharT> result(dataSize / sizeof(CharT), CharT(' '));

                // Call RegGetValue for the second time to read the string's content
                retCode = RegGetValue(m_hKey,
                                      nullptr,// no subkey
                                      valueName.c_str(), flags,
                                      nullptr,// type not required
                                      &result[0],// output buffer
                                      &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    return {};
                }

                // Remove the NUL terminator scribbled by RegGetValue from the string
                result.resize((dataSize / sizeof(CharT)) - 1);

                return result;
            }

            std::optional<std::basic_string<CharT>>
            TryGetExpandStringValue(const std::basic_string<CharT>& valueName,
                                    ExpandStringOption expandOption
                                    = ExpandStringOption::DontExpand) const
            {
                ASSERT(IsValid());

                DWORD flags = RRF_RT_REG_EXPAND_SZ;

                // Adjust the flag for RegGetValue considering the expand string option specified by the caller
                if(expandOption == ExpandStringOption::DontExpand)
                {
                    flags |= RRF_NOEXPAND;
                }

                // Get the size of the result string
                DWORD dataSize = 0;// size of data, in bytes
                LONG retCode = RegGetValue(m_hKey,
                                           nullptr,// no subkey
                                           valueName.c_str(), flags,
                                           nullptr,// type not required
                                           nullptr,// output buffer not needed now
                                           &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    return {};
                }

                // Allocate a string of proper size.
                // Note that dataSize is in bytes and includes the terminating NUL.
                // We must convert from bytes to wchar_ts for string::resize.
                std::basic_string<CharT> result(dataSize / sizeof(CharT), CharT(' '));

                // Call RegGetValue for the second time to read the string's content
                retCode = RegGetValue(m_hKey,
                                      nullptr,// no subkey
                                      valueName.c_str(), flags,
                                      nullptr,// type not required
                                      &result[0],// output buffer
                                      &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    return {};
                }

                // Remove the NUL terminator scribbled by RegGetValue from the string
                result.resize((dataSize / sizeof(CharT)) - 1);

                return result;
            }

            std::optional<std::vector<std::basic_string<CharT>>>
            TryGetMultiStringValue(const std::basic_string<CharT>& valueName) const
            {
                ASSERT(IsValid());

                // Request the size of the multi-string, in bytes
                DWORD dataSize = 0;
                constexpr DWORD flags = RRF_RT_REG_MULTI_SZ;
                LONG retCode = RegGetValue(m_hKey,
                                           nullptr,// no subkey
                                           valueName.c_str(), flags,
                                           nullptr,// type not required
                                           nullptr,// output buffer not needed now
                                           &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    return {};
                }

                // Allocate room for the result multi-string.
                // Note that dataSize is in bytes, but our vector<wchar_t>::resize
                // method requires size to be expressed in wchar_ts.
                std::vector<CharT> data(dataSize / sizeof(CharT), CharT(' '));

                // Read the multi-string from the registry into the vector object
                retCode = RegGetValue(m_hKey,
                                      nullptr,// no subkey
                                      valueName.c_str(), flags,
                                      nullptr,// no type required
                                      &data[0],// output buffer
                                      &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    return {};
                }

                // Resize vector to the actual size returned by GetRegValue.
                // Note that the vector is a vector of wchar_ts, instead the size returned by GetRegValue
                // is in bytes, so we have to scale from bytes to wchar_t count.
                data.resize(dataSize / sizeof(CharT));

                // Parse the double-NUL-terminated string into a vector<string>,
                // which will be returned to the caller
                std::vector<std::basic_string<CharT>> result;
                const CharT* currStringPtr = &data[0];
                while(*currStringPtr != CharT('\0'))
                {
                    // Current string is NUL-terminated, so get its length calling wcslen
                    const size_t currStringLength = detail::stringlength(currStringPtr);

                    // Add current string to the result vector
                    result.emplace_back(currStringPtr, currStringLength);

                    // Move to the next string
                    currStringPtr += currStringLength + 1;
                }

                return result;
            }

            std::optional<std::vector<BYTE>> TryGetBinaryValue(const std::basic_string<CharT>& valueName) const
            {
                ASSERT(IsValid());

                // Get the size of the binary data
                DWORD dataSize = 0;// size of data, in bytes
                constexpr DWORD flags = RRF_RT_REG_BINARY;
                LONG retCode = RegGetValue(m_hKey,
                                           nullptr,// no subkey
                                           valueName.c_str(), flags,
                                           nullptr,// type not required
                                           nullptr,// output buffer not needed now
                                           &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    return {};
                }

                // Allocate a buffer of proper size to store the binary data
                std::vector<BYTE> data(dataSize);

                // Call RegGetValue for the second time to read the data content
                retCode = RegGetValue(m_hKey,
                                      nullptr,// no subkey
                                      valueName.c_str(), flags,
                                      nullptr,// type not required
                                      &data[0],// output buffer
                                      &dataSize);
                if(retCode != ERROR_SUCCESS)
                {
                    return {};
                }

                return data;
            }

            //
            // Query Operations
            //

            void QueryInfoKey(DWORD& subKeys, DWORD& values, FILETIME& lastWriteTime) const
            {
                ASSERT(IsValid());

                subKeys = 0;
                values = 0;
                lastWriteTime.dwLowDateTime = lastWriteTime.dwHighDateTime = 0;

                LONG retCode = RegQueryInfoKey(m_hKey, nullptr, nullptr, nullptr, &subKeys,
                                               nullptr, nullptr, &values, nullptr,
                                               nullptr, nullptr, &lastWriteTime);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegQueryInfoKey failed." };
                }
            }

            // Return the DWORD type ID for the input registry value
            DWORD QueryValueType(const std::basic_string<CharT>& valueName) const
            {
                ASSERT(IsValid());

                DWORD typeId = 0;// will be returned by RegQueryValueEx

                LONG retCode = RegQueryValueEx(m_hKey, valueName.c_str(),
                                               nullptr,// reserved
                                               &typeId,
                                               nullptr,// not interested
                                               nullptr// not interested
                );

                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "Cannot get the value type: RegQueryValueEx failed." };
                }

                return typeId;
            }

            // Enumerate the subkeys of the registry key, using RegEnumKeyEx
            std::vector<std::basic_string<CharT>> EnumSubKeys() const
            {
                ASSERT(IsValid());

                // Get some useful enumeration info, like the total number of
                // subkeys and the maximum length of the subkey names
                DWORD subKeyCount = 0;
                DWORD maxSubKeyNameLen = 0;
                LONG retCode = RegQueryInfoKey(m_hKey,
                                               nullptr,// no user-defined class
                                               nullptr,// no user-defined class size
                                               nullptr,// reserved
                                               &subKeyCount, &maxSubKeyNameLen,
                                               nullptr,// no subkey class length
                                               nullptr,// no value count
                                               nullptr,// no value name max length
                                               nullptr,// no max value length
                                               nullptr,// no security descriptor
                                               nullptr// no last write time
                );
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegQueryInfoKey failed while preparing for subkey enumeration." };
                }

                // NOTE: According to the MSDN documentation, the size returned for subkey name max length
                // does *not* include the terminating NUL, so let's add +1 to take it into account
                // when I allocate the buffer for reading subkey names.
                maxSubKeyNameLen++;

                // Preallocate a buffer for the subkey names
                auto nameBuffer = std::make_unique<CharT[]>(maxSubKeyNameLen);

                // The result subkey names will be stored here
                std::vector<std::basic_string<CharT>> subkeyNames;

                // Reserve room in the vector to speed up the following insertion loop
                subkeyNames.reserve(subKeyCount);

                // Enumerate all the subkeys
                for(DWORD index = 0; index < subKeyCount; index++)
                {
                    // Get the name of the current subkey
                    DWORD subKeyNameLen = maxSubKeyNameLen;
                    retCode = RegEnumKeyEx(m_hKey, index, nameBuffer.get(), &subKeyNameLen,
                                           nullptr,// reserved
                                           nullptr,// no class
                                           nullptr,// no class
                                           nullptr// no last write time
                    );
                    if(retCode != ERROR_SUCCESS)
                    {
                        throw RegException{ retCode, "Cannot enumerate subkeys: RegEnumKeyEx failed." };
                    }

                    // On success, the ::RegEnumKeyEx API writes the length of the
                    // subkey name in the subKeyNameLen output parameter
                    // (not including the terminating NUL).
                    // So I can build a string based on that length.
                    subkeyNames.emplace_back(nameBuffer.get(), subKeyNameLen);
                }

                return subkeyNames;
            }

            // Enumerate the values under the registry key, using RegEnumValue.
            // Returns a vector of pairs: In each pair, the string is the value
            // name, the DWORD is the value type.
            std::vector<std::pair<std::basic_string<CharT>, DWORD>> EnumValues() const
            {
                ASSERT(IsValid());

                // Get useful enumeration info, like the total number of values
                // and the maximum length of the value names
                DWORD valueCount = 0;
                DWORD maxValueNameLen = 0;
                LONG retCode = RegQueryInfoKey(m_hKey,
                                               nullptr,// no user-defined class
                                               nullptr,// no user-defined class size
                                               nullptr,// reserved
                                               nullptr,// no subkey count
                                               nullptr,// no subkey max length
                                               nullptr,// no subkey class length
                                               &valueCount, &maxValueNameLen,
                                               nullptr,// no max value length
                                               nullptr,// no security descriptor
                                               nullptr// no last write time
                );
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegQueryInfoKey failed while preparing for value enumeration." };
                }

                // NOTE: According to the MSDN documentation, the size returned for value name max length
                // does *not* include the terminating NUL, so let's add +1 to take it into account
                // when I allocate the buffer for reading value names.
                maxValueNameLen++;

                // Preallocate a buffer for the value names
                auto nameBuffer = std::make_unique<CharT[]>(maxValueNameLen);

                // The value names and types will be stored here
                std::vector<std::pair<std::basic_string<CharT>, DWORD>> valueInfo;

                // Reserve room in the vector to speed up the following insertion loop
                valueInfo.reserve(valueCount);

                // Enumerate all the values
                for(DWORD index = 0; index < valueCount; index++)
                {
                    // Get the name and the type of the current value
                    DWORD valueNameLen = maxValueNameLen;
                    DWORD valueType = 0;
                    retCode = RegEnumValue(m_hKey, index, nameBuffer.get(), &valueNameLen,
                                           nullptr,// reserved
                                           &valueType,
                                           nullptr,// no data
                                           nullptr// no data size
                    );
                    if(retCode != ERROR_SUCCESS)
                    {
                        throw RegException{ retCode, "Cannot enumerate values: RegEnumValue failed." };
                    }

                    // On success, the RegEnumValue API writes the length of the
                    // value name in the valueNameLen output parameter
                    // (not including the terminating NUL).
                    // So we can build a string based on that.
                    valueInfo.emplace_back(std::basic_string<CharT>{ nameBuffer.get(), valueNameLen }, valueType);
                }

                return valueInfo;
            }

            //
            // Misc Registry API Wrappers
            //

            void DeleteValue(const std::basic_string<CharT>& valueName)
            {
                ASSERT(IsValid());

                LONG retCode = RegDeleteValue(m_hKey, valueName.c_str());
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegDeleteValue failed." };
                }
            }

            void DeleteKey(const std::basic_string<CharT>& subKey, REGSAM desiredAccess)
            {
                ASSERT(IsValid());

                LONG retCode = RegDeleteKeyEx(m_hKey, subKey.c_str(), desiredAccess, 0);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegDeleteKeyEx failed." };
                }
            }

            void DeleteTree(const std::basic_string<CharT>& subKey)
            {
                ASSERT(IsValid());

                LONG retCode = RegDeleteTree(m_hKey, subKey.c_str());
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegDeleteTree failed." };
                }
            }

            void CopyTree(const std::basic_string<CharT>& sourceSubKey, const RegKey& destKey)
            {
                ASSERT(IsValid());

                LONG retCode = RegCopyTree(m_hKey, sourceSubKey.c_str(), destKey.Get());
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegCopyTree failed." };
                }
            }
            void FlushKey()
            {
                ASSERT(IsValid());

                LONG retCode = RegFlushKey(m_hKey);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegFlushKey failed." };
                }
            }
            void LoadKey(const std::basic_string<CharT>& subKey, const std::basic_string<CharT>& filename)
            {
                Close();

                LONG retCode = RegLoadKey(m_hKey, subKey.c_str(), filename.c_str());
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegLoadKey failed." };
                }
            }

            void SaveKey(const std::basic_string<CharT>& filename, SECURITY_ATTRIBUTES* securityAttributes) const
            {
                ASSERT(IsValid());

                LONG retCode = RegSaveKey(m_hKey, filename.c_str(), securityAttributes);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegSaveKey failed." };
                }
            }
            void EnableReflectionKey()
            {
                LONG retCode = RegEnableReflectionKey(m_hKey);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegEnableReflectionKey failed." };
                }
            }

            void DisableReflectionKey()
            {
                LONG retCode = RegDisableReflectionKey(m_hKey);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegDisableReflectionKey failed." };
                }
            }
            bool QueryReflectionKey() const
            {
                BOOL isReflectionDisabled = FALSE;
                LONG retCode = RegQueryReflectionKey(m_hKey, &isReflectionDisabled);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegQueryReflectionKey failed." };
                }

                return (isReflectionDisabled ? true : false);
            }
            void ConnectRegistry(const std::basic_string<CharT>& machineName, HKEY hKeyPredefined)
            {
                // Safely close any previously opened key
                Close();

                HKEY hKeyResult = nullptr;
                LONG retCode
                = RegConnectRegistry(machineName.c_str(), hKeyPredefined, &hKeyResult);
                if(retCode != ERROR_SUCCESS)
                {
                    throw RegException{ retCode, "RegConnectRegistry failed." };
                }

                // Take ownership of the result key
                m_hKey = hKeyResult;
            }

            // Return a string representation of Windows registry types
            static std::string RegTypeToString(DWORD regType)
            {
                switch(regType)
                {
                    case REG_SZ:
                        return "REG_SZ";
                    case REG_EXPAND_SZ:
                        return "REG_EXPAND_SZ";
                    case REG_MULTI_SZ:
                        return "REG_MULTI_SZ";
                    case REG_DWORD:
                        return "REG_DWORD";
                    case REG_QWORD:
                        return "REG_QWORD";
                    case REG_BINARY:
                        return "REG_BINARY";

                    default:
                        return "Unknown/unsupported registry type";
                }
            }

            //
            // Relational comparison operators are overloaded as non-members
            // ==, !=, <, <=, >, >=
            //
    };


    //------------------------------------------------------------------------------
    //          Overloads of relational comparison operators for RegKey
    //------------------------------------------------------------------------------

    template<typename CharT>
    inline bool operator==(const RegKey<CharT>& a, const RegKey<CharT>& b) noexcept
    {
        return a.Get() == b.Get();
    }

    template<typename CharT>
    inline bool operator!=(const RegKey<CharT>& a, const RegKey<CharT>& b) noexcept
    {
        return a.Get() != b.Get();
    }

    template<typename CharT>
    inline bool operator<(const RegKey<CharT>& a, const RegKey<CharT>& b) noexcept
    {
        return a.Get() < b.Get();
    }

    template<typename CharT>
    inline bool operator<=(const RegKey<CharT>& a, const RegKey<CharT>& b) noexcept
    {
        return a.Get() <= b.Get();
    }

    template<typename CharT>
    inline bool operator>(const RegKey<CharT>& a, const RegKey<CharT>& b) noexcept
    {
        return a.Get() > b.Get();
    }

    template<typename CharT>
    inline bool operator>=(const RegKey<CharT>& a, const RegKey<CharT>& b) noexcept
    {
        return a.Get() >= b.Get();
    }



    //------------------------------------------------------------------------------
    //                          RegKey Inline Methods
    //------------------------------------------------------------------------------

    template<typename CharT>
    inline void swap(RegKey<CharT>& a, RegKey<CharT>& b) noexcept
    {
        a.SwapWith(b);
    }

    //------------------------------------------------------------------------------
    //                          RegException Inline Methods
    //------------------------------------------------------------------------------

    inline RegException::RegException(const LONG errorCode, const char* const message)
    : std::system_error{ errorCode, std::system_category(), message }
    {
    }

    inline RegException::RegException(const LONG errorCode, const std::string& message)
    : std::system_error{ errorCode, std::system_category(), message }
    {
    }
  
}// namespace winreg

#endif// GIOVANNI_DICANIO_WINREG_HPP_INCLUDED

# WinReg v1.0
## High-level C++ Wrappers Around Low-level Windows Registry C-interface APIs

by Giovanni Dicanio

The Windows Registry C-interface API is  _very low-level_ and _hard_ to use.

I developed some **C++ wrappers** around this low-level Win32 API, to raise the semantic level, using C++ classes like `std::wstring`, `std::vector`, etc. instead of raw C-style buffers and low-level mechanisms. 

For example, the `REG_MULTI_SZ` registry type associated to double-NUL-terminated C-style strings is handled using a much easier higher-level `vector<wstring>`. My C++ code does the _translation_ between high-level C++ STL-based stuff and low-level Win32 C-interface APIs.

Moreover, Win32 error codes are translated to C++ exceptions.

The Win32 registry value types are mapped to C++ higher-level types according the following table:

| Win32 Registry Type  | C++ Type                     |
| -------------------- |:----------------------------:| 
| `REG_DWORD`          | `DWORD`                      |
| `REG_QWORD`          | `ULONGLONG`                  |
| `REG_SZ`             | `std::wstring`               |
| `REG_EXPAND_SZ`      | `std::wstring`               |
| `REG_MULTI_SZ`       | `std::vector<std::wstring>`  |
| `REG_BINARY`         | `std::vector<BYTE>`          |

**NOTE**: I did some tests, and the code works correctly; however, I'd prefer doing _more_ testing. 

Currently, the code compiles cleanly at `/W4` in both 32-bit and 64-bit builds.

Being very busy right now, I preferred releasing this library on GitHub in current status; constructive feedback, bug reports, etc. are welcome.

I developed this code using **Visual Studio 2015 with Update 3**.

The library's code is contained in a **reusable** _header-only_ [`WinReg.hpp`](../master/WinReg/WinReg/WinReg.hpp) file.

`WinRegTest.cpp` contains some demo/test code for the library: check it out for some sample usage.

The library exposes two main classes:

* `RegKey`: a tiny efficient wrapper around raw Win32 `HKEY` handles
* `RegException`: an exception class to signal error conditions

There are many member functions inside the `RegKey` class, that wrap raw Win32 registry C-interface APIs
in a convenient C++ way.

For example, you can simply open a registry key and get a `DWORD` value with C++ code like this:

```c++
RegKey key{ HKEY_CURRENT_USER, L"SOFTWARE\\Connie" };
DWORD dw = key.GetDwordValue(L"MyValue");
```

Or you can enumerate all the values under a key with this simple code:
```c++
auto values = key.EnumValues();
```
 
The library stuff lives under the `winreg` namespace.

See the [**`WinReg.hpp`**](../master/WinReg/WinReg/WinReg.hpp) header for more details and **documentation**.

# WinReg v3.0.0
## High-level C++ Wrapper Around the Low-level Windows Registry C-interface API

by Giovanni Dicanio

The Windows Registry C-interface API is  _very low-level_ and _hard_ to use.

I developed some **C++ wrappers** around this low-level Win32 API, to raise the semantic level, using C++ classes like `std::wstring`, `std::vector`, etc. instead of raw C-style buffers and low-level mechanisms. 

For example, the `REG_MULTI_SZ` registry type associated to double-NUL-terminated C-style strings is handled using a much easier higher-level `vector<wstring>`. My C++ code does the _translation_ between high-level C++ STL-based stuff and the low-level Win32 C-interface API.

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


This code is currently developed using **Visual Studio 2019**. I have no longer tested the code with previous compilers. The code compiles cleanly at `/W4` in both 32-bit and 64-bit builds.

This is a **header-only** library, implemented in the **[`WinReg.hpp`](WinReg/WinReg.hpp)** header file.

`WinRegTest.cpp` contains some demo/test code for the library: check it out for some sample usage.

The library exposes two main classes:

* `RegKey`: a tiny efficient wrapper around raw Win32 `HKEY` handles
* `RegException`: an exception class to signal error conditions

There are many member functions inside the `RegKey` class, that wrap many parts of the native C-interface Windows Registry API, in a convenient C++ way.

For example, you can simply open a registry key and get registry values with C++ code like this:

```c++
RegKey  key{ HKEY_CURRENT_USER, L"SOFTWARE\\SomeKey" };

DWORD   dw = key.GetDwordValue (L"SomeDwordValue");
wstring s  = key.GetStringValue(L"SomeStringValue");
```

Or you can enumerate all the values under a given key with simple C++ code like this:

```c++
auto values = key.EnumValues();

for (const auto & v : values)
{
    //
    // Process current value:
    //
    //   - v.first  (wstring) is the value name
    //   - v.second (DWORD)   is the value type
    //
    ...
}
```

In addition, you can also use the `RegKey::TryGet...Value` methods, that return `std::optional` instead of throwing on errors:

```c++
// RegKey::TryGetDWordValue() returns a std::optional<DWORD>;
// the returned std::optional contains no value on error.

if (auto dw = key.TryGetDwordValue(L"SomeDwordValue"))
{
    // All right: Process the returned value ...
}
else
{
    // The method has failed: The returned std::optional contains no value.   
}
```


You can take a look at the test code in `WinRegTest.cpp` for some sample usage.

The library stuff lives under the `winreg` namespace.

See the [**`WinReg.hpp`**](WinReg/WinReg.hpp) header for more details and **documentation**.

Thanks to everyone who contributed to this project with some additional features and constructive comments and suggestions.


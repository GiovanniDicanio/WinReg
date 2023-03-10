cmake_minimum_required(VERSION 3.12)

set(CMAKE_CXX_STANDARD 17 CACHE STRING "The C++ standard to use")
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>" CACHE STRING "MSVC Runtime Library")

project(WinReg LANGUAGES CXX)

include(GNUInstallDirs)

set(HEADERS WinReg/WinReg.hpp)

add_library(WinReg OBJECT ${HEADERS})

set_target_properties(WinReg PROPERTIES LINKER_LANGUAGE CXX)

target_include_directories(
    WinReg PUBLIC "$<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/WinReg>" "$<INSTALL_INTERFACE:WinReg>"
)

install(
    TARGETS WinReg DESTINATION ${LIB_INSTALL_DIR}
    INCLUDES ${HEADERS} DESTINATION ${INCLUDE_INSTALL_DIR}
)

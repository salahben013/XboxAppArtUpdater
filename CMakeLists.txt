cmake_minimum_required(VERSION 3.20)
project(XboxAppArtUpdater LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

add_executable(XboxAppArtUpdater WIN32
  XboxAppArtUpdater.cpp
  XboxAppArtUpdater.rc
)

target_compile_definitions(XboxAppArtUpdater PRIVATE UNICODE _UNICODE)

target_link_libraries(XboxAppArtUpdater PRIVATE
  user32 gdi32 comctl32 shell32 ole32 oleaut32 advapi32 winhttp gdiplus
)

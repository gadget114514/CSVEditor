cmake_minimum_required(VERSION 3.10)
project(CSVEditor)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Define UNICODE for Win32 API
add_definitions(-DUNICODE -D_UNICODE)

add_executable(CSVEditor WIN32 
    src/main.cpp
    src/MemoryMappedFile.cpp
    src/PieceTable.cpp
    src/CsvDocument.cpp
    src/DirectXResources.cpp
    src/MainWindow.cpp
    src/MainWindow.cpp
    src/EditorState.cpp
    src/Localization.cpp
    src/ConfigManager.cpp
)

# Link standard Windows libraries and DirectWrite (as requested for high quality rendering)
target_link_libraries(CSVEditor PRIVATE user32 gdi32 dwrite shlwapi)

add_executable(CSVEditorTests 
    src/test_main.cpp
    src/MemoryMappedFile.cpp
    src/PieceTable.cpp
    src/CsvDocument.cpp
    src/Localization.cpp
)


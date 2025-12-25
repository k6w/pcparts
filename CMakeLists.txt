cmake_minimum_required(VERSION 3.10)
project(PCPartsInfo)

set(CMAKE_CXX_STANDARD 17)

add_executable(pcparts main.cpp)

target_link_libraries(pcparts ole32 oleaut32 wbemuuid)
# PC Parts Info

A C++ console application that retrieves and displays information about your PC components including CPU, RAM, GPU, Storage, and Motherboard. This information can be used to calculate the value of your PC.

## Features

- CPU Information: Name, Manufacturer, Cores, Logical Processors, Max Clock Speed
- RAM Information: Module capacities, Speed, Manufacturer, Type, Total RAM
- GPU Information: Name, Memory, Video Processor
- Storage Information: Model, Manufacturer, Size, Serial Number
- Motherboard Information: Manufacturer, Product, Serial Number, Version

## Requirements

- Windows operating system
- CMake 3.10 or higher
- C++17 compatible compiler

## Build Instructions

1. Clone or download the repository
2. Create a build directory: `mkdir build && cd build`
3. Run CMake: `cmake ..`
4. Build the project: `cmake --build .`

## Usage

Run the executable `pcparts.exe` from the build directory. The application will display all PC component information and wait for you to press Enter to exit.

## Note

This application uses Windows Management Instrumentation (WMI) to query hardware information. It requires administrative privileges to access some hardware details.
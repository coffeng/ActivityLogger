# Makefile for ActivityLogger using Visual Studio tools
# Use: nmake -f ActivityLogger.mak (to build) or nmake -f ActivityLogger.mak clean (to clean)

# Find VS installation (adjust path as needed)
VSINSTALLDIR = C:\Program Files\Microsoft Visual Studio\2022\Community
WINDOWSSDKDIR = C:\Program Files (x86)\Windows Kits\10

# Compiler and linker
CC = cl.exe
LINK = link.exe

# Include paths
INCLUDES = /I"$(VSINSTALLDIR)\VC\Tools\MSVC\14.44.35211\include" \
           /I"$(WINDOWSSDKDIR)\Include\10.0.22621.0\um" \
           /I"$(WINDOWSSDKDIR)\Include\10.0.22621.0\shared" \
           /I"$(WINDOWSSDKDIR)\Include\10.0.22621.0\winrt" \
           /I"$(WINDOWSSDKDIR)\Include\10.0.22621.0\ucrt"

# Library paths
LIBPATHS = /LIBPATH:"$(VSINSTALLDIR)\VC\Tools\MSVC\14.44.35211\lib\x64" \
           /LIBPATH:"$(WINDOWSSDKDIR)\Lib\10.0.22621.0\um\x64" \
           /LIBPATH:"$(WINDOWSSDKDIR)\Lib\10.0.22621.0\ucrt\x64"

# Compiler flags
CFLAGS = /EHsc /W3 /MD /O2 /DWIN32 /D_WINDOWS /DUNICODE /D_UNICODE $(INCLUDES)

# Linker flags and libraries
LDFLAGS = /SUBSYSTEM:WINDOWS /ENTRY:wWinMainCRTStartup $(LIBPATHS)
LIBS = user32.lib psapi.lib shell32.lib comctl32.lib comdlg32.lib kernel32.lib

# Target executable
TARGET = ActivityLogger.exe

# Source files
SOURCES = ActivityLogger.cpp

# Object files
OBJECTS = ActivityLogger.obj

# Default target
all: $(TARGET)

# Build the executable
$(TARGET): $(OBJECTS)
    $(LINK) $(LDFLAGS) $(OBJECTS) $(LIBS) /OUT:$(TARGET)

# Compile source files
ActivityLogger.obj: ActivityLogger.cpp
    $(CC) $(CFLAGS) /c ActivityLogger.cpp

# Clean target
clean:
    -del *.obj
    -del *.exe
    -del *.pdb
    -del *.ilk

# Debug build
debug:
    $(CC) /EHsc /W3 /MDd /Od /Zi /DWIN32 /D_WINDOWS /DUNICODE /D_UNICODE /D_DEBUG $(INCLUDES) /c ActivityLogger.cpp
    $(LINK) /SUBSYSTEM:WINDOWS /ENTRY:wWinMainCRTStartup /DEBUG $(LIBPATHS) $(OBJECTS) $(LIBS) /OUT:$(TARGET)

# Rebuild target
rebuild: clean all

.PHONY: all clean debug rebuild
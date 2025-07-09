### Desktop Icon Layout Manager
This is a simple Windows application that allows you to save and restore the layout of desktop icons. It uses the Win32 API to interact with the desktop and manage icon positions.
## Why This Tool?
I have my icons spread all over my Windows 11 (Preview) desktop. When I installed the latest Windows update (25H2 3/26/2025 26200.5670) it arranged my icons into some sort of default sort. I tried vibe coding an console app that would record the positions of all the icons and save them to a file. Then I wanted the desktop restore the icons positions to the save ones.  Vibe coding when delving into the Win32 api  and user32.dll is really tricky. It turned into a real rabbit hole. I could not get the restore to work and it seems like it also messed up the right-click sort on the desktop. I miss the fences app that grouped icons that Scott Henselman talked about.

## Features
Command line interface to save and restore desktop icon layouts. Command Line arguments: /s and /r for saving and restoring layouts, respectively.
- **Save Layout**: Capture the current positions of desktop icons and save them to a file.
- **Restore Layout**: Restore the desktop icons to their previously saved positions.
- **Command Line Support**: Use command line arguments to save (`/s`) or restore (`/r`) layouts without a GUI.
- **Cross-Version Compatibility**: Designed to work with Windows 11 and later versions.
- **Simple and Lightweight**: No complex installation required, just run the executable. The save file is stored as a json file.
- 
## Requirements

- Windows 11 (tested on 25H2 and later)
- .NET 9 SDK
- Administrator privileges may be required

## Limitations

- This tool interacts with internal Windows APIs and may break with future Windows updates.
- Special system icons (e.g., Recycle Bin, This PC) may not always be handled.
- The app does not group icons or provide "fences" functionality like some commercial tools.

## Credits

Inspired by the need to preserve desktop icon layouts and the challenges of working with the Win32 API. Special thanks to the Windows developer community and Scott Hanselman for sharing productivity tips.

---

**Note:** Use at your own risk. Always back up your data before running tools that interact with system internals.

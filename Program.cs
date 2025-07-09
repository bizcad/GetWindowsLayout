using System;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;
using System.Collections.Generic;
using System.Text.Json;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;

namespace WindowsDesktopLayout
{
    class Program
    {
        // Win32 constants
        const int PROCESS_VM_OPERATION = 0x0008;
        const int PROCESS_VM_READ = 0x0010;
        const int PROCESS_VM_WRITE = 0x0020;
        const int PROCESS_QUERY_INFORMATION = 0x0400;
        const int MEM_COMMIT = 0x1000;
        const int MEM_RELEASE = 0x8000;
        const int PAGE_READWRITE = 0x04;

        const int LVM_FIRST = 0x1000;
        const int LVM_GETITEMCOUNT = LVM_FIRST + 4;
        const int LVM_GETITEMPOSITION = LVM_FIRST + 16;
        const int LVM_GETITEMTEXTW = LVM_FIRST + 115;
        const int LVM_SETITEMPOSITION = LVM_FIRST + 15; // Add this with the other constants

        const int MAX_PATH = 260;
        private const string V = @"\Downloads\DesktopLayout.json";

        // Data structure for JSON serialization
        class DesktopIcon
        {
            public string Name { get; set; } = string.Empty;
            public int X { get; set; }
            public int Y { get; set; }
            public bool IsShortcut { get; set; }
            public bool IsDirectory { get; set; }
            public string? FilePath { get; set; }
            public string? TargetPath { get; set; }
            public string? FileType { get; set; }
            public string? SpecialType { get; set; }
            public string? IconImageBase64 { get; set; }
            public int Index { get; set; } // <-- Add this
            public int ZOrder { get; set; } // <-- Add this
        }

        // Win32 structures
        [StructLayout(LayoutKind.Sequential)]
        struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        struct LVITEM
        {
            public uint mask;
            public int iItem;
            public int iSubItem;
            public uint state;
            public uint stateMask;
            public IntPtr pszText;
            public int cchTextMax;
            public int iImage;
            public IntPtr lParam;
            public int iIndent;
            public int iGroupId;
            public uint cColumns;
            public IntPtr puColumns;
            public IntPtr piColFmt;
            public int iGroup;
        }

        // Win32 API imports
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern IntPtr FindWindow(string lpClassName, string? lpWindowName);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string? lpszWindow);

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr OpenProcess(int dwDesiredAccess, bool bInheritHandle, uint dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool CloseHandle(IntPtr hObject);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress, uint dwSize, uint flAllocationType, uint flProtect);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool VirtualFreeEx(IntPtr hProcess, IntPtr lpAddress, uint dwSize, uint dwFreeType);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool ReadProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, [Out] byte[] lpBuffer, int dwSize, out int lpNumberOfBytesRead);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, byte[] lpBuffer, int dwSize, out int lpNumberOfBytesWritten);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, ref LVITEM lParam);

        static void Main(string[] args)
        {
            List<string>? items = GetDesktopDirectoryItems();
             string? arg = args.Length > 0 ? args[0] : null;
            if (string.IsNullOrEmpty(arg) || arg.Equals("/s", StringComparison.OrdinalIgnoreCase))
            {
                SaveDesktop();
            }
            else if (arg.Equals("/r", StringComparison.OrdinalIgnoreCase))
            {
                RestoreDesktop();
            }
            else
            {
                Console.WriteLine("Usage: WindowsDesktopLayout [/s] [/r]");
                Console.WriteLine("  /s   Save desktop layout (default)");
                Console.WriteLine("  /r   Restore desktop layout");
            }
        }

        static void SaveDesktop()
        {
            try
            {
                IntPtr desktopListView = GetDesktopListViewHandle();
                if (desktopListView == IntPtr.Zero)
                {
                    Console.WriteLine("Could not find the desktop ListView handle.");
                    return;
                }

                // Get the process ID of explorer.exe
                GetWindowThreadProcessId(desktopListView, out uint explorerPid);

                IntPtr hProcess = OpenProcess(
                    PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE | PROCESS_QUERY_INFORMATION,
                    false,
                    explorerPid);

                if (hProcess == IntPtr.Zero)
                {
                    Console.WriteLine("Could not open explorer.exe process.");
                    return;
                }

                // Get icon count
                int iconCount = (int)SendMessage(desktopListView, LVM_GETITEMCOUNT, 0, IntPtr.Zero);

                // Prepare desktop and public desktop paths
                string userDesktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string commonDesktop = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory);

                
                var desktopFiles = new List<string>();
                if (Directory.Exists(userDesktop))
                {
                    desktopFiles.AddRange(
                        Directory.GetFiles(userDesktop)
                            .Where(f => !f.Contains("desktop.ini", StringComparison.OrdinalIgnoreCase))
                    );
                    desktopFiles.AddRange(Directory.GetDirectories(userDesktop));
                }


                // Collect icon data
                var icons = new List<DesktopIcon>();

                for (int i = 0; i < iconCount; i++)
                {
                    // Allocate memory for POINT in explorer process
                    IntPtr remotePoint = VirtualAllocEx(hProcess, IntPtr.Zero, (uint)Marshal.SizeOf<POINT>(), MEM_COMMIT, PAGE_READWRITE);
                    if (remotePoint == IntPtr.Zero)
                    {
                        Console.WriteLine("Failed to allocate memory in remote process.");
                        continue;
                    }

                    // Get icon position
                    SendMessage(desktopListView, LVM_GETITEMPOSITION, i, remotePoint);

                    // Read POINT from explorer process
                    byte[] pointBuffer = new byte[Marshal.SizeOf<POINT>()];
                    if (!ReadProcessMemory(hProcess, remotePoint, pointBuffer, pointBuffer.Length, out _))
                    {
                        Console.WriteLine($"Failed to read icon position for index {i}.");
                        VirtualFreeEx(hProcess, remotePoint, 0, MEM_RELEASE);
                        continue;
                    }
                    POINT pt = ByteArrayToStructure<POINT>(pointBuffer);

                    // Get icon text
                    string iconText = GetListViewItemText(desktopListView, hProcess, i);

                    // Try to find the directory first by name (case-insensitive)
                    string? filePath = desktopFiles
                        .Where(Directory.Exists)
                        .FirstOrDefault(d =>
                            string.Equals(Path.GetFileName(d), iconText, StringComparison.OrdinalIgnoreCase));

                    // If not found as a directory, try to find a matching file (by name or name without extension)
                    if (filePath == null)
                    {
                        filePath = desktopFiles
                            .Where(File.Exists)
                            .FirstOrDefault(f =>
                                string.Equals(Path.GetFileNameWithoutExtension(f), iconText, StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(Path.GetFileName(f), iconText, StringComparison.OrdinalIgnoreCase));
                    }

                    // Consistently check if the found path is a directory
                    bool isDirectory = !string.IsNullOrEmpty(filePath) && Directory.Exists(filePath);
                    bool isShortcut = false;
                    string? fileType = null;
                    string? targetPath = null;
                    string? specialType = null;
                    string? iconImageBase64 = null;

                    if (!string.IsNullOrEmpty(filePath) && File.Exists(filePath))
                    {
                        fileType = Path.GetExtension(filePath);
                        isShortcut = string.Equals(fileType, ".lnk", StringComparison.OrdinalIgnoreCase);

                        if (isShortcut)
                        {
                            // Try to resolve shortcut target
                            targetPath = GetShortcutTarget(filePath);
                        }

                        // Try to get icon image as base64
                        try
                        {
                            using (Icon? icon = Icon.ExtractAssociatedIcon(filePath))
                            {
                                if (icon != null)
                                {
                                    using (Bitmap bmp = icon.ToBitmap())
                                    using (MemoryStream ms = new MemoryStream())
                                    {
                                        bmp.Save(ms, ImageFormat.Png);
                                        iconImageBase64 = Convert.ToBase64String(ms.ToArray());
                                    }
                                }
                            }
                        }
                        catch { /* ignore icon extraction errors */ }
                    }
                    else if (isDirectory)
                    {
                        // Optionally, you could add directory-specific logic here if needed
                    }
                    else
                    {
                        // Handle special icons (Recycle Bin, This PC, etc.)
                        specialType = GetSpecialDesktopIconType(iconText);
                    }

                    Console.WriteLine($"Icon: {iconText}, Position: ({pt.X}, {pt.Y})");

                    // Add to list
                    icons.Add(new DesktopIcon
                    {
                        Name = iconText,
                        X = pt.X,
                        Y = pt.Y,
                        IsShortcut = isShortcut,
                        IsDirectory = isDirectory,
                        FilePath = filePath,
                        TargetPath = targetPath,
                        FileType = fileType,
                        SpecialType = specialType,
                        IconImageBase64 = iconImageBase64,
                        Index = i,         // Save the ListView index
                        ZOrder = i         // Z-order is the same as index in the ListView
                    });

                    // Free memory
                    VirtualFreeEx(hProcess, remotePoint, 0, MEM_RELEASE);
                }

                // Sort icons by Name before saving
                icons = icons.OrderBy(icon => icon.Index).ToList();

                // Save to JSON file
                string downloads = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + V;
                var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
                File.WriteAllText(downloads, JsonSerializer.Serialize(icons, jsonOptions));

                Console.WriteLine($"Desktop layout saved as JSON to: {downloads}");

                CloseHandle(hProcess);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        static void RestoreDesktop()
        {
            try
            {
                // Read the saved layout from JSON
                string jsonPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + V;
                if (!File.Exists(jsonPath))
                {
                    Console.WriteLine($"No saved layout found at: {jsonPath}");
                    return;
                }

                var jsonOptions = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                List<DesktopIcon>? savedIcons = JsonSerializer.Deserialize<List<DesktopIcon>>(File.ReadAllText(jsonPath), jsonOptions);
                if (savedIcons == null)
                {
                    Console.WriteLine("Failed to deserialize saved layout.");
                    return;
                }

                IntPtr desktopListView = GetDesktopListViewHandle();
                if (desktopListView == IntPtr.Zero)
                {
                    Console.WriteLine("Could not find the desktop ListView handle.");
                    return;
                }

                GetWindowThreadProcessId(desktopListView, out uint explorerPid);
                IntPtr hProcess = OpenProcess(
                    PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE | PROCESS_QUERY_INFORMATION,
                    false,
                    explorerPid);

                if (hProcess == IntPtr.Zero)
                {
                    Console.WriteLine("Could not open explorer.exe process.");
                    return;
                }

                int iconCount = (int)SendMessage(desktopListView, LVM_GETITEMCOUNT, 0, IntPtr.Zero);

                // Prepare desktop and public desktop paths
                string userDesktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string commonDesktop = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory);

                // Get all files and directories on both desktops
                var desktopFiles = new List<string>();
                if (Directory.Exists(userDesktop))
                {
                    desktopFiles.AddRange(Directory.GetFiles(userDesktop));
                    desktopFiles.AddRange(Directory.GetDirectories(userDesktop));
                }
                if (Directory.Exists(commonDesktop))
                {
                    desktopFiles.AddRange(Directory.GetFiles(commonDesktop));
                    desktopFiles.AddRange(Directory.GetDirectories(commonDesktop));
                }

                for (int i = 0; i < iconCount; i++)
                {                    
                    // Get icon text
                    string iconText = GetListViewItemText(desktopListView, hProcess, i);

                    // Try to find the directory first by name (case-insensitive)
                    string? folderPath = desktopFiles
                        .Where(Directory.Exists)
                        .FirstOrDefault(d =>
                            string.Equals(Path.GetDirectoryName(d), iconText, StringComparison.OrdinalIgnoreCase));

                    // Try to find the file first by name (case-insensitive)
                    string? filePath = desktopFiles
                        .Where(File.Exists)
                        .FirstOrDefault(d =>
                            string.Equals(Path.GetFileName(d), iconText, StringComparison.OrdinalIgnoreCase));

                    // If not found as a directory, try to find a matching file (by name or name without extension)
                    if (filePath == null)
                    {
                        filePath = desktopFiles
                            .Where(File.Exists)
                            .FirstOrDefault(f =>
                                string.Equals(Path.GetFileNameWithoutExtension(f), iconText, StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(Path.GetFileName(f), iconText, StringComparison.OrdinalIgnoreCase));
                    }
                    if (filePath is not null && filePath.Contains("CPUID"))
                    {
                        Debug.WriteLine($"Icon: {iconText}, FilePath: {filePath}");
                    }

                    // Only process if this is a file or directory (not a special icon)
                    if (!string.IsNullOrEmpty(filePath))
                    {
                        // Try to find a matching saved icon by Name and FilePath
                        var match = savedIcons.FirstOrDefault(saved =>
                            string.Equals(saved.Name, iconText, StringComparison.OrdinalIgnoreCase) &&
                            string.Equals(saved.FilePath ?? "", filePath, StringComparison.OrdinalIgnoreCase));

                        if (match != null)
                        {
                            // Read current icon position
                            IntPtr remotePoint = VirtualAllocEx(hProcess, IntPtr.Zero, (uint)Marshal.SizeOf<POINT>(), MEM_COMMIT, PAGE_READWRITE);
                            if (remotePoint == IntPtr.Zero)
                            {
                                Console.WriteLine($"Failed to allocate memory for icon '{iconText}'.");
                                continue;
                            }

                            // Get current icon position
                            SendMessage(desktopListView, LVM_GETITEMPOSITION, i, remotePoint);
                            byte[] pointBuffer = new byte[Marshal.SizeOf<POINT>()];
                            if (!ReadProcessMemory(hProcess, remotePoint, pointBuffer, pointBuffer.Length, out _))
                            {
                                Console.WriteLine($"Failed to read icon position for '{iconText}'.");
                                VirtualFreeEx(hProcess, remotePoint, 0, MEM_RELEASE);
                                continue;
                            }
                            POINT currentPt = ByteArrayToStructure<POINT>(pointBuffer);

                            // If position matches, skip moving
                            if (currentPt.X == match.X && currentPt.Y == match.Y)
                            {
                                VirtualFreeEx(hProcess, remotePoint, 0, MEM_RELEASE);
                                continue;
                            }

                            // Write the saved POINT to the remote process
                            POINT pt = new POINT { X = match.X, Y = match.Y };
                            byte[] ptBytes = StructureToByteArray(pt);
                            WriteProcessMemory(hProcess, remotePoint, ptBytes, ptBytes.Length, out _);

                            // Set the icon position
                            SendMessage(desktopListView, LVM_SETITEMPOSITION, i, remotePoint); 

                            // Free memory
                            VirtualFreeEx(hProcess, remotePoint, 0, MEM_RELEASE);

                            Console.WriteLine($"Restored icon '{iconText}' to ({match.X}, {match.Y})");
                        }
                        // else: No match in JSON, so this is a new file/dir and should not be moved
                    }
                    // else: Not a file/dir, skip (special icon)
                }

                CloseHandle(hProcess);
                Console.WriteLine("Desktop icon positions restored.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        static IntPtr GetDesktopListViewHandle()
        {
            // Try to find the desktop ListView in the current Windows 11 layout
            IntPtr progman = FindWindow("Progman", null);
            IntPtr desktopWnd = IntPtr.Zero;

            // Sometimes the desktop is a child of WorkerW, not Progman
            IntPtr shellViewWin = FindWindowEx(progman, IntPtr.Zero, "SHELLDLL_DefView", null);
            if (shellViewWin == IntPtr.Zero)
            {
                // Try to find SHELLDLL_DefView under WorkerW windows
                IntPtr workerw = IntPtr.Zero;
                do
                {
                    workerw = FindWindowEx(IntPtr.Zero, workerw, "WorkerW", null);
                    shellViewWin = FindWindowEx(workerw, IntPtr.Zero, "SHELLDLL_DefView", null);
                } while (workerw != IntPtr.Zero && shellViewWin == IntPtr.Zero);
            }

            if (shellViewWin != IntPtr.Zero)
            {
                desktopWnd = FindWindowEx(shellViewWin, IntPtr.Zero, "SysListView32", "FolderView");
            }

            return desktopWnd;
        }

        static string GetListViewItemText(IntPtr listView, IntPtr hProcess, int index)
        {
            // Allocate memory for text buffer in remote process
            int textBufferSize = MAX_PATH * 2;
            IntPtr remoteBuffer = VirtualAllocEx(hProcess, IntPtr.Zero, (uint)textBufferSize, MEM_COMMIT, PAGE_READWRITE);
            if (remoteBuffer == IntPtr.Zero)
                return string.Empty;

            LVITEM lvItem = new LVITEM
            {
                mask = 0x0001, // LVIF_TEXT
                iItem = index,
                iSubItem = 0,
                pszText = remoteBuffer,
                cchTextMax = MAX_PATH
            };

            // Allocate memory for LVITEM in remote process
            int lvItemSize = Marshal.SizeOf<LVITEM>();
            IntPtr remoteLvItem = VirtualAllocEx(hProcess, IntPtr.Zero, (uint)lvItemSize, MEM_COMMIT, PAGE_READWRITE);
            if (remoteLvItem == IntPtr.Zero)
            {
                VirtualFreeEx(hProcess, remoteBuffer, 0, MEM_RELEASE);
                return string.Empty;
            }

            // Write LVITEM to remote process
            byte[] lvItemBytes = StructureToByteArray(lvItem);
            WriteProcessMemory(hProcess, remoteLvItem, lvItemBytes, lvItemBytes.Length, out _);

            // Send LVM_GETITEMTEXTW
            SendMessage(listView, LVM_GETITEMTEXTW, index, remoteLvItem);

            // Read text from remote process
            byte[] textBuffer = new byte[textBufferSize];
            ReadProcessMemory(hProcess, remoteBuffer, textBuffer, textBuffer.Length, out _);

            // Free memory
            VirtualFreeEx(hProcess, remoteBuffer, 0, MEM_RELEASE);
            VirtualFreeEx(hProcess, remoteLvItem, 0, MEM_RELEASE);

            // Convert to string
            string text = Encoding.Unicode.GetString(textBuffer);
            int nullIndex = text.IndexOf('\0');
            if (nullIndex >= 0)
                text = text.Substring(0, nullIndex);

            return text;
        }

        // Resolve .lnk shortcut target using COM interop
        static string? GetShortcutTarget(string shortcutPath)
        {
            if (string.IsNullOrEmpty(shortcutPath) || !File.Exists(shortcutPath))
                return null;

            try
            {
                Type shellLinkType = Type.GetTypeFromProgID("WScript.Shell");
                if (shellLinkType == null)
                    return null;
                dynamic shell = Activator.CreateInstance(shellLinkType)!;
                dynamic shortcut = shell.CreateShortcut(shortcutPath);
                string? target = shortcut.TargetPath as string;
                Marshal.FinalReleaseComObject(shortcut);
                Marshal.FinalReleaseComObject(shell);
                return string.IsNullOrEmpty(target) ? null : target;
            }
            catch
            {
                return null;
            }
        }

        // Identify special desktop icons by name (basic heuristic)
        static string? GetSpecialDesktopIconType(string name)
        {
            if (string.IsNullOrEmpty(name))
                return null;

            string n = name.ToLowerInvariant();

            // Use switch expression for pattern matching
            return n switch
            {
                var s when s.Contains("recycle bin") => "RecycleBin",
                var s when s.Contains("this pc") || s.Contains("my computer") => "ThisPC",
                var s when s.Contains("network") => "Network",
                var s when s.Contains("control panel") => "ControlPanel",
                var s when s.Contains("user") || s.Contains("home") => "UserFolder",
                _ => null
            };
        }

        static T ByteArrayToStructure<T>(byte[] bytes) where T : struct
        {
            GCHandle handle = GCHandle.Alloc(bytes, GCHandleType.Pinned);
            try
            {
                return Marshal.PtrToStructure<T>(handle.AddrOfPinnedObject());
            }
            finally
            {
                handle.Free();
            }
        }

        static byte[] StructureToByteArray<T>(T obj) where T : struct
        {
            int size = Marshal.SizeOf<T>();
            byte[] arr = new byte[size];
            IntPtr ptr = Marshal.AllocHGlobal(size);
            try
            {
                Marshal.StructureToPtr(obj, ptr, false);
                Marshal.Copy(ptr, arr, 0, size);
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
            }
            return arr;
        }

        static List<string> GetDesktopDirectoryItems()
        {
            var items = new List<string>();
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

            if (Directory.Exists(desktopPath))
            {
                // Add files (excluding "desktop.ini")
                items.AddRange(
                    Directory.GetFiles(desktopPath)
                        .Where(f => !f.Contains("desktop.ini", StringComparison.OrdinalIgnoreCase))
                );
                // Add directories
                items.AddRange(Directory.GetDirectories(desktopPath));
            }

            return items;
        }
    }
}
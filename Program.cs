using System;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;

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

    const int MAX_PATH = 260;

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

    static void Main()
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

                Console.WriteLine($"Icon: {iconText}, Position: ({pt.X}, {pt.Y})");

                // Free memory
                VirtualFreeEx(hProcess, remotePoint, 0, MEM_RELEASE);
            }

            CloseHandle(hProcess);
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
}

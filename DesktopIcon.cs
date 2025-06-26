namespace WindowsDesktopLayout
   {
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
       }
   }
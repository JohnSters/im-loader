using System.Collections.Generic;

namespace IMLoader
{
    public class ExcelFileModel
    {
        public required string FilePath { get; set; }
        public List<string> Sheets { get; set; } = new List<string>();
        public string? SelectedSheet { get; set; }
        public bool IsMaster { get; set; } = false;
        public override string ToString() => System.IO.Path.GetFileName(FilePath) + (IsMaster ? " (Master)" : "");
    }
} 
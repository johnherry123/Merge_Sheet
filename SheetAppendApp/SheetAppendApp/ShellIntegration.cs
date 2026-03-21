using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace SheetAppendApp
{
    public static class ShellIntegration
    {
        private const string Root = @"Software\Classes";

        // Thêm .csv
        private const string KeyXlsx_SystemAssoc = Root + @"\SystemFileAssociations\.xlsx\shell\SheetAppendApp.Merge";
        private const string KeyXls_SystemAssoc = Root + @"\SystemFileAssociations\.xls\shell\SheetAppendApp.Merge";
        private const string KeyXlsm_SystemAssoc = Root + @"\SystemFileAssociations\.xlsm\shell\SheetAppendApp.Merge";
        private const string KeyCsv_SystemAssoc = Root + @"\SystemFileAssociations\.csv\shell\SheetAppendApp.Merge";

        private const string KeyXlsx_ExcelProgId = Root + @"\Excel.Sheet.12\shell\SheetAppendApp.Merge";

        private const string KeyDir = Root + @"\Directory\shell\SheetAppendApp.Merge";
        private const string KeyFolder = Root + @"\Folder\shell\SheetAppendApp.Merge";
        private const string KeyDirBg = Root + @"\Directory\Background\shell\SheetAppendApp.Merge";

        private const string MarkerKey = @"Software\SheetAppendApp";
        private const string MarkerVal = "ContextMenuInstalled";

        public static void EnsureRegistered(string exePath)
        {
            if (string.IsNullOrWhiteSpace(exePath)) throw new ArgumentException("exePath is required.");

            using (var mk = Registry.CurrentUser.CreateSubKey(MarkerKey)!)
            {
                var installed = mk.GetValue(MarkerVal) as string;
                if (string.Equals(installed, exePath, StringComparison.OrdinalIgnoreCase))
                { RefreshExplorerShell(); return; }
            }

            // Đăng ký cho các đuôi file
            CreateVerb(KeyXlsx_SystemAssoc, "Merge (Append Sheets)", exePath, $"\"{exePath}\" --merge \"%1\"");
            CreateVerb(KeyXls_SystemAssoc, "Merge (Append Sheets)", exePath, $"\"{exePath}\" --merge \"%1\"");
            CreateVerb(KeyXlsm_SystemAssoc, "Merge (Append Sheets)", exePath, $"\"{exePath}\" --merge \"%1\"");
            CreateVerb(KeyCsv_SystemAssoc, "Merge (Append Sheets)", exePath, $"\"{exePath}\" --merge \"%1\""); // Dành cho .csv
            CreateVerb(KeyXlsx_ExcelProgId, "Merge (Append Sheets)", exePath, $"\"{exePath}\" --merge \"%1\"");

            CreateVerb(KeyDir, "Merge Excel/CSV in this folder", exePath, $"\"{exePath}\" --merge \"%1\"");
            CreateVerb(KeyFolder, "Merge Excel/CSV in this folder", exePath, $"\"{exePath}\" --merge \"%1\"");
            CreateVerb(KeyDirBg, "Merge Excel/CSV here", exePath, $"\"{exePath}\" --merge \"%V\"");

            using (var mk = Registry.CurrentUser.CreateSubKey(MarkerKey)!) mk.SetValue(MarkerVal, exePath);
            RefreshExplorerShell();
        }

        public static void Unregister()
        {
            DeleteTreeIfExists(KeyXlsx_SystemAssoc); DeleteTreeIfExists(KeyXls_SystemAssoc);
            DeleteTreeIfExists(KeyXlsm_SystemAssoc); DeleteTreeIfExists(KeyCsv_SystemAssoc);
            DeleteTreeIfExists(KeyXlsx_ExcelProgId); DeleteTreeIfExists(KeyDir);
            DeleteTreeIfExists(KeyFolder); DeleteTreeIfExists(KeyDirBg);

            try { using var mk = Registry.CurrentUser.CreateSubKey(MarkerKey)!; mk.DeleteValue(MarkerVal, throwOnMissingValue: false); } catch { }
            RefreshExplorerShell();
        }

        private static void CreateVerb(string keyPath, string verbText, string iconPath, string command)
        {
            using var key = Registry.CurrentUser.CreateSubKey(keyPath)!;
            key.SetValue("MUIVerb", verbText); key.SetValue("Icon", iconPath);
            using var cmdKey = key.CreateSubKey("command")!; cmdKey.SetValue("", command);
        }

        private static void DeleteTreeIfExists(string keyPath) { try { Registry.CurrentUser.DeleteSubKeyTree(keyPath, throwOnMissingSubKey: false); } catch { } }

        private const uint SHCNE_ASSOCCHANGED = 0x08000000; private const uint SHCNF_IDLIST = 0x0000;
        [DllImport("shell32.dll")] private static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
        public static void RefreshExplorerShell() { try { SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, IntPtr.Zero, IntPtr.Zero); } catch { } }
    }
}
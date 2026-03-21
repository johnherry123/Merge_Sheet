using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SheetAppendApp
{
    public partial class MainForm : Form
    {
        private RadioButton _rbSingleFile = null!;
        private RadioButton _rbFolder = null!;

        private TextBox _txtInput = null!;
        private TextBox _txtFolder = null!;
        private TextBox _txtOutput = null!;

        private Button _btnBrowseInput = null!;
        private Button _btnBrowseFolder = null!;
        private Button _btnBrowseOutput = null!;
        private bool _outputUserEdited = false;
        private bool _isSettingOutput = false;
        private string? _lastAutoOutput = null;
        private CheckBox _chkRepair = null!;
        private CheckBox _chkTempCopy = null!;
        private CheckBox _chkCopyStyles = null!;
        private CheckBox _chkIncludeSubfolders = null!;

        private Button _btnRun = null!;
        private ProgressBar _progress = null!;
        private TextBox _log = null!;

        // ĐÃ THÊM .CSV VÀO ĐÂY
        private readonly string[] _validExts = { ".xls", ".xlsx", ".xlsm", ".csv" };

        public MainForm()
        {
            Text = "Merge Excel & CSV - Tối ưu NPOI";
            Width = 980; Height = 700;
            StartPosition = FormStartPosition.CenterScreen;

            BuildUi(); EnableDragDrop(); ApplyMode();
        }

        private void EnableDragDrop()
        {
            AllowDrop = true;
            DragEnter += OnDragEnter; DragDrop += OnDragDrop;
            _txtInput.AllowDrop = true; _txtInput.DragEnter += OnDragEnter; _txtInput.DragDrop += OnDragDrop;
            _txtFolder.AllowDrop = true; _txtFolder.DragEnter += OnDragEnter; _txtFolder.DragDrop += OnDragDrop;
            _txtOutput.AllowDrop = true; _txtOutput.DragEnter += OnDragEnter; _txtOutput.DragDrop += OnDragDrop;
        }

        private void OnDragEnter(object? sender, DragEventArgs e)
        {
            if (e.Data != null && e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
            else e.Effect = DragDropEffects.None;
        }

        private void OnDragDrop(object? sender, DragEventArgs e)
        {
            try
            {
                if (e.Data == null) return;
                var items = (string[]?)e.Data.GetData(DataFormats.FileDrop);
                if (items == null || items.Length == 0) return;

                var first = items[0];
                if (Directory.Exists(first)) { _rbFolder.Checked = true; SetFolderPath(first); ApplyMode(); return; }
                if (File.Exists(first))
                {
                    var ext = Path.GetExtension(first).ToLowerInvariant();
                    if (!_validExts.Contains(ext))
                    {
                        MessageBox.Show(this, "Chỉ hỗ trợ file Excel/CSV (.xls, .xlsx, .xlsm, .csv)", "Invalid file", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    _rbSingleFile.Checked = true; SetInputPath(first); ApplyMode(); return;
                }
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "DragDrop Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private void BuildUi()
        {
            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 10, Padding = new Padding(12) };
            for (int i = 0; i < 9; i++) root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); Controls.Add(root);

            var modePanel = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true };
            _rbSingleFile = new RadioButton { Text = "Single file (gộp các sheet trong file)", AutoSize = true, Checked = true };
            _rbFolder = new RadioButton { Text = "Folder (gộp tất cả file Excel/CSV trong folder)", AutoSize = true };
            _rbSingleFile.CheckedChanged += (_, __) => ApplyMode(); _rbFolder.CheckedChanged += (_, __) => ApplyMode();
            modePanel.Controls.Add(new Label { Text = "Mode:", AutoSize = true, Padding = new Padding(0, 6, 6, 0) });
            modePanel.Controls.Add(_rbSingleFile); modePanel.Controls.Add(_rbFolder); root.Controls.Add(modePanel);

            _txtOutput = new TextBox { Dock = DockStyle.Fill };
            _txtOutput.TextChanged += (_, __) => { if (_isSettingOutput) return; _outputUserEdited = true; };

            var rowIn = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 3, AutoSize = true };
            rowIn.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220)); rowIn.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); rowIn.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
            rowIn.Controls.Add(new Label { Text = "Input file (Excel/CSV):", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            _txtInput = new TextBox { Dock = DockStyle.Fill }; rowIn.Controls.Add(_txtInput, 1, 0);
            _btnBrowseInput = new Button { Text = "Browse...", Dock = DockStyle.Fill }; _btnBrowseInput.Click += (_, __) => BrowseInput(); rowIn.Controls.Add(_btnBrowseInput, 2, 0); root.Controls.Add(rowIn);

            var rowFolder = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 3, AutoSize = true };
            rowFolder.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220)); rowFolder.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); rowFolder.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
            rowFolder.Controls.Add(new Label { Text = "Folder:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            _txtFolder = new TextBox { Dock = DockStyle.Fill }; rowFolder.Controls.Add(_txtFolder, 1, 0);
            _btnBrowseFolder = new Button { Text = "Select...", Dock = DockStyle.Fill }; _btnBrowseFolder.Click += (_, __) => BrowseFolder(); rowFolder.Controls.Add(_btnBrowseFolder, 2, 0); root.Controls.Add(rowFolder);

            var rowOut = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 3, AutoSize = true };
            rowOut.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220)); rowOut.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); rowOut.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
            rowOut.Controls.Add(new Label { Text = "Output (.xlsx):", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            rowOut.Controls.Add(_txtOutput, 1, 0);
            _btnBrowseOutput = new Button { Text = "Save as...", Dock = DockStyle.Fill }; _btnBrowseOutput.Click += (_, __) => BrowseOutput(); rowOut.Controls.Add(_btnBrowseOutput, 2, 0); root.Controls.Add(rowOut);

            var optRow = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true };
            _chkRepair = new CheckBox { Text = "Auto repair broken refs", AutoSize = true, Checked = false, Enabled = false };
            _chkTempCopy = new CheckBox { Text = "Work on temp copy", AutoSize = true, Checked = false, Enabled = false };
            _chkCopyStyles = new CheckBox { Text = "Copy styles", AutoSize = true, Checked = false, Enabled = false };
            _chkIncludeSubfolders = new CheckBox { Text = "Include subfolders", AutoSize = true, Checked = false };
            optRow.Controls.Add(_chkRepair); optRow.Controls.Add(_chkTempCopy); optRow.Controls.Add(_chkCopyStyles); optRow.Controls.Add(_chkIncludeSubfolders); root.Controls.Add(optRow);

            _btnRun = new Button { Text = "RUN (FAST MERGE)", Height = 44, Dock = DockStyle.Top }; _btnRun.Click += async (_, __) => await RunAsync(); root.Controls.Add(_btnRun);
            _progress = new ProgressBar { Dock = DockStyle.Top, Style = ProgressBarStyle.Marquee, Visible = false }; root.Controls.Add(_progress);
            root.Controls.Add(new Label { AutoSize = true, Text = "Tip: Kéo-thả file/folder vào đây. Hỗ trợ .xls, .xlsx, .xlsm, .csv" });

            _log = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, ReadOnly = true }; root.Controls.Add(_log);
        }

        private void ApplyMode() { bool f = _rbFolder.Checked; _txtInput.Enabled = !f; _btnBrowseInput.Enabled = !f; _txtFolder.Enabled = f; _btnBrowseFolder.Enabled = f; _chkIncludeSubfolders.Enabled = f; }
        private void BrowseInput() { using var dlg = new OpenFileDialog { Filter = "Excel & CSV Files (*.xls;*.xlsx;*.xlsm;*.csv)|*.xls;*.xlsx;*.xlsm;*.csv|All files (*.*)|*.*", Title = "Chọn file" }; if (dlg.ShowDialog(this) == DialogResult.OK) SetInputPath(dlg.FileName); }
        private void BrowseFolder() { using var dlg = new FolderBrowserDialog { Description = "Chọn folder" }; if (dlg.ShowDialog(this) == DialogResult.OK) SetFolderPath(dlg.SelectedPath); }
        private void BrowseOutput() { using var dlg = new SaveFileDialog { Filter = "Excel Workbook (*.xlsx)|*.xlsx", Title = "Lưu file", FileName = string.IsNullOrWhiteSpace(_txtOutput.Text) ? "merged.xlsx" : Path.GetFileName(_txtOutput.Text) }; if (dlg.ShowDialog(this) == DialogResult.OK) _txtOutput.Text = dlg.FileName; }
        private void SetInputPath(string filePath) { _txtInput.Text = filePath; if (string.IsNullOrWhiteSpace(_txtOutput.Text)) _txtOutput.Text = Path.Combine(Path.GetDirectoryName(filePath)!, Path.GetFileNameWithoutExtension(filePath) + "_merged.xlsx"); }
        private void SetFolderPath(string folderPath) { _txtFolder.Text = folderPath; if (string.IsNullOrWhiteSpace(_txtOutput.Text)) _txtOutput.Text = Path.Combine(folderPath, "merged_all.xlsx"); }
        private void Log(string s) { if (InvokeRequired) { BeginInvoke(new Action<string>(Log), s); return; } _log.AppendText($"[{DateTime.Now:HH:mm:ss}] {s}{Environment.NewLine}"); }

        private bool IsFileLocked(string filePath) { try { if (!File.Exists(filePath)) return false; using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { stream.Close(); } } catch (IOException) { return true; } return false; }

        private async Task RunAsync()
        {
            var output = _txtOutput.Text.Trim();
            if (string.IsNullOrWhiteSpace(output)) { MessageBox.Show(this, "Bạn chưa chọn Output file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            if (IsFileLocked(output)) { MessageBox.Show(this, "File Output đang mở. Vui lòng đóng lại!", "File in use", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            try
            {
                SetBusy(true); _log.Clear(); Log("Bắt đầu xử lý (Fast Merge Engine)...");
                string[] filesToMerge;
                if (_rbSingleFile.Checked)
                {
                    var input = _txtInput.Text.Trim();
                    if (!File.Exists(input)) { MessageBox.Show(this, "Input file không tồn tại.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    filesToMerge = new string[] { input };
                }
                else
                {
                    var folder = _txtFolder.Text.Trim();
                    if (!Directory.Exists(folder)) { MessageBox.Show(this, "Folder không tồn tại.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                    var searchOption = _chkIncludeSubfolders.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                    filesToMerge = Directory.GetFiles(folder, "*.*", searchOption).Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.OrdinalIgnoreCase)).Where(f => _validExts.Contains(Path.GetExtension(f).ToLowerInvariant())).Where(f => !string.Equals(Path.GetFullPath(f), Path.GetFullPath(output), StringComparison.OrdinalIgnoreCase)).OrderBy(f => f, StringComparer.OrdinalIgnoreCase).ToArray();
                    if (filesToMerge.Length == 0) { MessageBox.Show(this, "Folder không có file Excel/CSV hợp lệ.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                }

                var report = await Task.Run(() => NpoiMerger.MergeFast_ToOneXlsx(filesToMerge, output, trySkipHeaderIfMatchesBase: true, Log));
                Log($"====================================="); Log($"HOÀN THÀNH!"); Log($"File thành công: {report.FilesSucceeded}/{report.InputFiles}"); Log($"Tổng số dòng đã gộp: {report.RowsWritten}");
                MessageBox.Show(this, "Đã gộp xong thành công!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { Log("ERROR: " + ex.Message); MessageBox.Show(this, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { SetBusy(false); }
        }

        private void SetBusy(bool busy) { _btnRun.Enabled = !busy; _btnBrowseInput.Enabled = !busy; _btnBrowseFolder.Enabled = !busy; _btnBrowseOutput.Enabled = !busy; _rbSingleFile.Enabled = !busy; _rbFolder.Enabled = !busy; _chkIncludeSubfolders.Enabled = !busy; _progress.Visible = busy; }
    }
}
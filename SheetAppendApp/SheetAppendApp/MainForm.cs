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


        private CheckBox _chkIncludeSubfolders = null!;

        private Button _btnRun = null!;
        private ProgressBar _progress = null!;
        private TextBox _log = null!;
        private TableLayoutPanel _rowInFile = null!;
        private TableLayoutPanel _rowInFolder = null!;
        private FlowLayoutPanel _optRow = null!;

        // ĐÃ THÊM .CSV VÀO ĐÂY
        private readonly string[] _validExts = { ".xls", ".xlsx", ".xlsm", ".csv" };

        public MainForm()
        {
            Text = "Merge Excel & CSV - Tối ưu NPOI";
            Width = 980; Height = 700;
            StartPosition = FormStartPosition.CenterScreen;

            BuildUi(); EnableDragDrop(); ApplyMode();

            // Load lịch sử log toàn cục (khi chạy ngầm trước đó)
            _log.Text = AppLog.GetAllText() + Environment.NewLine;

            // Đăng ký nhận log theo thời gian thực từ Background/Tray
            AppLog.Message += OnAppLogMessage;
            TrayAppContext.TargetSelected += OnTrayTargetSelected;
            TrayAppContext.BusyStateChanged += OnTrayBusyStateChanged;

            FormClosed += (_, __) =>
            {
                AppLog.Message -= OnAppLogMessage;
                TrayAppContext.TargetSelected -= OnTrayTargetSelected;
                TrayAppContext.BusyStateChanged -= OnTrayBusyStateChanged;
            };

            // Khôi phục trạng thái nếu background đang chạy
            if (TrayAppContext.CurrentTarget != null) OnTrayTargetSelected(TrayAppContext.CurrentTarget);
            if (TrayAppContext.IsMergeRunning) OnTrayBusyStateChanged(true);
        }

        private void OnTrayTargetSelected(string targetPath)
        {
            if (InvokeRequired) { BeginInvoke(new Action<string>(OnTrayTargetSelected), targetPath); return; }
            if (Directory.Exists(targetPath))
            {
                _rbFolder.Checked = true;
                SetFolderPath(targetPath);
            }
            else if (File.Exists(targetPath))
            {
                _rbSingleFile.Checked = true;
                SetInputPath(targetPath);
            }
        }

        private void OnTrayBusyStateChanged(bool isBusy)
        {
            if (InvokeRequired) { BeginInvoke(new Action<bool>(OnTrayBusyStateChanged), isBusy); return; }
            SetBusy(isBusy);
        }

        private void OnAppLogMessage(string msg)
        {
            if (InvokeRequired) { BeginInvoke(new Action<string>(OnAppLogMessage), msg); return; }
            _log.AppendText(msg + Environment.NewLine);
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
            BackColor = Color.FromArgb(32, 34, 37);
            ForeColor = Color.FromArgb(220, 220, 220);
            Font = new Font("Segoe UI", 10F, FontStyle.Regular);

            var topPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                ColumnCount = 1,
                Padding = new Padding(24, 24, 24, 10)
            };

            // 1. TITLE & TIP
            var lblTitle = new Label { Text = "🚀 EXCEL & CSV MERGER", Font = new Font("Segoe UI", 18F, FontStyle.Bold), AutoSize = true, ForeColor = Color.FromArgb(88, 101, 242), Margin = new Padding(0, 0, 0, 10) };
            topPanel.Controls.Add(lblTitle);

            var dragTip = new Label { AutoSize = true, Text = "💡 Tip: You can drag & drop files or folders directly anywhere into this window.", ForeColor = Color.FromArgb(170, 170, 170), Margin = new Padding(0, 0, 0, 20) };
            topPanel.Controls.Add(dragTip);

            // 2. MODE SELECTION
            var modePanel = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(0, 0, 0, 15) };
            _rbSingleFile = new RadioButton { Text = "Merge Sheets in a Single File", AutoSize = true, Checked = true, Cursor = Cursors.Hand, Margin = new Padding(15, 0, 25, 0), ForeColor = Color.White };
            _rbFolder = new RadioButton { Text = "Merge Multiple Files in a Folder", AutoSize = true, Cursor = Cursors.Hand, ForeColor = Color.White };
            _rbSingleFile.CheckedChanged += (_, __) => ApplyMode(); _rbFolder.CheckedChanged += (_, __) => ApplyMode();
            var lblMode = new Label { Text = "Mode:", AutoSize = true, Padding = new Padding(0, 2, 0, 0), Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.White };
            modePanel.Controls.Add(lblMode); modePanel.Controls.Add(_rbSingleFile); modePanel.Controls.Add(_rbFolder);
            topPanel.Controls.Add(modePanel);

            // 3. SOURCE (Dynamic)
            _rowInFile = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, AutoSize = true, Margin = new Padding(0, 0, 0, 10) };
            _rowInFile.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110)); _rowInFile.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); _rowInFile.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            _rowInFile.Controls.Add(new Label { Text = "Input File:", AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = Color.White }, 0, 0);
            _txtInput = CreateDarkTextBox(); _rowInFile.Controls.Add(_txtInput, 1, 0);
            _btnBrowseInput = CreateDarkButton("Browse"); _btnBrowseInput.Click += (_, __) => BrowseInput(); _rowInFile.Controls.Add(_btnBrowseInput, 2, 0);
            topPanel.Controls.Add(_rowInFile);

            _rowInFolder = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, AutoSize = true, Margin = new Padding(0, 0, 0, 0) };
            _rowInFolder.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110)); _rowInFolder.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); _rowInFolder.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            _rowInFolder.Controls.Add(new Label { Text = "Input Folder:", AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = Color.White }, 0, 0);
            _txtFolder = CreateDarkTextBox(); _rowInFolder.Controls.Add(_txtFolder, 1, 0);
            _btnBrowseFolder = CreateDarkButton("Browse"); _btnBrowseFolder.Click += (_, __) => BrowseFolder(); _rowInFolder.Controls.Add(_btnBrowseFolder, 2, 0);
            topPanel.Controls.Add(_rowInFolder);

            _optRow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(110, 0, 0, 10) };
            _chkIncludeSubfolders = new CheckBox { Text = "Include Subfolders", AutoSize = true, Checked = false, Cursor = Cursors.Hand, ForeColor = Color.White };
            _optRow.Controls.Add(_chkIncludeSubfolders);
            topPanel.Controls.Add(_optRow);

            // 4. DESTINATION
            _txtOutput = CreateDarkTextBox();
            var rowOut = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, AutoSize = true, Margin = new Padding(0, 0, 0, 20) };
            rowOut.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110)); rowOut.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); rowOut.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            rowOut.Controls.Add(new Label { Text = "Output File:", AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = Color.White }, 0, 0);
            rowOut.Controls.Add(_txtOutput, 1, 0);
            _btnBrowseOutput = CreateDarkButton("Save As"); _btnBrowseOutput.Click += (_, __) => BrowseOutput(); rowOut.Controls.Add(_btnBrowseOutput, 2, 0);
            topPanel.Controls.Add(rowOut);

            // 5. RUN BUTTON
            _btnRun = new Button { Text = "START MERGE", Height = 50, Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 12F, FontStyle.Bold), BackColor = Color.FromArgb(88, 101, 242), ForeColor = Color.White, Cursor = Cursors.Hand, Margin = new Padding(0, 0, 0, 10) };
            _btnRun.FlatAppearance.BorderSize = 0;
            _btnRun.Click += async (_, __) => await RunAsync();
            topPanel.Controls.Add(_btnRun);

            _progress = new ProgressBar { Dock = DockStyle.Fill, Style = ProgressBarStyle.Marquee, Visible = false, Height = 6, Margin = new Padding(0, 0, 0, 10) };
            topPanel.Controls.Add(_progress);

            // 6. LOG CONSOLE
            _log = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                ReadOnly = true,
                BackColor = Color.FromArgb(20, 22, 24),
                ForeColor = Color.FromArgb(87, 242, 135), // Xanh lá cây sáng
                Font = new Font("Consolas", 10F),
                BorderStyle = BorderStyle.None
            };

            var logContainer = new Panel { Dock = DockStyle.Fill, Padding = new Padding(24, 0, 24, 24) };
            logContainer.Controls.Add(_log);

            Controls.Add(logContainer);
            Controls.Add(topPanel);

            ApplyMode();
        }

        private TextBox CreateDarkTextBox()
        {
            var txt = new TextBox
            {
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.FromArgb(43, 45, 49),
                ForeColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 10.5F)
            };
            return txt;
        }

        private Button CreateDarkButton(string text)
        {
            var btn = new Button
            {
                Text = text,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(71, 82, 196),
                ForeColor = Color.White,
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private void ApplyMode()
        {
            bool isFolder = _rbFolder.Checked;
            _txtInput.Enabled = !isFolder;
            _btnBrowseInput.Enabled = !isFolder;
            _txtFolder.Enabled = isFolder;
            _btnBrowseFolder.Enabled = isFolder;
            _chkIncludeSubfolders.Enabled = isFolder;

            if (_rowInFile != null) _rowInFile.Visible = !isFolder;
            if (_rowInFolder != null) _rowInFolder.Visible = isFolder;
            if (_optRow != null) _optRow.Visible = isFolder;
        }
        private void BrowseInput() { using var dlg = new OpenFileDialog { Filter = "Excel & CSV Files (*.xls;*.xlsx;*.xlsm;*.csv)|*.xls;*.xlsx;*.xlsm;*.csv|All files (*.*)|*.*", Title = "Chọn file" }; if (dlg.ShowDialog(this) == DialogResult.OK) SetInputPath(dlg.FileName); }
        private void BrowseFolder() { using var dlg = new FolderBrowserDialog { Description = "Chọn folder" }; if (dlg.ShowDialog(this) == DialogResult.OK) SetFolderPath(dlg.SelectedPath); }
        private void BrowseOutput() { using var dlg = new SaveFileDialog { Filter = "Excel Workbook (*.xlsx)|*.xlsx", Title = "Lưu file", FileName = string.IsNullOrWhiteSpace(_txtOutput.Text) ? "merged.xlsx" : Path.GetFileName(_txtOutput.Text) }; if (dlg.ShowDialog(this) == DialogResult.OK) _txtOutput.Text = dlg.FileName; }
        private void SetInputPath(string filePath) { _txtInput.Text = filePath; if (string.IsNullOrWhiteSpace(_txtOutput.Text)) _txtOutput.Text = Path.Combine(Path.GetDirectoryName(filePath)!, Path.GetFileNameWithoutExtension(filePath) + "_merged.xlsx"); }
        private void SetFolderPath(string folderPath) { _txtFolder.Text = folderPath; if (string.IsNullOrWhiteSpace(_txtOutput.Text)) _txtOutput.Text = Path.Combine(folderPath, "merged_all.xlsx"); }

        private bool IsFileLocked(string filePath) { try { if (!File.Exists(filePath)) return false; using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { stream.Close(); } } catch (IOException) { return true; } return false; }

        private async Task RunAsync()
        {
            if (TrayAppContext.IsMergeRunning) { MessageBox.Show(this, "Hệ thống đang bận gộp một tác vụ khác. Vui lòng chờ!", "Busy", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            var output = _txtOutput.Text.Trim();
            if (string.IsNullOrWhiteSpace(output)) { MessageBox.Show(this, "Bạn chưa chọn Output file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            if (IsFileLocked(output)) { MessageBox.Show(this, "File Output đang mở. Vui lòng đóng lại!", "File in use", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            if (!TrayAppContext.TrySetBusy()) return;

            try
            {
                SetBusy(true);
                AppLog.Clear(); _log.Clear();
                AppLog.Write("Bắt đầu xử lý (Fast Merge Engine)...");

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

                var report = await Task.Run(() => NpoiMerger.MergeFast_ToOneXlsx(filesToMerge, output, trySkipHeaderIfMatchesBase: true, AppLog.Write));
                AppLog.Write($"=====================================");
                AppLog.Write($"HOÀN THÀNH!");
                AppLog.Write($"File thành công: {report.FilesSucceeded}/{report.InputFiles}");
                AppLog.Write($"Tổng số dòng đã gộp: {report.RowsWritten}");
                MessageBox.Show(this, "Đã gộp xong thành công!", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { AppLog.Write("ERROR: " + ex.Message); MessageBox.Show(this, ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { TrayAppContext.SetIdle(); SetBusy(false); }
        }

        private void SetBusy(bool busy) { _btnRun.Enabled = !busy; _btnBrowseInput.Enabled = !busy; _btnBrowseFolder.Enabled = !busy; _btnBrowseOutput.Enabled = !busy; _rbSingleFile.Enabled = !busy; _rbFolder.Enabled = !busy; _chkIncludeSubfolders.Enabled = !busy; _progress.Visible = busy; }
    }
}
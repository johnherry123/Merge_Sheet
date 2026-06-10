using System;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SheetAppendApp
{
    public sealed class TrayAppContext : ApplicationContext
    {
        private readonly string _pipeName;
        private readonly NotifyIcon _tray;
        private readonly CancellationTokenSource _cts = new();

        // UI invoker để marshal về UI thread
        private readonly Control _uiInvoker = new Control();

        private MainForm? _mainForm;

        public static event Action<string>? TargetSelected;
        public static event Action<bool>? BusyStateChanged;
        public static string? CurrentTarget { get; private set; }
        public static bool IsMergeRunning { get; private set; }

        private static readonly object _globalLock = new();

        public static bool TrySetBusy()
        {
            lock (_globalLock)
            {
                if (IsMergeRunning) return false;
                IsMergeRunning = true;
                BusyStateChanged?.Invoke(true);
                return true;
            }
        }

        public static void SetIdle()
        {
            lock (_globalLock)
            {
                IsMergeRunning = false;
                BusyStateChanged?.Invoke(false);
            }
        }

        // ===== Merge single-flight (không spawn nhiều task) =====
        private readonly object _mergeSync = new();
        private Task? _currentMergeTask;
        private string? _pendingTarget;   // chỉ giữ "latest" request khi đang bận

        public TrayAppContext(string pipeName, string? initialTarget)
        {
            _pipeName = pipeName;

            _uiInvoker.CreateControl();

            var menu = new ContextMenuStrip();
            menu.Items.Add("Open", null, (_, __) => ShowMain());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Merge File...", null, (_, __) => PickAndMergeFile());
            menu.Items.Add("Merge Folder...", null, (_, __) => PickAndMergeFolder());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Exit", null, (_, __) => ExitApp());

            _tray = new NotifyIcon
            {
                Visible = true,
                Text = "SheetAppendApp (running)",
                Icon = SystemIcons.Application,
                ContextMenuStrip = menu
            };
            _tray.DoubleClick += (_, __) => ShowMain();

            // Auto register menu (nếu bạn dùng ShellIntegration.EnsureRegistered)
            try
            {
                ShellIntegration.EnsureRegistered(Application.ExecutablePath);
                AppLog.Write("Context menu ensured.");
            }
            catch (Exception ex)
            {
                AppLog.Write("Auto-register failed: " + ex.Message);
            }

            AppLog.Write("App started (tray mode).");

            // IPC server loop
            _ = Task.Run(() => IpcServerLoopAsync(_cts.Token));

            // Nếu chạy từ context menu lúc chưa có instance
            if (!string.IsNullOrWhiteSpace(initialTarget))
                RequestMerge(initialTarget);
        }

        // ===================== UI marshal =====================
        private void RunOnUi(Action action)
        {
            if (_uiInvoker.IsDisposed) return;

            if (_uiInvoker.InvokeRequired)
                _uiInvoker.BeginInvoke(action);
            else
                action();
        }

        private void Balloon(string title, string text, ToolTipIcon icon)
        {
            RunOnUi(() =>
            {
                try { _tray.ShowBalloonTip(2500, title, text, icon); } catch { }
            });
        }

        private void ShowMain()
        {
            RunOnUi(() =>
            {
                if (_mainForm == null || _mainForm.IsDisposed)
                {
                    _mainForm = new MainForm();
                    _mainForm.FormClosed += (_, __) => { };
                }

                _mainForm.Show();
                _mainForm.WindowState = FormWindowState.Normal;
                _mainForm.BringToFront();
                _mainForm.Activate();
            });
        }

        // ===================== Pickers =====================
        private void PickAndMergeFile()
        {
            RunOnUi(() =>
            {
                using var dlg = new OpenFileDialog
                {
                    Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx",
                    Title = "Chọn file Excel để merge (gộp sheet trong file)"
                };
                if (dlg.ShowDialog() == DialogResult.OK)
                    RequestMerge(dlg.FileName);
            });
        }

        private void PickAndMergeFolder()
        {
            RunOnUi(() =>
            {
                using var dlg = new FolderBrowserDialog
                {
                    Description = "Chọn folder để merge (gộp tất cả file Excel trong folder)"
                };
                if (dlg.ShowDialog() == DialogResult.OK)
                    RequestMerge(dlg.SelectedPath);
            });
        }

        // ===================== IPC =====================
        private async Task IpcServerLoopAsync(CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                try
                {
                    var msg = await IpcServer.WaitForMessageAsync(_pipeName, ct).ConfigureAwait(false);
                    if (string.IsNullOrWhiteSpace(msg)) continue;

                    if (msg.StartsWith("MERGE|", StringComparison.OrdinalIgnoreCase))
                    {
                        var target = msg.Substring("MERGE|".Length);
                        RequestMerge(target);
                    }
                    else if (msg.Equals("SHOW", StringComparison.OrdinalIgnoreCase))
                    {
                        ShowMain();
                    }
                }
                catch (OperationCanceledException) { }
                catch (Exception ex)
                {
                    AppLog.Write("IPC server error: " + ex.Message);
                }
            }
        }

        // ===================== SINGLE-FLIGHT MERGE =====================
        /// <summary>
        /// Không spawn nhiều task. Nếu đang merge thì chỉ giữ "latest" request.
        /// </summary>
        private void RequestMerge(string targetPath)
        {
            targetPath = (targetPath ?? "").Trim().Trim('"');
            if (string.IsNullOrWhiteSpace(targetPath)) return;

            CurrentTarget = targetPath;
            TargetSelected?.Invoke(targetPath);

            lock (_mergeSync)
            {
                if (_currentMergeTask != null && !_currentMergeTask.IsCompleted)
                {
                    // đang chạy -> chỉ giữ request cuối
                    _pendingTarget = targetPath;
                    AppLog.Write($"[BUSY] Merge đang chạy. Queue latest: {targetPath}");
                    Balloon("Busy", "Đang merge... yêu cầu mới sẽ chạy sau khi xong.", ToolTipIcon.Info);
                    return;
                }

                _currentMergeTask = RunMergeAsync(targetPath);
            }
        }

        private async Task RunMergeAsync(string targetPath)
        {
            if (!TrySetBusy()) return;
            try
            {
                AppLog.Write($"[MERGE] Start: {targetPath}");
                Balloon("Merging...", targetPath, ToolTipIcon.Info);

                var result = await MergeRunner.MergeAsync(targetPath).ConfigureAwait(false);

                AppLog.Write($"[MERGE] Done: {result.OutputPath}");
                RunOnUi(() =>
                {
                    try
                    {
                        
                        var r = result.Report;
                        MessageBox.Show(
                            $"Merge xong!\n\nOutput:\n{result.OutputPath}\n\n" +
                            $"SheetsMerged: {r.SheetsMerged}\n" +
                            $"RowsWritten: {r.RowsWritten}\n" +
                            $"DataRows: {r.DataRows}\n" +
                            $"BlankRowsPreserved: {r.BlankRowsPreserved}\n" +
                            $"HeadersSkipped: {r.HeadersSkipped}",
                            "SheetAppendApp",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    catch { }
                });

                Balloon("Merge done", result.OutputPath, ToolTipIcon.Info);
            }
            catch (Exception ex)
            {
                AppLog.Write("[MERGE] Failed: " + ex);

                RunOnUi(() =>
                {
                    try
                    {
                        MessageBox.Show(
                            ex.Message,
                            "Merge failed",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                    }
                    catch { }
                });

                Balloon("Merge failed", ex.Message, ToolTipIcon.Error);
            }
            finally
            {
                SetIdle();

                // ✅ Task kết thúc -> nếu có pending latest thì chạy tiếp đúng 1 cái
                string? next = null;
                lock (_mergeSync)
                {
                    next = _pendingTarget;
                    _pendingTarget = null;
                    _currentMergeTask = null;
                }

                if (!string.IsNullOrWhiteSpace(next))
                {
                    AppLog.Write($"[QUEUE] Run latest queued: {next}");
                    _ = Task.Run(() => RequestMerge(next));
                }
            }
        }

        // ===================== Exit =====================
        private void ExitApp()
        {
            _cts.Cancel();

            RunOnUi(() =>
            {
                try
                {
                    _tray.Visible = false;
                    _tray.Dispose();
                }
                catch { }
            });

            // ✅ thoát hẳn
            Application.Exit();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _cts.Cancel();

                try
                {
                    RunOnUi(() =>
                    {
                        _tray.Visible = false;
                        _tray.Dispose();
                    });
                }
                catch { }

                _cts.Dispose();
                _uiInvoker.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
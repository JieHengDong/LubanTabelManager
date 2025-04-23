using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Text.Json;

namespace LubanManagerTool
{
    public partial class MainWindow : Window
    {
        private const string ConfigFilePath = "config.json";
        private AppConfig appConfig = new AppConfig();
        
        private string selectedPath = "";

        public MainWindow()
        {
            InitializeComponent();

            LoadConfig();

            // 恢复 UI 上的配置内容
            LubanBatPathTextBox.Text = appConfig.LubanBatPath ?? "";
            if (!string.IsNullOrEmpty(appConfig.ExcelFolderPath))
            {
                selectedPath = appConfig.ExcelFolderPath;
                TxtPathDisplay.Text = selectedPath;
                RefreshExcelList();
            }

            Log("请先选择表格目录。");
        }
        
        private void LoadConfig()
        {
            if (File.Exists(ConfigFilePath))
            {
                try
                {
                    string json = File.ReadAllText(ConfigFilePath);
                    appConfig = JsonSerializer.Deserialize<AppConfig>(json) ?? new AppConfig();
                }
                catch
                {
                    appConfig = new AppConfig();
                }
            }
        }

        private void SaveConfig()
        {
            try
            {
                string json = JsonSerializer.Serialize(appConfig, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigFilePath, json);
            }
            catch (Exception ex)
            {
                Log("保存配置失败：" + ex.Message);
            }
        }
        private void BtnChooseLubanBat_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "批处理文件 (*.bat)|*.bat",
                Title = "选择 Luban 导出脚本"
            };

            if (dialog.ShowDialog() == true)
            {
                appConfig.LubanBatPath = dialog.FileName;
                LubanBatPathTextBox.Text = appConfig.LubanBatPath;
                SaveConfig();
            }
        }


        private void BtnChooseFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "请选择包含 Excel 表格的文件夹";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    selectedPath = dialog.SelectedPath;
                    TxtPathDisplay.Text = selectedPath;
                    RefreshExcelList();
                    appConfig.ExcelFolderPath = selectedPath;
                    SaveConfig();
                }
            }
        }

        private void RefreshExcelList()
        {
            ExcelListBox.Items.Clear();
            if (string.IsNullOrEmpty(selectedPath) || !Directory.Exists(selectedPath))
            {
                Log("路径无效，未加载文件。");
                return;
            }

            try
            {
                // 获取当前目录文件
                var files = new List<string>(Directory.GetFiles(selectedPath, "*.xlsx"));
        
                // 获取一级子目录文件
                foreach (var subDir in Directory.GetDirectories(selectedPath))
                {
                    files.AddRange(Directory.GetFiles(subDir, "*.xlsx"));
                }

                foreach (var file in files)
                    ExcelListBox.Items.Add(Path.GetFileName(file));

                Log($"加载完成：共 {files.Count} 个表格文件。");
            }
            catch (Exception ex)
            {
                Log($"加载错误：{ex.Message}");
            }
        }
        
        private void ExcelListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // 确保有选中项
            if (ExcelListBox.SelectedItem is string fileName && !string.IsNullOrEmpty(selectedPath))
            {
                string fullPath = System.IO.Path.Combine(selectedPath, fileName);

                if (File.Exists(fullPath))
                {
                    try
                    {
                        Process.Start(new ProcessStartInfo(fullPath) { UseShellExecute = true });
                        Log($"打开文件：{fileName}");
                    }
                    catch (Exception ex)
                    {
                        Log($"打开失败：{ex.Message}");
                    }
                }
                else
                {
                    Log("文件不存在，可能已被删除或移动。");
                }
            }
        }


        private void BtnSvnUpdate_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(selectedPath)) return;

            RunTortoiseProc("update", selectedPath);
        }

        private void BtnSvnCommit_Click(object sender, RoutedEventArgs e)
        {
            var selectedFiles = ExcelListBox.SelectedItems.Cast<string>().ToList();
            if (selectedFiles.Count == 0)
            {
                Log("请先选择要提交的表格。");
                return;
            }

            // 检查文件是否正在被占用
            foreach (var file in selectedFiles)
            {
                var fullPath = Path.Combine(selectedPath, file);
                if (IsFileLocked(fullPath))
                {
                    Log($"文件正在被使用，请先关闭：{file}");
                    return;
                }
            }

            // 构建路径参数
            var paths = string.Join("*", selectedFiles.Select(f => Path.Combine(selectedPath, f)));
            RunTortoiseProc("commit", paths, "/logmsg:\"提交选中表格\"");
        }
        
        private bool IsFileLocked(string filePath)
        {
            try
            {
                using (FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    return false;
                }
            }
            catch (IOException)
            {
                return true; // 文件被占用
            }
        }


        private void BtnExportSelected_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(appConfig.LubanBatPath) || !File.Exists(appConfig.LubanBatPath))
            {
                Log("Luban 导出脚本路径无效，请重新选择。");
                return;
            }

            Log("开始执行 Luban 导出脚本...");
            RunCmd("cmd.exe", $"/C \"{appConfig.LubanBatPath}\"", selectedPath);
        }

        private void RunTortoiseProc(string command, string path, string args = "")
        {
            var proc = new Process();
            proc.StartInfo.FileName = "TortoiseProc.exe";
            proc.StartInfo.Arguments = $"/command:{command} /path:\"{path}\" {args} /closeonend:0";
            proc.StartInfo.UseShellExecute = false;
            try
            {
                proc.Start();
                Log($"执行 SVN {command}：{path}");
            }
            catch (Exception ex)
            {
                Log($"SVN 操作失败：{ex.Message}");
            }
        }

        private void RunCmd(string fileName, string arguments, string workingDirectory)
        {
            var proc = new Process();
            proc.StartInfo.FileName = fileName;
            proc.StartInfo.Arguments = arguments;
            proc.StartInfo.WorkingDirectory = workingDirectory;
            proc.StartInfo.UseShellExecute = false;
            proc.StartInfo.RedirectStandardOutput = true;
            proc.StartInfo.RedirectStandardError = true;
            proc.StartInfo.CreateNoWindow = true;

            proc.OutputDataReceived += (s, e) => { if (e.Data != null) Log(e.Data); };
            proc.ErrorDataReceived += (s, e) => { if (e.Data != null) Log("[ERROR] " + e.Data); };

            try
            {
                proc.Start();
                proc.BeginOutputReadLine();
                proc.BeginErrorReadLine();
                Log($"执行命令：{arguments}");
            }
            catch (Exception ex)
            {
                Log($"命令执行失败：{ex.Message}");
            }
        }

        private void Log(string msg)
        {
            LogBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {msg}\n");
            LogBox.ScrollToEnd();
        }
    }
}

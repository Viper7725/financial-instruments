using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace ReconcileToolLauncher
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            string exeDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string scriptPath = Path.Combine(exeDirectory, "Start-ReconcileTool.ps1");

            if (!File.Exists(scriptPath))
            {
                MessageBox.Show(
                    "未找到启动脚本:\n" + scriptPath,
                    "游戏财务对账工具",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                return;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"" + scriptPath + "\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                WorkingDirectory = exeDirectory
            };

            try
            {
                Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "启动失败:\n" + ex.Message,
                    "游戏财务对账工具",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }
    }
}

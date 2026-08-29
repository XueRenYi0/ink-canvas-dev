using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Windows.Forms;

namespace InkCanvasSetup
{
    internal static class Program
    {
        // 运行期间的安装元数据（发布时改这里即可）
        public const string AppName = "Ink Canvas";
        public const string AppVersion = "5.1.0";
        public const string DisplayVersion = "5.1.2026.0829";
        public const string Publisher = "Ink Canvas Team";
        public const string UninstallRegistryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" + AppName;
        public static readonly string[] PayloadFolders = new[] { "payload" };

        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length > 0 && StringComparer.OrdinalIgnoreCase.Equals(args[0], "/uninstall"))
            {
                UninstallRun();
                return;
            }
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                MessageBox.Show("安装程序异常：" + ex.Message + Environment.NewLine + ex.StackTrace,
                    AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ————————————————————————————————————————————
        // 核心：把 payload 目录（相对 setup 同目录）拷到目标路径 + 快捷 + 卸载注册表
        // ————————————————————————————————————————————
        public static bool IsAdmin()
        {
            using (var id = WindowsIdentity.GetCurrent())
            {
                return new WindowsPrincipal(id).IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        public static void RequireAdminRelaunch(string extraArg)
        {
            var start = new ProcessStartInfo();
            start.FileName = Application.ExecutablePath;
            start.Arguments = extraArg;
            start.UseShellExecute = true;
            start.Verb = "runas";
            try { Process.Start(start); }
            catch (Exception ex) { MessageBox.Show("需要管理员权限才能继续：" + ex.Message); }
            Application.Exit();
        }

        public static void CopyDirectory(string src, string dst)
        {
            var srcDir = new DirectoryInfo(src);
            if (!srcDir.Exists) return;
            Directory.CreateDirectory(dst);
            foreach (var f in srcDir.GetFiles()) f.CopyTo(Path.Combine(dst, f.Name), true);
            foreach (var d in srcDir.GetDirectories()) CopyDirectory(d.FullName, Path.Combine(dst, d.Name));
        }

        // 单文件安装模式：把嵌入资源的 payload.zip（逻辑名以 payload.zip 结尾）释放到临时目录并解压
        public static string ExtractEmbeddedPayload()
        {
            var asm = System.Reflection.Assembly.GetExecutingAssembly();
            string resName = null;
            foreach (var n in asm.GetManifestResourceNames())
            {
                if (n.EndsWith("payload.zip", StringComparison.OrdinalIgnoreCase)) { resName = n; break; }
            }
            if (resName == null) return null;
            string dir = Path.Combine(Path.GetTempPath(), "InkCanvasSetup_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            string zipPath = dir + ".zip";
            using (var s = asm.GetManifestResourceStream(resName))
            using (var fs = File.Create(zipPath))
            {
                s.CopyTo(fs);
            }
            System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, dir);
            try { File.Delete(zipPath); } catch { }
            return dir;
        }

        public static void CreateShortcut(string shortcutPath, string targetPath, string workDir, string description, string iconPath = null)
        {
            Type t = Type.GetTypeFromProgID("WScript.Shell");
            dynamic shell = Activator.CreateInstance(t);
            dynamic sc = shell.CreateShortcut(shortcutPath);
            sc.TargetPath = targetPath;
            if (!String.IsNullOrEmpty(workDir)) sc.WorkingDirectory = workDir;
            if (!String.IsNullOrEmpty(description)) sc.Description = description;
            if (!String.IsNullOrEmpty(iconPath)) sc.IconLocation = iconPath;
            sc.Save();
            Marshal.ReleaseComObject(sc);
            Marshal.ReleaseComObject(shell);
        }

        // 卸载器入口
        private static void UninstallRun()
        {
            if (!IsAdmin()) { RequireAdminRelaunch("/uninstall"); return; }
            try
            {
                // 读注册表取安装目录
                using (var k = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(UninstallRegistryKey))
                {
                    if (k == null) { MessageBox.Show("未检测到已安装的" + AppName, "卸载"); return; }
                    string installLoc = (string)k.GetValue("InstallLocation");
                    if (String.IsNullOrEmpty(installLoc) || !Directory.Exists(installLoc))
                    {
                        MessageBox.Show("找不到安装目录：" + installLoc, "卸载");
                        return;
                    }
                    string desktopShortcut = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), AppName + ".lnk");
                    string startShortcut = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs", AppName + ".lnk");
                    if (File.Exists(desktopShortcut)) File.Delete(desktopShortcut);
                    if (File.Exists(startShortcut)) File.Delete(startShortcut);
                    try { if (Directory.Exists(installLoc)) Directory.Delete(installLoc, true); }
                    catch (Exception ex) { MessageBox.Show("安装目录删除失败：" + ex.Message + Environment.NewLine + installLoc); }
                    try { Microsoft.Win32.Registry.LocalMachine.DeleteSubKeyTree(UninstallRegistryKey); } catch { }
                    MessageBox.Show(AppName + " 已卸载完成。感谢使用！", "卸载");
                }
            }
            catch (Exception ex) { MessageBox.Show("卸载失败：" + ex.Message); }
        }
    }

    public class MainForm : Form
    {
        private TextBox txtPath;
        private CheckBox chkDesktop;
        private CheckBox chkStartMenu;
        private Button btnBrowse;
        private Button btnInstall;
        private Button btnCancel;
        private Label lblInfo;

        public MainForm()
        {
            Text = Program.AppName + "  v" + Program.AppVersion + "  安装向导";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            ClientSize = new System.Drawing.Size(540, 300);
            Font = new System.Drawing.Font("Microsoft YaHei UI", 9f);

            Controls.Add(new Label { Text = "欢迎安装  " + Program.AppName, Left = 20, Top = 18, Width = 500, Height = 30, Font = new System.Drawing.Font("Microsoft YaHei UI", 14f, System.Drawing.FontStyle.Bold) });
            Controls.Add(new Label { Text = "作者：" + Program.Publisher + "         版本：" + Program.DisplayVersion, Left = 20, Top = 50, Width = 500, ForeColor = System.Drawing.Color.Gray });
            Controls.Add(new Label { Text = "安装位置：", Left = 20, Top = 90 });
            txtPath = new TextBox { Left = 20, Top = 110, Width = 380, Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Ink Canvas") };
            btnBrowse = new Button { Text = "浏览…", Left = 410, Top = 109, Width = 100, Height = 26 };
            btnBrowse.Click += (s, e) =>
            {
                using (var d = new FolderBrowserDialog { Description = "选择安装目录" })
                {
                    if (d.ShowDialog() == DialogResult.OK) txtPath.Text = d.SelectedPath;
                }
            };
            Controls.Add(txtPath);
            Controls.Add(btnBrowse);

            chkDesktop = new CheckBox { Left = 20, Top = 150, Width = 220, Text = "创建桌面快捷方式", Checked = true };
            chkStartMenu = new CheckBox { Left = 250, Top = 150, Width = 260, Text = "创建开始菜单快捷方式（所有用户）", Checked = true };
            Controls.Add(chkDesktop);
            Controls.Add(chkStartMenu);

            lblInfo = new Label { Left = 20, Top = 185, Width = 490, Height = 38, ForeColor = System.Drawing.Color.DimGray,
                Text = "· 需要管理员权限写入 Program Files 和注册表卸载项。\r\n· 卸载：控制面板→程序→" + Program.AppName + "→卸载。" };
            Controls.Add(lblInfo);

            btnCancel = new Button { Text = "取消", Left = 330, Top = 250, Width = 90, Height = 32, DialogResult = DialogResult.Cancel };
            btnInstall = new Button { Text = "安装", Left = 430, Top = 250, Width = 90, Height = 32 };
            btnInstall.Click += BtnInstall_Click;
            AcceptButton = btnInstall; CancelButton = btnCancel;
            Controls.Add(btnCancel);
            Controls.Add(btnInstall);
        }

        private void BtnInstall_Click(object sender, EventArgs e)
        {
            if (!Program.IsAdmin()) { Program.RequireAdminRelaunch(""); return; }
            btnInstall.Enabled = false;
            string targetDir = txtPath.Text;
            try
            {
                // 查找 payload：优先 exe 同目录下的 payload 文件夹（zip 分发模式）；
                // 找不到则释放嵌入资源 payload.zip（单文件 exe 安装包模式）
                string exeDir = Path.GetDirectoryName(Application.ExecutablePath) ?? ".";
                string payloadRoot = Path.Combine(exeDir, "payload");
                bool embedded = false;
                if (!Directory.Exists(payloadRoot) || Directory.GetFileSystemEntries(payloadRoot).Length == 0)
                {
                    payloadRoot = Program.ExtractEmbeddedPayload();
                    embedded = payloadRoot != null;
                }
                if (payloadRoot == null || !Directory.Exists(payloadRoot))
                    throw new Exception("找不到安装数据（payload）：外部 payload 文件夹与内置数据均不存在，安装包可能已损坏。");
                if (embedded) lblInfo.Text = "已释放内置安装数据，正在复制文件…";
                Application.DoEvents();

                Directory.CreateDirectory(targetDir);
                lblInfo.Text = "正在复制文件…";
                Application.DoEvents();
                Program.CopyDirectory(payloadRoot, targetDir);

                string mainExe = Path.Combine(targetDir, "Ink Canvas.exe");
                if (!File.Exists(mainExe)) throw new Exception("未找到主程序 Ink Canvas.exe，复制可能失败。");

                // 开始菜单快捷（All Users）
                if (chkStartMenu.Checked)
                {
                    string startMenu = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs");
                    string shortcut = Path.Combine(startMenu, Program.AppName + ".lnk");
                    Program.CreateShortcut(shortcut, mainExe, targetDir, "启动 " + Program.AppName, mainExe);
                }

                // 桌面快捷（当前用户）
                if (chkDesktop.Checked)
                {
                    // 复制到所有用户公共桌面，供所有人看到；写入失败就写当前用户
                    string commonDesktop = Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory);
                    string sc1 = Path.Combine(commonDesktop, Program.AppName + ".lnk");
                    try { Program.CreateShortcut(sc1, mainExe, targetDir, "启动 " + Program.AppName, mainExe); }
                    catch
                    {
                        string sc2 = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), Program.AppName + ".lnk");
                        Program.CreateShortcut(sc2, mainExe, targetDir, "启动 " + Program.AppName, mainExe);
                    }
                }

                // 写卸载注册表
                try
                {
                    long sizeBytes = 0;
                    foreach (var f in new DirectoryInfo(targetDir).GetFiles("*", SearchOption.AllDirectories)) sizeBytes += f.Length;
                    using (var key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(Program.UninstallRegistryKey))
                    {
                        key.SetValue("DisplayName", Program.AppName + "  " + Program.AppVersion);
                        key.SetValue("DisplayVersion", Program.DisplayVersion);
                        key.SetValue("Publisher", Program.Publisher);
                        key.SetValue("InstallLocation", targetDir);
                        key.SetValue("UninstallString", "\"" + Application.ExecutablePath + "\" /uninstall");
                        key.SetValue("QuietUninstallString", "\"" + Application.ExecutablePath + "\" /uninstall");
                        key.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"));
                        key.SetValue("EstimatedSize", (int)(sizeBytes / 1024));
                        key.SetValue("NoModify", 1, Microsoft.Win32.RegistryValueKind.DWord);
                        key.SetValue("NoRepair", 1, Microsoft.Win32.RegistryValueKind.DWord);
                        try { key.SetValue("DisplayIcon", mainExe + ",0"); } catch { }
                    }
                    // 同时复制卸载主程序到安装目录下，保证控制面板卸载以后用户删除 setup.exe 也能卸载
                    string installSetup = Path.Combine(targetDir, "unins000.exe");
                    File.Copy(Application.ExecutablePath, installSetup, true);
                    using (var key = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(Program.UninstallRegistryKey))
                    {
                        key.SetValue("UninstallString", "\"" + installSetup + "\" /uninstall");
                        key.SetValue("QuietUninstallString", "\"" + installSetup + "\" /uninstall");
                    }
                }
                catch (Exception regEx) { MessageBox.Show("写入卸载信息失败（但软件已安装完成）：" + regEx.Message); }

                lblInfo.Text = "安装完成！";
                if (MessageBox.Show(Program.AppName + " 安装成功！\r\n是否立即启动？",
                    "安装完成", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                {
                    Process.Start(new ProcessStartInfo { FileName = mainExe, WorkingDirectory = targetDir });
                }
                Close();
            }
            catch (Exception ex)
            {
                lblInfo.Text = "安装失败";
                MessageBox.Show("安装失败：" + ex.Message + Environment.NewLine + ex.StackTrace,
                    Program.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnInstall.Enabled = true;
            }
        }
    }
}

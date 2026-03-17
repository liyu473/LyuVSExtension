using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Task = System.Threading.Tasks.Task;

namespace LyuVSExtension
{
    internal sealed class OpenLogsCommand
    {
        public const int CommandId = 0x0100;
        public static readonly Guid CommandSet = new Guid("a1b2c3d4-e5f6-4a5b-8c9d-0e1f2a3b4c5d");

        private readonly AsyncPackage package;

        private OpenLogsCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new OleMenuCommand(this.Execute, menuCommandID);
            menuItem.BeforeQueryStatus += MenuItem_BeforeQueryStatus;
            commandService.AddCommand(menuItem);
        }

        public static OpenLogsCommand Instance { get; private set; }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider => this.package;

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new OpenLogsCommand(package, commandService);
        }

        private void MenuItem_BeforeQueryStatus(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var menuCommand = sender as OleMenuCommand;
            if (menuCommand != null)
            {
                menuCommand.Visible = IsProjectSelected();
            }
        }

        private bool IsProjectSelected()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var dte = Package.GetGlobalService(typeof(DTE)) as DTE;
                if (dte?.SelectedItems == null) return false;

                foreach (SelectedItem item in dte.SelectedItems)
                {
                    if (item.Project != null)
                    {
                        return true;
                    }
                }
            }
            catch { }
            return false;
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var dte = Package.GetGlobalService(typeof(DTE)) as DTE;
                if (dte?.SelectedItems == null) return;

                foreach (SelectedItem item in dte.SelectedItems)
                {
                    if (item.Project != null)
                    {
                        string projectPath = item.Project.FullName;
                        string projectDir = Path.GetDirectoryName(projectPath);

                        // 获取项目目标框架
                        string targetFramework = GetProjectTargetFramework(item.Project);

                        // 查找 bin/debug/netx/logs 文件夹
                        string logsPath = FindLogsFolder(projectDir, targetFramework);

                        if (!string.IsNullOrEmpty(logsPath))
                        {
                            OpenInVSCode(logsPath);
                        }
                        else
                        {
                            VsShellUtilities.ShowMessageBox(
                                this.package,
                                $"未找到 logs 文件夹。\n目标框架: {targetFramework ?? "未知"}\n请确保项目已编译并且存在 bin/debug/{targetFramework}/logs 路径。",
                                "提示",
                                OLEMSGICON.OLEMSGICON_INFO,
                                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                        }
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    $"打开日志文件夹时出错: {ex.Message}",
                    "错误",
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        private string GetProjectTargetFramework(Project project)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                // 方法1：直接从项目文件读取
                string projectPath = project.FullName;
                if (File.Exists(projectPath))
                {
                    string projectContent = File.ReadAllText(projectPath);

                    // 查找 <TargetFramework> 标签
                    var match = System.Text.RegularExpressions.Regex.Match(
                        projectContent,
                        @"<TargetFramework>([^<]+)</TargetFramework>",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                    if (match.Success)
                    {
                        return match.Groups[1].Value.Trim();
                    }
                }

                // 方法2：从 TargetFrameworkMoniker 解析
                var property = project.Properties?.Item("TargetFrameworkMoniker");
                if (property != null)
                {
                    string moniker = property.Value?.ToString();
                    if (!string.IsNullOrEmpty(moniker))
                    {
                        // 格式: ".NETCoreApp,Version=v10.0" 或 ".NETFramework,Version=v4.7.2"
                        var parts = moniker.Split(',');
                        if (parts.Length >= 2)
                        {
                            var versionPart = parts[1].Trim();
                            if (versionPart.StartsWith("Version=v"))
                            {
                                var version = versionPart.Substring("Version=v".Length);

                                // 判断是 .NET Core/5+/6+ 还是 .NET Framework
                                if (parts[0].Contains("NETCoreApp"))
                                {
                                    return "net" + version; // 如 "net10.0"
                                }
                                else if (parts[0].Contains("NETFramework"))
                                {
                                    return "net" + version.Replace(".", ""); // Framework 版本去掉点号
                                }
                            }
                        }
                    }
                }
            }
            catch { }

            return null;
        }

        private string FindLogsFolder(string projectDir, string targetFramework)
        {
            string binDebugPath = Path.Combine(projectDir, "bin", "Debug");

            if (!Directory.Exists(binDebugPath))
                return null;

            // 如果有目标框架，优先查找对应的文件夹
            if (!string.IsNullOrEmpty(targetFramework))
            {
                string specificLogsPath = Path.Combine(binDebugPath, targetFramework, "logs");
                if (Directory.Exists(specificLogsPath))
                {
                    return specificLogsPath;
                }
            }

            // 如果没有找到或没有目标框架信息，查找所有 net* 文件夹（降级方案）
            var netFolders = Directory.GetDirectories(binDebugPath, "net*")
                .OrderByDescending(d => d);

            foreach (var netFolder in netFolders)
            {
                string logsPath = Path.Combine(netFolder, "logs");
                if (Directory.Exists(logsPath))
                {
                    return logsPath;
                }
            }

            return null;
        }

        private string FindVSCodePath()
        {
            // 尝试常见的 VSCode 安装路径
            string[] possiblePaths = new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Microsoft VS Code", "Code.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft VS Code", "Code.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft VS Code", "Code.exe")
            };

            foreach (var path in possiblePaths)
            {
                if (File.Exists(path))
                {
                    return path;
                }
            }

            // 尝试从 PATH 环境变量查找
            try
            {
                var processStartInfo = new ProcessStartInfo
                {
                    FileName = "where",
                    Arguments = "code",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using (var process = System.Diagnostics.Process.Start(processStartInfo))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();

                    if (process.ExitCode == 0 && !string.IsNullOrWhiteSpace(output))
                    {
                        return output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                    }
                }
            }
            catch { }

            return null;
        }

        private void OpenInVSCode(string path)
        {
            try
            {
                string vscodePath = FindVSCodePath();

                if (string.IsNullOrEmpty(vscodePath))
                {
                    VsShellUtilities.ShowMessageBox(
                        this.package,
                        "未找到 VSCode。请确保已安装 Visual Studio Code。\n\n常见安装位置：\n- %LOCALAPPDATA%\\Programs\\Microsoft VS Code\n- %ProgramFiles%\\Microsoft VS Code",
                        "提示",
                        OLEMSGICON.OLEMSGICON_WARNING,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    return;
                }

                var processStartInfo = new ProcessStartInfo
                {
                    FileName = vscodePath,
                    Arguments = $"\"{path}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                System.Diagnostics.Process.Start(processStartInfo);
            }
            catch (Exception ex)
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    $"无法启动 VSCode。\n错误: {ex.Message}",
                    "错误",
                    OLEMSGICON.OLEMSGICON_CRITICAL,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }
    }
}

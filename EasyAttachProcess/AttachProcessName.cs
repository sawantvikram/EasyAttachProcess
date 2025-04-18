using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace EasyAttachProcess
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class AttachProcessName
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("ce59d21a-826a-4117-9b27-4c66f4f26e43");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;
        private const string ProcessNameFile = "processname.txt";

        private AttachProcessName(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static AttachProcessName Instance { get; private set; }

        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider => this.package;

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new AttachProcessName(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string processName = LoadProcessName();
            if (string.IsNullOrEmpty(processName))
            {
                processName = PromptForProcessName();
                if (string.IsNullOrEmpty(processName))
                {
                    ShowMessage("Process name is required.");
                    return;
                }
                SaveProcessName(processName);
            }

            var dte = (DTE2)Package.GetGlobalService(typeof(DTE));
            foreach (Process proc in dte.Debugger.LocalProcesses)
            {
                if (proc.Name.EndsWith(processName))
                {
                    proc.Attach();
                    ShowMessage($"Successfully attached debugger to {processName}.");
                    return;
                }
            }

            ShowMessage($"Process {processName} was not found.");
        }

        private string LoadProcessName()
        {
            string path = GetSettingsFilePath();
            return File.Exists(path) ? File.ReadAllText(path) : string.Empty;
        }

        private void SaveProcessName(string processName)
        {
            string path = GetSettingsFilePath();
            File.WriteAllText(path, processName);
        }

        private string PromptForProcessName()
        {
            string input = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter the process name (e.g., tcserver.exe):",
                "Process Name Input",
                "");
            return input?.Trim();
        }

        private string GetSettingsDirectory()
        {
            // Define the base path using the Roaming folder for user-specific data
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string extensionFolder = Path.Combine(folderPath, "EasyAttach");

            // Ensure the directory exists
            if (!Directory.Exists(extensionFolder))
            {
                Directory.CreateDirectory(extensionFolder);
            }

            return extensionFolder;
        }

        private string GetSettingsFilePath()
        {
            // Combine the directory with the filename
            return Path.Combine(GetSettingsDirectory(), ProcessNameFile);
        }
        private void ShowMessage(string message)
        {
            string title = "MyDebuggerCommand";
            VsShellUtilities.ShowMessageBox(
                this.package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
using Microsoft.VisualBasic;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace EasyAttachProcess
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class EditProcessName
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 4129;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("ce59d21a-826a-4117-9b27-4c66f4f26e43");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;
        private const string ProcessNameFile = "processname.txt";

        private EditProcessName(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static EditProcessName Instance { get; private set; }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new EditProcessName(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread(); // Ensure UI thread access

            string currentProcessName = LoadProcessName();
            string newProcessName = PromptForProcessName(currentProcessName);

            if (!string.IsNullOrEmpty(newProcessName))
            {
                SaveProcessName(newProcessName);
                ShowMessage($"Process name updated to: {newProcessName}");
            }
            else
            {
                ShowMessage("Process name was not changed.");
            }
        }

        private string PromptForProcessName(string currentName)
        {
            var input = Interaction.InputBox(
                "Edit the process name (e.g., tcserver.exe):",
                "Edit Process Name",
                currentName);
            return input.Trim();
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

        private string GetSettingsDirectory()
        {
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string extensionFolder = Path.Combine(folderPath, "EasyAttach");
            if (!Directory.Exists(extensionFolder))
            {
                Directory.CreateDirectory(extensionFolder);
            }
            return extensionFolder;
        }

        private string GetSettingsFilePath()
        {
            return Path.Combine(GetSettingsDirectory(), ProcessNameFile);
        }

        private void ShowMessage(string message)
        {
            MessageBox.Show(message, "Process Name Editor", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
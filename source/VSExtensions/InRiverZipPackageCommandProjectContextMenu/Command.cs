using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace InRiverZipPackageCommandProjectContextMenu
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class Command
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("2e8e062a-e031-496c-9382-a2c35035658b");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="Command"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private Command(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static Command Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in Command's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new Command(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            string message = string.Empty;
            string title = string.Empty;
            OLEMSGICON icon = OLEMSGICON.OLEMSGICON_INFO;
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                var result = CreateInRiverZipPackage();
                message = result.Message;
                if (result.Result == PackageCreationStatus.Warning)
                {
                    icon = OLEMSGICON.OLEMSGICON_WARNING;
                }
            }
            catch(Exception ex)
            {
                message = ex.Message;
                title = "Some error =(";
                icon = OLEMSGICON.OLEMSGICON_CRITICAL;
            }
            finally
            {
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    message,
                    title,
                    icon,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
        }

        private PackageCreationResult CreateInRiverZipPackage()
        {
            var projectPath = GetProjectPath();
            var projectFolderPath = System.IO.Path.GetDirectoryName(projectPath);
            var projectDebugPath = System.IO.Path.Combine(projectFolderPath, "bin", "Debug");
            if (!Directory.Exists(projectDebugPath))
            {
                return new PackageCreationResult() { Message = "Debug directory does not exist! =\\", Result = PackageCreationStatus.Warning };
            }
            var files = GetPackageFiles(projectDebugPath);
            if (files.Count() == 0)
            {
                return new PackageCreationResult() { Message = "There is no files to pack! =\\", Result = PackageCreationStatus.Warning };
            }
            var zipFileName = Path.GetFileNameWithoutExtension(projectPath) + ".zip";
            var zipFilePath = Path.Combine(projectDebugPath, zipFileName);
            CreateZipFile(files, zipFilePath);
            return new PackageCreationResult() { Message = "Hey! Package created =)", Result = PackageCreationStatus.Success };
        }

        private string GetProjectPath()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            IntPtr hierarchyPointer, selectionContainerPointer;
            Object selectedObject = null;
            IVsMultiItemSelect multiItemSelect;
            uint projectItemId;

            IVsMonitorSelection monitorSelection =
                    (IVsMonitorSelection)Package.GetGlobalService(
                    typeof(SVsShellMonitorSelection));

            monitorSelection.GetCurrentSelection(out hierarchyPointer,
                                                 out projectItemId,
                                                 out multiItemSelect,
                                                 out selectionContainerPointer);

            IVsHierarchy selectedHierarchy = Marshal.GetTypedObjectForIUnknown(
                                                 hierarchyPointer,
                                                 typeof(IVsHierarchy)) as IVsHierarchy;

            if (selectedHierarchy != null)
            {
                ErrorHandler.ThrowOnFailure(selectedHierarchy.GetProperty(
                                                  projectItemId,
                                                  (int)__VSHPROPID.VSHPROPID_ExtObject,
                                                  out selectedObject));
            }

            Project selectedProject = selectedObject as Project;

            string projectPath = selectedProject.FullName;
            return projectPath;
        }

        private IEnumerable<string> GetPackageFiles(string projectDebugPath)
        {
            var extensions = new string[] { "*.dll", "*.xml", "*.config" };
            var files = new List<string>();
            foreach (var ext in extensions)
            {
                files.AddRange(System.IO.Directory.GetFiles(projectDebugPath, ext));
            }
            return files;
        }

        private void CreateZipFile(IEnumerable<string> files, string zipFilePath)
        {
            using (FileStream zipToOpen = new FileStream(zipFilePath, FileMode.Create))
            {
                using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Update))
                {
                    foreach(var file in files)
                    {
                        ZipArchiveEntry entry = archive.CreateEntry(Path.GetFileName(file));
                        using (BinaryWriter writer = new BinaryWriter(entry.Open()))
                        {
                            writer.Write(File.ReadAllBytes(file));
                        }
                    }
                }
            }
        }
    }

    public class PackageCreationResult
    {
        public string Message { get; set; }
        public PackageCreationStatus Result { get; set; }
    }

    public enum PackageCreationStatus { Error, Warning, Success }
}

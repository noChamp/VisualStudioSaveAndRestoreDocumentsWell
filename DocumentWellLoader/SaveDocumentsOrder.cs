using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Task = System.Threading.Tasks.Task;
using System.Diagnostics;
using Microsoft.VisualStudio;
using System.Runtime.InteropServices;

namespace DocumentWellLoader
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class SaveDocumentsOrder
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("502e18a8-af11-4a9f-a3b8-74dd15e1159f");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="SaveDocumentsOrder"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private SaveDocumentsOrder(AsyncPackage package, OleMenuCommandService commandService)
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
        public static SaveDocumentsOrder Instance
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
            // Switch to the main thread - the call to AddCommand in SaveDocumentsOrder's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new SaveDocumentsOrder(package, commandService);
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
            ThreadHelper.ThrowIfNotOnUIThread();

            SimpleContainer.Stream = SaveDocumentWindowPositions(SimpleContainer.WinManager);

        }

        private IStream SaveDocumentWindowPositions(IVsUIShellDocumentWindowMgr windowsMgr)
        {
            if (windowsMgr == null)
            {
                Debug.Assert(false, "IVsUIShellDocumentWindowMgr", String.Empty, 0);
                return null;
            }

            IStream stream;
            NativeMethods.CreateStreamOnHGlobal(IntPtr.Zero, true, out stream);
            if (stream == null)
            {
                Debug.Assert(false, "CreateStreamOnHGlobal", String.Empty, 0);
                return null;
            }
            int hr = windowsMgr.SaveDocumentWindowPositions(0, stream);
            if (hr != VSConstants.S_OK)
            {
                Debug.Assert(false, "SaveDocumentWindowPositions", String.Empty, hr);
                return null;
            }

            // Move to the beginning of the stream 
            // In preparation for reading
            LARGE_INTEGER l = new LARGE_INTEGER();
            ULARGE_INTEGER[] ul = new ULARGE_INTEGER[1];
            ul[0] = new ULARGE_INTEGER();
            l.QuadPart = 0;
            //Seek to the beginning of the stream
            stream.Seek(l, 0, ul);
            return stream;
        }
    }

    internal class NativeMethods
    {
        [DllImport("Ole32.dll", EntryPoint = "CreateStreamOnHGlobal")]
        internal static extern void CreateStreamOnHGlobal(IntPtr hGlobal, [MarshalAs(UnmanagedType.Bool)] bool deleteOnRelease, [Out] out Microsoft.VisualStudio.OLE.Interop.IStream stream);
    }
}

using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Text.Outlining;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace CollapseRegionExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class RegionCommands
    {
        /// <summary>
        /// Expand Command ID.
        /// </summary>
        public const int ExpandCommandId = 0x0100;

        /// <summary>
        /// Collapse Command ID.
        /// </summary>
        public const int CollapseCommandId = 0x0101;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("155b6515-ed99-4e79-8950-2bdf525871db");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="RegionCommands"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private RegionCommands(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var explandCommandID = new CommandID(CommandSet, ExpandCommandId);
            var expandMenuItem = new MenuCommand(this.Expand, explandCommandID);
            commandService.AddCommand(expandMenuItem);

            var collapseCommandID = new CommandID(CommandSet, CollapseCommandId);
            var collapseMenuItem = new MenuCommand(this.Collapse, collapseCommandID);
            commandService.AddCommand(collapseMenuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static RegionCommands Instance
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
            // Switch to the main thread - the call to AddCommand in RegionCommands's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new RegionCommands(package, commandService);
        }

        private async void Expand(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            var (manager, regions) = await GetCurrentDocInfoAsync();
            foreach (var region in regions)
            {
                if (region is ICollapsed collapsed)
                {
                    manager.Expand(collapsed);
                }
            }
        }

        private async void Collapse(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            var (manager, regions) = await GetCurrentDocInfoAsync();
            foreach (var region in regions)
            {
                if (region.IsCollapsed == false)
                {
                    manager.TryCollapse(region);
                }
            }
        }

        private async Task<(IOutliningManager manager, IEnumerable<ICollapsible> regions)> GetCurrentDocInfoAsync()
        {
            var emptyResult = (default(IOutliningManager), new ICollapsible[] { });

            var textManager = await ServiceProvider.GetServiceAsync(typeof(SVsTextManager)) as IVsTextManager;
            var componentModel = await this.ServiceProvider.GetServiceAsync(typeof(SComponentModel)) as IComponentModel;
            var outlining = componentModel.GetService<IOutliningManagerService>();

            if (textManager == null || componentModel == null || outlining == null)
            {
                System.Diagnostics.Debug.WriteLine("CollapseRegionExtension failed to retrieve vital service.");
                return emptyResult;
            }

            var editor = componentModel.GetService<IVsEditorAdaptersFactoryService>();

            textManager.GetActiveView(1, null, out IVsTextView textViewCurrent);
            if (textViewCurrent == null)
            {
                return emptyResult;
            }

            var wpfView = editor.GetWpfTextView(textViewCurrent);
            var outliningManager = outlining.GetOutliningManager(wpfView);
            if (outliningManager == null)
            {
                return emptyResult;
            }

            List<ICollapsible> regions = new List<ICollapsible>();
            var snapSHot = new SnapshotSpan(wpfView.TextSnapshot, 0, wpfView.TextSnapshot.Length);
            foreach (var region in outliningManager.GetAllRegions(snapSHot))
            {
                var regionSnapshot = region.Extent.TextBuffer.CurrentSnapshot;
                var text = region.Extent.GetText(regionSnapshot);
                if (text.StartsWith("#region", StringComparison.CurrentCultureIgnoreCase))
                {
                    regions.Add(region);
                }
                if (text.StartsWith("<!--"))
                {
                    var textLow = text.ToLower();
                    if (textLow.Contains("region"))
                    {
                        regions.Add(region);
                    }
                }
            }

            return (outliningManager, regions);
        }
    }
}

//------------------------------------------------------------------------------
// <copyright file="IncreaseFontSize.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace PresentationMode
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class IncreaseFontSize
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid MenuGroup = new Guid("907712e3-9ec2-4bdb-a38a-904088d1bd7f");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="IncreaseFontSize"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private IncreaseFontSize(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                CommandID menuCommandID = new CommandID(MenuGroup, CommandId);
                EventHandler eventHandler = this.HandleIncreaseFontSize;
                MenuCommand menuItem = new MenuCommand(eventHandler, menuCommandID);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static IncreaseFontSize Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
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
        public static void Initialize(Package package)
        {
            Instance = new IncreaseFontSize(package);
        }

        /// <summary>
        /// Shows a message box when the menu item is clicked.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void HandleIncreaseFontSize(object sender, EventArgs e)
        {
            const string CodeLensCategory = "{FC88969A-CBED-4940-8F48-142A503E2381}";
            IncreaseFontSizePackage.AdjustFontSize(ServiceProvider, new Guid(FontsAndColorsCategory.TextEditor), 2);
            IncreaseFontSizePackage.AdjustFontSize(ServiceProvider, new Guid(FontsAndColorsCategory.StatementCompletion), 2);
            IncreaseFontSizePackage.AdjustFontSize(ServiceProvider, new Guid(FontsAndColorsCategory.TextOutputToolWindows), 2);
            IncreaseFontSizePackage.AdjustFontSize(ServiceProvider, new Guid(FontsAndColorsCategory.Tooltip), 2);
            IncreaseFontSizePackage.AdjustFontSize(ServiceProvider, new Guid(CodeLensCategory), 2);
        }
    }
}

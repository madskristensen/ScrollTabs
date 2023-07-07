﻿global using System;
global using Community.VisualStudio.Toolkit;
global using Microsoft.VisualStudio.Shell;
global using Task = System.Threading.Tasks.Task;
using System.ComponentModel.Design;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using Microsoft.VisualStudio;
using Microsoft.Win32;

namespace ScrollTabs
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration(Vsix.Name, Vsix.Description, Vsix.Version)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuids.ScrollTabsString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.ShellInitialized_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class ScrollTabsPackage : ToolkitPackage
    {
        private UIElement _tabWell;

        // OtherContextMenus.EasyMDIToolWindow.ShowTabsInMultipleRows
        private static readonly CommandID _command = new(new Guid("{43F755C7-7916-454D-81A9-90D4914019DD}"), 0x14);

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            Application.Current.MainWindow.PreviewMouseWheel += OnMouseWheel;
        }

        private void OnMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Alt + MouseWheel
            if ((Keyboard.IsKeyDown(Key.LeftAlt) || Keyboard.IsKeyDown(Key.RightAlt)) &&
                !Keyboard.IsKeyDown(Key.LeftShift) && !Keyboard.IsKeyDown(Key.RightShift) &&
                !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl))
            {
                ActivateNextOrPreviousTab(e);
            }

            // MouseWheel over tab well with no modifier keys down
            else if (!Keyboard.IsKeyDown(Key.LeftAlt) && !Keyboard.IsKeyDown(Key.RightAlt) &&
                     !Keyboard.IsKeyDown(Key.LeftShift) && !Keyboard.IsKeyDown(Key.RightShift) &&
                     !Keyboard.IsKeyDown(Key.LeftCtrl) && !Keyboard.IsKeyDown(Key.RightCtrl))
            {
                ToggleMultiRowSetting(e);
            }
        }

        private void ToggleMultiRowSetting(MouseWheelEventArgs e)
        {
            if (_tabWell == null || !_tabWell.IsVisible)
            {
                _tabWell = Application.Current.MainWindow.FindChild<UIElement>("InsertTabPreviewDockTarget");
            }

            if (_tabWell?.IsMouseOver == true)
            {
                bool isMultiRowsEnabled = IsMultiRowsEnabled();

                // MouseWheel up: disable multi rows
                if (e.Delta > 0 && isMultiRowsEnabled)
                {
                    _command.ExecuteAsync().FireAndForget();
                }

                // MouseWheel down: enable multi rows
                else if (e.Delta < 0 && !isMultiRowsEnabled)
                {
                    _command.ExecuteAsync().FireAndForget();
                }

                e.Handled = true;
            }
        }

        private static void ActivateNextOrPreviousTab(MouseWheelEventArgs e)
        {
            string commandName = e.Delta > 0 ? "Window.PreviousTab" : "Window.NextTab";
            VS.Commands.ExecuteAsync(commandName).FireAndForget();
            e.Handled = true;
        }

        private bool IsMultiRowsEnabled()
        {
            using (RegistryKey key = UserRegistryRoot.OpenSubKey("ApplicationPrivateSettings\\WindowManagement\\Options"))
            {
                return ((string)key.GetValue("IsMultiRowTabsEnabled", "0*System.Boolean*False")).EndsWith("True", StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
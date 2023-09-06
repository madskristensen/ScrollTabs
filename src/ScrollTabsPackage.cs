global using System;
global using Community.VisualStudio.Toolkit;
global using Microsoft.VisualStudio.Shell;
global using Task = System.Threading.Tasks.Task;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Reflection;
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
        // Cached reflection MemberInfo
        private static readonly FieldInfo _getFrameField = typeof(WindowFrame).GetRuntimeFields().FirstOrDefault(f => f.Name == "_frame");
        private static readonly Dictionary<Type, PropertyInfo> _getContentProperties = new();
        private static readonly Dictionary<Type, PropertyInfo> _getContainingWindowProperties = new();
        private static readonly Dictionary<Type, PropertyInfo> _getIsActiveProperties = new();
        private static bool _activeFrameChangeDisabled;

        private DateTime _showMultiLineTabsDate;
        private RatingPrompt _rating;

        // OtherContextMenus.EasyMDIToolWindow.ShowTabsInMultipleRows
        private static readonly CommandID _command = new(new Guid("{43F755C7-7916-454D-81A9-90D4914019DD}"), 0x14);

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            _rating = new("MadsKristensen.ScrollTabs", Vsix.Name, await General.GetLiveInstanceAsync(), 2);

            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            // Register mouse wheel events only on Tab well implementation type.
            // At this time that type is named "Microsoft.VisualStudio.PlatformUI.Shell.Controls.DockTarget" and resides in "Microsoft.VisualStudio.Shell.ViewManager.dll".
            // Later filtered by name.
            Type dockTargetType = Assembly.Load("Microsoft.VisualStudio.Shell.ViewManager")
                .GetType("Microsoft.VisualStudio.PlatformUI.Shell.Controls.DockTarget");
            EventManager.RegisterClassHandler(dockTargetType, Mouse.PreviewMouseWheelEvent, new MouseWheelEventHandler(OnMouseWheelOverTabWell), false);

            // Capture all mouse wheel events on windows.
            // Later filtered to windows with documents.
            EventManager.RegisterClassHandler(typeof(Window), Mouse.PreviewMouseWheelEvent, new MouseWheelEventHandler(OnMouseWheelOverWindowAsync), false);

            // Capture active frame changed event.
            // QuickFix: Causes some windows to not work, https://github.com/madskristensen/ScrollTabs/issues/6
            //VS.Events.WindowEvents.ActiveFrameChanged += OnActiveFrameChanged;
        }

        /// <summary>
        /// Handler to handle the mouse-wheel event over the tab well.
        /// </summary>
        private void OnMouseWheelOverTabWell(object sender, MouseWheelEventArgs e)
        {
            // Process only events from tab well element.
            // At this time the name is InsertTabPreviewDockTarget.
            if (sender is not FrameworkElement { Name: "InsertTabPreviewDockTarget" })
            {
                return;
            }

            // MouseWheel over Tab Well with no modifier keys down
            if (!IsAnyAlt() && !IsAnyShift() && !IsAnyCtrl())
            {
                // Enable active frame changed event if is disabled.
                _activeFrameChangeDisabled = false;

                ToggleMultiRowSetting(e);
            }
        }

        /// <summary>
        /// Method to toggle multi-row setting for tab control.
        /// </summary>
        private void ToggleMultiRowSetting(MouseWheelEventArgs e)
        {
            bool isMultiRowsEnabled = IsMultiRowsEnabled();
            bool disableMultiRows = e.Delta > 0 && isMultiRowsEnabled; // MouseWheel up: disable multi rows
            bool enableMultiRows = e.Delta < 0 && !isMultiRowsEnabled; // MouseWheel down: enable multi rows

            if (disableMultiRows || enableMultiRows)
            {
                _command.ExecuteAsync().FireAndForget();
                e.Handled = true;
            }

            if (enableMultiRows)
            {
                _showMultiLineTabsDate = DateTime.Now;
            }
        }

        /// <summary>
        /// Helper method to check if multi-row tabs are enabled.
        /// </summary>
        /// <returns>Returns true if multi-row tabs are enabled; otherwise, false.</returns>
        private bool IsMultiRowsEnabled()
        {
            using (RegistryKey key = UserRegistryRoot.OpenSubKey("ApplicationPrivateSettings\\WindowManagement\\Options"))
            {
                return ((string)key.GetValue("IsMultiRowTabsEnabled", "0*System.Boolean*False")).EndsWith("True", StringComparison.OrdinalIgnoreCase);
            }
        }

        /// <summary>
        /// Handler to handle the mouse-wheel event over the whole window.
        /// </summary>
        private async void OnMouseWheelOverWindowAsync(object sender, MouseWheelEventArgs e)
        {
            // Sender should be always of type Window, but to be sure check and cast.
            if (sender is not Window window)
            {
                return;
            }

            // Process only windows of type MainWindow or FloatingWindow, only these can have document holders with tab well.
            string name = sender.GetType().Name;
            if (name is not ("MainWindow" or "FloatingWindow"))
            {
                return;
            }

            // Alt + MouseWheel
            if (IsAnyAlt() && !IsAnyShift() && !IsAnyCtrl())
            {
                // Activate the document tab holder under mouse (or at least the window under mouse).
                await ActivateTabUnderMouseAsync(window);

                // And move to next or previous tab.
                ActivateNextOrPreviousTab(e);
            }
        }

        /// <summary>
        /// Activate the document tab holder under mouse.
        /// Or at least tries to activate the window under mouse.
        /// </summary>
        private async Task ActivateTabUnderMouseAsync(Window window)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync();

            IEnumerable<WindowFrame> allTabs = await VS.Windows.GetAllDocumentWindowsAsync();
            
            // Get tabs that are visible and inside window under mouse.
            List<TabFrameData> visibleWindowTabs = allTabs
                .Select(GetVsFrameData)
                .Where(tuple => tuple.ContainingWindow == window && tuple.TabContent?.IsVisible == true)
                .ToList();
            
            if (visibleWindowTabs.Count == 1)
            {
                if (window.IsActive) return;

                // Activate the single visible tab in window under mouse.
                TabFrameData singleTab = visibleWindowTabs[0];
                await singleTab.Frame.ShowAsync();
                return;
            }

            if (visibleWindowTabs.Count > 1)
            {
                // Try to find the tab under mouse if there are more then one.
                // TODO Mouse.DirectlyOver returns VsMenu if Alt focused the top left menu and not element directly under mouse.
                // TODO maybe we can use mouse position, frame rectangles and intersection. (tabs shouldn't overlap inside the same window)
                TabFrameData tabUnderMouse = visibleWindowTabs.FirstOrDefault(t => Mouse.DirectlyOver.HasParent(t.TabContent));
                if (tabUnderMouse != null)
                {
                    // And activate it if it is not already active.
                    if (tabUnderMouse.IsActive) return;

                    await tabUnderMouse.Frame.ShowAsync();
                    return;
                }

                // Else if we cant find the tab under mouse (Mouse is over some other element) activate any tab if not already active any.
                if (visibleWindowTabs.Any(t => t.IsActive)) return;

                await visibleWindowTabs.First().Frame.ShowAsync();
            }
        }

        /// <summary>
        /// Use reflection to get some extended data about tab.
        /// </summary>
        private static TabFrameData GetVsFrameData(WindowFrame frame)
        {
            try
            {
                object innerFrame = _getFrameField.GetValue(frame);
                GetFrameProperties(innerFrame.GetType(), out PropertyInfo getContent, out PropertyInfo getContainingWindow, out PropertyInfo getIsActive);
                FrameworkElement content = (FrameworkElement)getContent.GetValue(innerFrame);
                Window containingWindow = (Window)getContainingWindow.GetValue(innerFrame);
                bool isActive = (bool)getIsActive.GetValue(innerFrame);
                return new(frame,  containingWindow, content, isActive);
            }
            catch
            {
                return default;
            }
        }

        /// <summary>
        /// Return cached getters for frame properties of specified type.
        /// This method should be called only in main thread context to prevent race conditions.
        /// </summary>
        private static void GetFrameProperties(Type type, out PropertyInfo getContent, out PropertyInfo getContainingWindow, out PropertyInfo getIsActive)
        {
            if (!_getContentProperties.TryGetValue(type, out getContent))
            {
                getContent = type.GetRuntimeProperties().FirstOrDefault(p => p.Name == "Content");
                _getContentProperties[type] = getContent;
            }

            if (!_getContainingWindowProperties.TryGetValue(type, out getContainingWindow))
            {
                getContainingWindow = type.GetRuntimeProperties().FirstOrDefault(p => p.Name == "ContainingWindow");
                _getContainingWindowProperties[type] = getContainingWindow;
            }

            if (!_getIsActiveProperties.TryGetValue(type, out getIsActive))
            {
                getIsActive = type.GetRuntimeProperties().FirstOrDefault(p => p.Name == "IsActive");
                _getIsActiveProperties[type] = getIsActive;
            }
        }

        /// <summary>
        /// Helper method to activate the next or previous tab based on the mouse wheel delta value.
        /// </summary>
        private static void ActivateNextOrPreviousTab(MouseWheelEventArgs e)
        {
            string commandName = e.Delta > 0 ? "Window.PreviousTab" : "Window.NextTab";

            // Disable active frame changed event, else it will automatically collapse tab well.
            _activeFrameChangeDisabled = true;
            VS.Commands.ExecuteAsync(commandName)
                .ContinueWith(_ => _activeFrameChangeDisabled = false) // And after command restore active frame changed event.
                .FireAndForget();
            e.Handled = true;
        }

        /// <summary>
        /// Handler to handle the active frame changed event that gets triggered when a frame gets loaded or unloaded.
        /// </summary>
        private void OnActiveFrameChanged(ActiveFrameChangeEventArgs obj)
        {
            // The event is temporarily disabled by Move to next or previous tab event processing.
            if (_activeFrameChangeDisabled) return;

            // If less than 5 seconds have passed since the user enabled multi rows, then disable multi rows
            if (_showMultiLineTabsDate.AddSeconds(5) > DateTime.Now && IsMultiRowsEnabled())
            {
                _command.ExecuteAsync().FireAndForget();
                _rating.RegisterSuccessfulUsage();
            }
        }

        /// <summary>
        /// Is pressed left or right Alt key.
        /// </summary>
        private static bool IsAnyAlt() => Keyboard.IsKeyDown(Key.LeftAlt) || Keyboard.IsKeyDown(Key.RightAlt);

        /// <summary>
        /// Is pressed left or right Shift key.
        /// </summary>
        private static bool IsAnyShift() => Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift);

        /// <summary>
        /// Is pressed left or right Ctrl key.
        /// </summary>
        private static bool IsAnyCtrl() => Keyboard.IsKeyDown(Key.LeftCtrl) || Keyboard.IsKeyDown(Key.RightCtrl);

        /// <summary>
        /// Extended data about tab.
        /// </summary>
        private class TabFrameData
        {
            public TabFrameData(WindowFrame frame, Window containingWindow, FrameworkElement tabContent, bool isActive)
            {
                Frame = frame;
                ContainingWindow = containingWindow;
                TabContent = tabContent;
                IsActive = isActive;
            }

            /// <summary>
            /// The frame itself to call commands from VS toolkit.
            /// </summary>
            public WindowFrame Frame { get; }

            /// <summary>
            /// Window in which the tab resides.
            /// </summary>
            public Window ContainingWindow { get; }

            /// <summary>
            /// Element of tab content.
            /// Can be <see langword="null"/> if tab is not visible and never was.
            /// </summary>
            public FrameworkElement TabContent { get; }

            /// <summary>
            /// The tab is active for VS.
            /// </summary>
            public bool IsActive { get; }
        }
    }
}
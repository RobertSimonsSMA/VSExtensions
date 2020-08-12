using CodeValue.VSCopyLocal.OptionsPages;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using VSUtils;
using VSUtils.Project;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Task = System.Threading.Tasks.Task;

namespace CodeValue.VSCopyLocal
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.GuidVsPackageTemplatePkgString)]
    [ProvideOptionPage(typeof(OptionsPage), "VSCopyLocal", "Settings", 0, 0, true)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasSingleProject_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionHasMultipleProjects_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class VSCopyLocalPackage : AsyncPackage, IVsSolutionEvents
    {
        private static DTE _dte;
        private static OptionsPage _options;

        private IVsSolution _solution;
        private uint _hSolutionEvents = uint.MaxValue;
        private IList<HierarchyHandler> _hierarchyHandlers = new List<HierarchyHandler>();


        private const string CActionTextFormat = "Turn {0} Copy Local";
        private const string CProjectActionTextFormat = CActionTextFormat + " for this Project";
        private const string CSolutionActionTextFormat = CActionTextFormat + " for this Solution";

        public string Name
        {
            get { return "VSCopyLocal"; }
        }

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress) {
            await AdviseSolutionEvents();

            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            // Add our command handlers for menu (commands must exist in the .vsct file)
            var mcs = await GetServiceAsync(typeof (IMenuCommandService)) as OleMenuCommandService;
            if (null == mcs) return;

            // Solution level command
            var menuCommandId = new CommandID(GuidList.GuidVSCopyLocalCmdSolutionSet, (int)PkgCmdIDList.CmdidSolutionCommand);
            var menuItem = new OleMenuCommand(SolutionMenuItemCallback, menuCommandId);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
            mcs.AddCommand(menuItem);

            // Project level command
            menuCommandId = new CommandID(GuidList.GuidVSCopyLocalCmdSet, (int) PkgCmdIDList.CmdidProjectCommand);
            menuItem = new OleMenuCommand(ProjectMenuItemCallback, menuCommandId);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
            mcs.AddCommand(menuItem);

            // References level command
            menuCommandId = new CommandID(GuidList.GuidVSCopyLocalCmdReferencesSet, (int) PkgCmdIDList.CmdidReferencesCommand);
            menuItem = new OleMenuCommand(ReferencesMenuItemCallback, menuCommandId);
            menuItem.BeforeQueryStatus += OnBeforeQueryStatus;
            mcs.AddCommand(menuItem);
            
            _dte = (DTE)GetGlobalService(typeof(DTE));
            _options = GetDialogPage(typeof(OptionsPage))as OptionsPage;
        }

        protected override void Dispose(bool disposing) {
            UnadviseSolutionEvents();
            foreach (var hierarchyHandler in _hierarchyHandlers) {
                hierarchyHandler.Unadvise();
            }
            _hierarchyHandlers.Clear();

            base.Dispose(disposing);
        }

        private async Task AdviseSolutionEvents() {
            _solution = (IVsSolution) await GetServiceAsync(typeof(SVsSolution));
            var projectsInSolution = GetProjectsInSolution(_solution);
            foreach (var project in projectsInSolution) {
                AddHierarchyHandler(project);
            }

            _solution.AdviseSolutionEvents(this, out _hSolutionEvents);
        }


        private static IEnumerable<IVsHierarchy> GetProjectsInSolution(IVsSolution solution) {
            __VSENUMPROJFLAGS flags = __VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION;
            Guid guid = Guid.Empty;
            IEnumHierarchies enumHierarchies;
            solution.GetProjectEnum((uint)flags, ref guid, out enumHierarchies);
            if (enumHierarchies == null) {
                yield break;
            }

            IVsHierarchy[] hierarchies = new IVsHierarchy[1];

            uint fetched;
            while (enumHierarchies.Next(1, hierarchies, out fetched) == VSConstants.S_OK && fetched == 1) {
                if (hierarchies.Length > 0 && hierarchies[0] != null) {
                    yield return hierarchies[0];
                }
            }
        }

        private void UnadviseSolutionEvents() {
            if (_solution != null) {
                if (_hSolutionEvents != uint.MaxValue) {
                    _solution.UnadviseSolutionEvents(_hSolutionEvents);

                    _hSolutionEvents = uint.MaxValue;
                }

                _solution = null;
            }
        }

        private static void OnBeforeQueryStatus(object sender, EventArgs e)
        {
            var command = sender as OleMenuCommand;
            UpdateLabelText(command);
        }

        private static void ProjectMenuItemCallback(object sender, EventArgs e)
        {
            SetCopyLocalForProject();
        }

        private static void ReferencesMenuItemCallback(object sender, EventArgs e)
        {
            SetCopyLocalForProject();
        }

        private static void SolutionMenuItemCallback(object sender, EventArgs e)
        {
            SetCopyLocalForSolution();
        }

        /// <summary>
        /// Set Copy Local for the active project.
        /// </summary>
        private static void SetCopyLocalForProject()
        {
            var projects = (Array) _dte.ActiveSolutionProjects;
            var activeProject = (Project) projects.GetValue(0);
            if (!_options.Skip(activeProject.Name))
            {
                int changeCount = ReferencesHelper.SetCopyLocalFlag(activeProject, _options.CopyLocalFlag,
                                                                    _options.PreviewMode);
                SaveProjectIfNeeded(activeProject);
                LogChangesToOutput(changeCount);
            }
            else
            {
                Common.WriteToDTEOutput(_dte, string.Format(
                    "'{0}' was skipped from processing (Set in Tools -> Options -> VSCopyLocal).",
                    activeProject.Name));
            }
        }

        /// <summary>
        /// Set Copy Local for the current solution.
        /// </summary>
        private static void SetCopyLocalForSolution()
        {
            var projects = SolutionHelper.GetProjects(_dte);
            int changeCount = 0;

            foreach (var project in projects.Where(project => !_options.Skip(project.Name)))
            {
                changeCount += ReferencesHelper.SetCopyLocalFlag(project, _options.CopyLocalFlag, _options.PreviewMode);
                if (!_options.PreviewMode)
                    SaveProjectIfNeeded(project);
            }

            LogChangesToOutput(changeCount);
        }

        private static void LogChangesToOutput(int changeCount)
        {
            var msg = changeCount > 0
                          ? string.Format(CultureInfo.CurrentCulture,
                                          "Copy Local set to {0} in {1} references{2}.", _options.CopyLocalFlag, changeCount,
                                          _options.PreviewMode ? " (Preview)" : string.Empty)
                          : "No Copy Local references found to set.";
            
            Common.WriteToDTEOutput(_dte, msg);
        }

        private static void SaveProjectIfNeeded(Project project)
        {
            if (!project.IsDirty) return;

            project.Save();
            Common.WriteToDTEOutput(_dte, string.Format("'{0}' saved.", project.Name));
        }

        private static void UpdateLabelText(OleMenuCommand cmd)
        {
            if (cmd == null) return;

            var action = _options.CopyLocalFlag ? "On" : "Off";

            if (cmd.CommandID.ID.Equals((int)PkgCmdIDList.CmdidSolutionCommand)) // solution node
            {
                cmd.Text = string.Format(CSolutionActionTextFormat, action);
            }
            if (cmd.CommandID.ID.Equals((int)PkgCmdIDList.CmdidProjectCommand)) // project node
            {
                cmd.Text = string.Format(CProjectActionTextFormat, action);
            }
            if (cmd.CommandID.ID.Equals((int)PkgCmdIDList.CmdidReferencesCommand)) // references node
            {
                cmd.Text = string.Format(CActionTextFormat, action);
            }
        }

        int IVsSolutionEvents.OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded) {
            AddHierarchyHandler(pHierarchy);

            return VSConstants.S_OK;
        }

        private void AddHierarchyHandler(IVsHierarchy pHierarchy) {
            var handler = new HierarchyHandler(pHierarchy);

            uint handlerCode;
            pHierarchy.AdviseHierarchyEvents(handler, out handlerCode);
            handler.SetCode(handlerCode);
            _hierarchyHandlers.Add(handler);
        }

        int IVsSolutionEvents.OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel) {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved) {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy) {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel) {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy) {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnAfterOpenSolution(object pUnkReserved, int fNewSolution) {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnQueryCloseSolution(object pUnkReserved, ref int pfCancel) {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnBeforeCloseSolution(object pUnkReserved) {
            return VSConstants.S_OK;
        }

        int IVsSolutionEvents.OnAfterCloseSolution(object pUnkReserved) {
            foreach (var hierarchyHandler in _hierarchyHandlers) {
                hierarchyHandler.Unadvise();
            }
            _hierarchyHandlers.Clear();

            return VSConstants.S_OK;
        }

        private class HierarchyHandler : IVsHierarchyEvents {
            private readonly IVsHierarchy _hierarchy;
            private uint _handler;

            public HierarchyHandler(IVsHierarchy hierarchy) {
                _hierarchy = hierarchy;
            }

            public void SetCode(uint code) {
                _handler = code;
            }

            public bool MatchesHierarchy(IVsHierarchy hierarchy) {
                return ReferenceEquals(hierarchy, _hierarchy);
            }

            public void Unadvise() {
                _hierarchy.UnadviseHierarchyEvents(_handler);
            }

            public int OnItemAdded(uint itemidParent, uint itemidSiblingPrev, uint itemidAdded) {
                string name;
                _hierarchy.GetCanonicalName(itemidAdded, out name);

                object propName;
                var propNameResult = _hierarchy.GetProperty(itemidAdded, (int)__VSHPROPID.VSHPROPID_Name, out propName);

                Guid propTypeGuid;
                var propTypeGuidResult = _hierarchy.GetGuidProperty(itemidAdded, (int)__VSHPROPID.VSHPROPID_TypeGuid, out propTypeGuid);

                // TODO: is there a better way to find out if this is a reference? 
                if (propNameResult == VSConstants.S_OK && propTypeGuidResult == VSConstants.S_OK) {
                    if (!string.IsNullOrEmpty((string)propName) && propTypeGuid == Guid.Empty) {
                        // This could have been a reference!
                        // TODO: only change this project, or even only this reference!
                        SetCopyLocalForSolution();
                    }
                }
                return VSConstants.S_OK;
            }

            public int OnItemsAppended(uint itemidParent) {
                return VSConstants.S_OK;
            }

            public int OnItemDeleted(uint itemid) {
                return VSConstants.S_OK;
            }

            public int OnPropertyChanged(uint itemid, int propid, uint flags) {
                return VSConstants.S_OK;
            }

            public int OnInvalidateItems(uint itemidParent) {
                return VSConstants.S_OK;
            }

            public int OnInvalidateIcon(IntPtr hicon) {
                return VSConstants.S_OK;
            }
        }
    }

    /// <summary>
    /// VSCopyLocalPackageExtensions
    /// </summary>
    public static class VSCopyLocalPackageExtensions
    {
        /// <summary>
        /// Returns a filtered list of projects.
        /// </summary>
        public static IEnumerable<Project> Filtered(this IList<Project> projects, IEnumerable<string> projectsToSkipList)
        {
            var skipList = projectsToSkipList.ToList();
            var filteredProjects = new List<Project>();

            foreach (var s in skipList)
            {
                bool wildCardStart = s.StartsWith("*");
                bool wildCardEnd = s.EndsWith("*");

                var p = s.Trim('*').ToLowerInvariant();

                if (wildCardStart && wildCardEnd)
                    filteredProjects.AddRange(projects.Where(proj => !proj.Name.ToLowerInvariant().Contains(p)));
                else if (wildCardStart)
                    filteredProjects.AddRange(projects.Where(proj => proj.Name.ToLowerInvariant().EndsWith(p)));
                else if (wildCardEnd)
                    filteredProjects.AddRange(projects.Where(proj => proj.Name.ToLowerInvariant().StartsWith(p)));
                else
                    filteredProjects.AddRange(projects.Where(proj => proj.Name.ToLowerInvariant().Equals(p)));
            }

            return projects.Except(filteredProjects);
        }
    }
}

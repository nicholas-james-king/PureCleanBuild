﻿using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;
using Task = System.Threading.Tasks.Task;

namespace PureCleanBuild
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class PureCleanBuild
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("ee7f2b7d-5dfd-4527-b998-e66e95a5da07");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="PureCleanBuild"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        /// 

        private int FoldersRemoved =  0;
        private int FilesDeleted = 0;

        private PureCleanBuild(AsyncPackage package, OleMenuCommandService commandService)
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
        public static PureCleanBuild Instance
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
            // Switch to the main thread - the call to AddCommand in PureCleanBuild's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new PureCleanBuild(package, commandService);
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
            package.JoinableTaskFactory.Run( async () => 
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                IVsSolution solution = (IVsSolution)Package.GetGlobalService(typeof(IVsSolution));
                IVsOutputWindow outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;
                Guid buildPaneGuid = VSConstants.OutputWindowPaneGuid.BuildOutputPane_guid;
                IVsOutputWindowPane buildPane;
                int projectsCleaned = 0;
                FoldersRemoved = 0;
                FilesDeleted = 0;
                outWindow.GetPane(ref buildPaneGuid, out buildPane);

                buildPane.Activate(); // Brings this pane into view
                buildPane.OutputString($"Beginning Build Nuke:\n");

                foreach (Project project in GetProjects(solution))
                {
                    if (!string.IsNullOrEmpty(project.FileName))
                    {
                        projectsCleaned++;
                        var path = project.Properties.Item("FullPath").Value.ToString();
                        var paths = new List<string>
                        { 
                            Path.Combine(path, "bin"),
                            Path.Combine(path, "obj")
                        };
                        try
                        {
                            paths.ForEach(x =>
                            {
                                package.JoinableTaskFactory.Run(async () =>
                                {
                                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                                    while (Directory.Exists(x))
                                    {
                                        DirectoryInfo di = new DirectoryInfo(x);

                                        foreach (FileInfo file in di.GetFiles())
                                        {
                                            buildPane.OutputString($"Removing file {file.FullName}\n");
                                            file.Delete();
                                            FilesDeleted++;
                                            buildPane.OutputString($"Removed file {file.FullName}\n");
                                        }
                                        foreach (DirectoryInfo dir in di.GetDirectories())
                                        {
                                            dir.Delete(true);
                                        }

                                        Directory.Delete(x);
                                        buildPane.OutputString($"Removing {x}\n");
                                        if (!Directory.Exists(x))
                                        {
                                            buildPane.OutputString($"Removed {x}\n");
                                            FoldersRemoved++;
                                        }
                                    }
                                    await Task.FromResult(true);
                                });
                            });
                        }
                        catch (UnauthorizedAccessException)
                        {
                            MessageBox.Show("Your current user does not have permissions to delete the bin/obj folders try restarting Visual Studio as administrator.", "Something has gone wrong :(");
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Something has gone wrong, please try again.");
                        }
                    }
                }
                buildPane.OutputString($"Nuke Completed, deleted {FilesDeleted} files, from {FoldersRemoved} folders inside of {projectsCleaned} projects. \n");
            });
        }
        
        public static IEnumerable<Project> GetProjects(IVsSolution solution)
        {
            foreach (IVsHierarchy hier in GetProjectsInSolution(solution))
            {
                Project project = GetDTEProject(hier);
                if (project != null)
                    yield return project;
            }
        }

        public static IEnumerable<IVsHierarchy> GetProjectsInSolution(IVsSolution solution)
        {
            return GetProjectsInSolution(solution, __VSENUMPROJFLAGS.EPF_LOADEDINSOLUTION);
        }

        public static IEnumerable<IVsHierarchy> GetProjectsInSolution(IVsSolution solution, __VSENUMPROJFLAGS flags)
        {
            if (solution == null)
                yield break;

            IEnumHierarchies enumHierarchies;
            Guid guid = Guid.Empty;
            solution.GetProjectEnum((uint)flags, ref guid, out enumHierarchies);
            if (enumHierarchies == null)
                yield break;

            IVsHierarchy[] hierarchy = new IVsHierarchy[1];
            uint fetched;
            while (enumHierarchies.Next(1, hierarchy, out fetched) == VSConstants.S_OK && fetched == 1)
            {
                if (hierarchy.Length > 0 && hierarchy[0] != null)
                    yield return hierarchy[0];
            }
        }

        public static Project GetDTEProject(IVsHierarchy hierarchy)
        {
            if (hierarchy == null)
                throw new ArgumentNullException("hierarchy");

            object obj;
            hierarchy.GetProperty(VSConstants.VSITEMID_ROOT, (int)__VSHPROPID.VSHPROPID_ExtObject, out obj);
            return obj as Project;
        }
    }
}

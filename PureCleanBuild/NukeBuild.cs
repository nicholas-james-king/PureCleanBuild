using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

using Task = System.Threading.Tasks.Task;

namespace PureCleanBuild
{
  /// <summary>
  /// Command handler
  /// </summary>
  internal sealed class NukeBuild
  {
    /// <summary>
    /// Command ID.
    /// </summary>
    public const int CommandId = 0x0100;

    /// <summary>
    /// Command menu group (command set GUID).
    /// </summary>
    public static readonly Guid CommandSet = new Guid("b24c565c-4581-4b74-8b45-7d8ff8ef285b");

    /// <summary>
    /// VS Package that provides this command, not null.
    /// </summary>
    private readonly AsyncPackage package;

    public struct Summary
    {
      public int ProjectsCleaned;
      public int FoldersRemoved;
      public int FilesDeleted;

      public override string ToString()
      {
        return $"Top Level Files Deleted: {FilesDeleted}\r\nDirectories Removed: {FoldersRemoved}\r\nProjects Cleaned: {ProjectsCleaned}\r\n";
      }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="NukeBuild"/> class.
    /// Adds our command handlers for menu (commands must exist in the command table file)
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    /// <param name="commandService">Command service to add command to, not null.</param>
    private NukeBuild(AsyncPackage package, OleMenuCommandService commandService)
    {
      this.package = package ?? throw new ArgumentNullException(nameof(package));
      commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

      var menuCommandId = new CommandID(CommandSet, CommandId);
      var menuItem = new MenuCommand(this.Execute, menuCommandId);
      commandService.AddCommand(menuItem);
    }

    /// <summary>
    /// Gets the instance of the command.
    /// </summary>
    public static NukeBuild Instance
    {
      get;
      private set;
    }

    /// <summary>
    /// Gets the service provider from the owner package.
    /// </summary>
    private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider => this.package;

    /// <summary>
    /// Initializes the singleton instance of the command.
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    public static async Task InitializeAsync(AsyncPackage package)
    {
      // Switch to the main thread - the call to AddCommand in Command1's constructor requires
      // the UI thread.
      await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

      var commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
      Instance = new NukeBuild(package, commandService);
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

      var verbose = false;

#if DEBUG
      verbose = true;
#endif

      package.JoinableTaskFactory.Run(async () =>
      {
        await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

        var solution = (await package.GetServiceAsync(typeof(IVsSolution)) as IVsSolution);
        var outputWin = (await package.GetServiceAsync(typeof(SVsOutputWindow)) as IVsOutputWindow);

        outputWin.GetPane(VSConstants.OutputWindowPaneGuid.BuildOutputPane_guid, out var buildPane);

        var summary = new Summary
        {
          ProjectsCleaned = 0,
          FoldersRemoved = 0,
          FilesDeleted = 0
        };

        buildPane.Activate();

        Action<string> output = (s) =>
        {
          buildPane.OutputString($"{s}\r\n");
        };

        try
        {
          output("--- Beginning Build Nuke ---");

          foreach (var prj in GetProjects(solution))
          {
            if (!string.IsNullOrEmpty(prj.FileName))
            {
              var paths = GetTargetPaths(prj.Properties.Item("FullPath").Value.ToString());

              try
              {
                paths.ForEach(p => 
                {
                  package.JoinableTaskFactory.Run(async () =>
                    {
                      await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                      while (Directory.Exists(p))
                      {
                        var di = new DirectoryInfo(p);

                        foreach (var file in di.GetFiles())
                        {
                          try
                          {
                            file.Delete();
                            summary.FilesDeleted++;

                            if(verbose)
                              output($"Removed file '{file.FullName}'");
                          }
                          catch (Exception ex)
                          {
                            output($"Delete Failed ({ex.Message}) for file '{file.FullName}'");
                          }
                        }

                        foreach (var dir in di.GetDirectories())
                        {
                          try
                          {
                            dir.Delete(true);
                            
                            if (verbose)
                              output($"Removed directory '{dir.FullName}'");
                          }
                          catch (Exception ex)
                          {
                            output($"Delete Failed ({ex.Message}) for directory '{dir.FullName}'");
                          }
                        }

                        try
                        {
                          Directory.Delete(p);
                          summary.FoldersRemoved++;

                          if (verbose)
                            output($"Removed directory '{p}'");
                        }
                        catch (Exception ex)
                        {
                          output($"Delete Failed ({ex.Message}) for directory '{p}'");
                        }
                      }
                    });
                });

                summary.ProjectsCleaned++;
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
          output("--- Build Nuke Completed ---");
          output(summary.ToString());
        }
        catch (Exception ex)
        {

        }
      });
    }

    public static List<string> GetTargetPaths(string path)
    {
      return new List<string>
      {
        Path.Combine(path, "bin"),
        Path.Combine(path, "obj")
      };
    }

    public static IEnumerable<Project> GetProjects(IVsSolution solution)
    {
      foreach (var h in GetProjectsInSolution(solution))
      {
        var project = GetDTEProject(h);
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

      var guid = Guid.Empty;

      solution.GetProjectEnum((uint)flags, ref guid, out var enumHierarchies);

      if (enumHierarchies == null)
        yield break;

      var hierarchy = new IVsHierarchy[1];

      while (enumHierarchies.Next(1, hierarchy, out var fetched) == VSConstants.S_OK && fetched == 1)
      {
        if (hierarchy.Length > 0 && hierarchy[0] != null)
          yield return hierarchy[0];
      }
    }

    public static Project GetDTEProject(IVsHierarchy hierarchy)
    {
      if (hierarchy == null)
        throw new ArgumentNullException(nameof(hierarchy));

      hierarchy.GetProperty(VSConstants.VSITEMID_ROOT, (int)__VSHPROPID.VSHPROPID_ExtObject, out var obj);

      return obj as Project;
    }
  }
}
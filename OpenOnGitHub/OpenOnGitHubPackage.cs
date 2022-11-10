using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace OpenOnGitHub
{
    // change to async package. ref: https://github.com/Microsoft/VSSDK-Extensibility-Samples/tree/master/AsyncPackageMigration
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112",  PackageVersion.Version, IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuids.guidOpenOnGitHubPkgString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExistsAndFullyLoaded_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class OpenOnGitHubPackage : AsyncPackage
    {
        private static DTE2 _dte;

        internal static DTE2 DTE
        {
            get
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                if (_dte == null)
                {
                    _dte = ServiceProvider.GlobalProvider.GetService(typeof(DTE)) as DTE2;
                }

                return _dte;
            }
        }

        protected override async System.Threading.Tasks.Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);

            //Switches to the UI thread in order to consume some services used in command initialization
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            //link menu items to their commands. No need to link them to the parent menu or to set their text - that's done via OpenOnGitHub.vsct
            if (await GetServiceAsync(typeof(IMenuCommandService)) is OleMenuCommandService menuCommandService)
            {
                CommandID openOnGithubCommandId = new CommandID(PackageGuids.guidOpenOnGitHubCmdSet, PackageCommandIDs.OpenOnGitHub);
                OleMenuCommand openOnGithubMenuItem = new OleMenuCommand(OpenOnGitHub, openOnGithubCommandId);
                openOnGithubMenuItem.BeforeQueryStatus += EnableDisableMenuItems;
                menuCommandService.AddCommand(openOnGithubMenuItem);

                CommandID copyToClipboardCommandId = new CommandID(PackageGuids.guidOpenOnGitHubCmdSet, PackageCommandIDs.CopyLinkToClipboard);
                OleMenuCommand copyToClipboardMenuItem = new OleMenuCommand(CopyLinkToClipboard, copyToClipboardCommandId);
                openOnGithubMenuItem.BeforeQueryStatus += EnableDisableMenuItems;
                menuCommandService.AddCommand(copyToClipboardMenuItem);

                CommandID openGithubBlameCommandId = new CommandID(PackageGuids.guidOpenOnGitHubCmdSet, PackageCommandIDs.OpenGitHubBlame);
                OleMenuCommand openGithubBlameMenuItem = new OleMenuCommand(OpenBlame, openGithubBlameCommandId);
                openOnGithubMenuItem.BeforeQueryStatus += EnableDisableMenuItems;
                menuCommandService.AddCommand(openGithubBlameMenuItem);
            }
        }

        private void EnableDisableMenuItems(object sender, EventArgs e)
        {
            OleMenuCommand menuItem = (OleMenuCommand)sender;

            //if this solution doesn't have GitHub enabled, don't create GitHub menu items
            using (GitHubRepository githubRepository = new GitHubRepository(GetActiveFilePath()))
            {
                if (!githubRepository.IsDiscoveredGitRepository)
                {
                    menuItem.Enabled = false;
                }
            }
        }

        private void OpenOnGitHub(object sender, EventArgs e)
        {
            using (GitHubRepository githubRepository = new GitHubRepository(GetActiveFilePath()))
            {
                EditorSelection editorSelection = GetEditorSelection();
                string fileUrl = githubRepository.GetFileUrl(editorSelection);
                System.Diagnostics.Process.Start(fileUrl); //open browser
            }
        }

        private void CopyLinkToClipboard(object sender, EventArgs e)
        {
            using (GitHubRepository githubRepository = new GitHubRepository(GetActiveFilePath()))
            {
                EditorSelection editorSelection = GetEditorSelection();
                string fileUrl = githubRepository.GetFileUrl(editorSelection);
                Clipboard.SetText(fileUrl);
            }
        }

        private void OpenBlame(object sender, EventArgs e)
        {
            using (GitHubRepository githubRepository = new GitHubRepository(GetActiveFilePath()))
            {
                EditorSelection editorSelection = GetEditorSelection();
                string fileUrl = githubRepository.GetBlameUrl(editorSelection);
                System.Diagnostics.Process.Start(fileUrl); //open browser
            }
        }

        private string GetActiveFilePath()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            // sometimes, DTE.ActiveDocument.Path is ToLower but GitHub can't open lower path.
            // fix proper-casing | http://stackoverflow.com/questions/325931/getting-actual-file-name-with-proper-casing-on-windows-with-net
            string activeFilePath = GetExactPathName($"{DTE.ActiveDocument.Path}{DTE.ActiveDocument.Name}");
            return activeFilePath;
        }

        private string GetExactPathName(string pathName)
        {
            if (!(File.Exists(pathName) || Directory.Exists(pathName)))
            {
                return pathName;
            }

            DirectoryInfo directoryInfo = new DirectoryInfo(pathName);

            if (directoryInfo.Parent == null)
            {
                return directoryInfo.Name.ToUpper();
            }
            else
            {
                string directoryName = GetExactPathName(directoryInfo.Parent.FullName);
                FileSystemInfo[] fileSystemInfos = directoryInfo.Parent.GetFileSystemInfos(directoryInfo.Name);
                FileSystemInfo fileSystemInfo = fileSystemInfos[0];
                string fileName = fileSystemInfo.Name;

                string exactPathName = Path.Combine(directoryName, fileName);
                return exactPathName;
            }
        }

        private EditorSelection GetEditorSelection()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!(DTE.ActiveDocument.Selection is TextSelection selection))
            {
                return null;
            }

            EditorSelection editorSelection = new EditorSelection()
            {
                StartLine = selection.TopPoint?.Line,
                EndLine = selection.BottomPoint?.Line
            };

            return editorSelection;
        }
    }
}

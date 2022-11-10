using LibGit2Sharp;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace OpenOnGitHub
{
    public sealed class GitHubRepository : IDisposable
    {
        private readonly Repository repository;
        private readonly string targetFullPath;

        public bool IsDiscoveredGitRepository => repository != null;

        public GitHubRepository(string targetFullPath)
        {
            this.targetFullPath = targetFullPath;
            string repositoryPath = LibGit2Sharp.Repository.Discover(targetFullPath);
            if (repositoryPath != null)
            {
                this.repository = new LibGit2Sharp.Repository(repositoryPath);
            }
        }

        public string GetFileUrl(EditorSelection editorSelection)
        {
            //get the base URL components
            string baseRepositoryUrl = BuildBaseRepositoryUrl();
            string commitHash = repository.Commits.First().Id.Sha;
            string filePath = BuildFileUrlPath();
            string editorSelectionUrlSelector = BuildEditorSelectionUrlSelector(editorSelection);

            //build and format the URL pieces
            baseRepositoryUrl = baseRepositoryUrl.Trim('/');
            filePath = filePath.Trim('/');

            string fileUrl = $"{baseRepositoryUrl}/blob/{commitHash}/{filePath}{editorSelectionUrlSelector}";
            return fileUrl;
        }

        public string GetBlameUrl(EditorSelection editorSelection)
        {
            //get the base URL components
            string baseRepositoryUrl = BuildBaseRepositoryUrl();
            string commitHash = repository.Commits.First().Id.Sha;
            string filePath = BuildFileUrlPath();
            string editorSelectionUrlSelector = BuildEditorSelectionUrlSelector(editorSelection);

            //build and format the URL pieces
            baseRepositoryUrl = baseRepositoryUrl.Trim('/');
            filePath = filePath.Trim('/');

            string fileUrl = $"{baseRepositoryUrl}/blame/{commitHash}/{filePath}{editorSelectionUrlSelector}";
            return fileUrl;
        }

        private string BuildBaseRepositoryUrl()
        {
            // https://github.com/user/repo.git
            ConfigurationEntry<string> baseRepoUrlEntry = repository.Config.Get<string>("remote.origin.url");
            string baseRepoUrl = baseRepoUrlEntry?.Value;

            if (baseRepoUrl == null)
            {
                throw new InvalidOperationException("GitHub repository OriginUrl can't be found");
            }

            // https://github.com/user/repo
            if (baseRepoUrl.EndsWith(".git", StringComparison.OrdinalIgnoreCase))
            {
                baseRepoUrl = baseRepoUrl.Substring(0, baseRepoUrl.Length - 4); //remove ".git"
            }

            // git@github.com:user/repo -> https://github.com/user/repo
            baseRepoUrl = Regex.Replace(baseRepoUrl,
                                        "^git@(.+):(.+)/(.+)$",
                                        match => "https://" + string.Join("/", match.Groups.OfType<Group>().Skip(1).Select(group => group.Value)),
                                        RegexOptions.IgnoreCase);

            // https://user@github.com/user/repo -> https://github.com/user/repo
            baseRepoUrl = Regex.Replace(baseRepoUrl, "(?<=^https?://)([^@/]+)@", "");

            return baseRepoUrl;
        }

        private string BuildFileUrlPath()
        {
            // foo/bar.cs
            string rootDir = repository.Info.WorkingDirectory;
            string filePath = targetFullPath.Substring(rootDir.Length).Replace("\\", "/");

            return filePath;
        }

        private string BuildEditorSelectionUrlSelector(EditorSelection editorSelection)
        {
            int? startLine = editorSelection?.StartLine;
            int? endLine = editorSelection?.EndLine;

            if (startLine.HasValue && !endLine.HasValue)
            {
                return $"#L{startLine}";
            }

            if (!startLine.HasValue && endLine.HasValue)
            {
                return $"#L{endLine}";
            }

            if (startLine.HasValue && endLine.HasValue && startLine.Value == endLine.Value)
            {
                return $"#L{startLine}";
            }

            if (startLine.HasValue && endLine.HasValue && startLine.Value < endLine.Value)
            {
                return $"#L{startLine}-L{endLine}";
            }

            return "";
        }

        void Dispose(bool disposing)
        {
            if (repository != null)
            {
                repository.Dispose();
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~GitHubRepository()
        {
            Dispose(false);
        }
    }
}
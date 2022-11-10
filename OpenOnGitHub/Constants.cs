using System;

namespace OpenOnGitHub
{
    static class PackageGuids
    {
        public const string guidOpenOnGitHubPkgString = "465b40b6-311a-4e37-9556-95fced2de9c6";
        public const string guidOpenOnGitHubCmdSetString = "a674aaec-a6f5-4df0-9749-e7bef776df5d";

        public static readonly Guid guidOpenOnGitHubCmdSet = new Guid(guidOpenOnGitHubCmdSetString);
    };

    static class PackageCommandIDs
    {
        public const int OpenOnGitHub = 0x100;
        public const int CopyLinkToClipboard = 0x200;
        public const int OpenGitHubBlame = 0x300;
    };

    static class PackageVersion
    {
        public const string Version = "1.4"; 
    }
}

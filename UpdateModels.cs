using System;
using System.IO;
using System.Runtime.Serialization;

namespace JarvisPowerPoint
{
    [DataContract]
    internal sealed class GitHubRelease
    {
        [DataMember(Name = "tag_name", IsRequired = true)] public string Tag { get; set; }
        [DataMember(Name = "html_url", IsRequired = true)] public string Url { get; set; }
        [DataMember(Name = "draft", IsRequired = true)] public bool Draft { get; set; }
        [DataMember(Name = "prerelease", IsRequired = true)] public bool Prerelease { get; set; }
        [DataMember(Name = "body", IsRequired = true)] public string Notes { get; set; }
        [DataMember(Name = "assets", IsRequired = true)] public GitHubAsset[] Assets { get; set; }
    }

    [DataContract]
    internal sealed class GitHubAsset
    {
        [DataMember(Name = "name", IsRequired = true)] public string Name { get; set; }
        [DataMember(Name = "browser_download_url", IsRequired = true)] public string Url { get; set; }
        [DataMember(Name = "size", IsRequired = true)] public long Size { get; set; }
        [DataMember(Name = "digest", IsRequired = true)] public string Digest { get; set; }
    }

    internal sealed class UpdateRelease
    {
        public string Tag { get; private set; }
        public Version Version { get; private set; }
        public string Notes { get; private set; }
        public Uri DownloadUri { get; private set; }
        public long Size { get; private set; }
        public string Sha256 { get; private set; }

        public UpdateRelease(string tag, Version version, string notes, Uri uri, long size, string sha256)
        {
            Tag = tag;
            Version = version;
            Notes = notes;
            DownloadUri = uri;
            Size = size;
            Sha256 = sha256;
        }
    }

    [DataContract]
    internal sealed class UpdateManifest
    {
        [DataMember(IsRequired = true)] public string Target;
        [DataMember(IsRequired = true)] public string OldHash;
        [DataMember(IsRequired = true)] public int ParentId;
        [DataMember(IsRequired = true)] public long ParentStartUtcTicks;
        [DataMember(IsRequired = true)] public string Tag;
        [DataMember(IsRequired = true)] public long Size;
        [DataMember(IsRequired = true)] public string Sha256;
    }

    internal sealed class PreparedUpdate
    {
        public string DirectoryPath { get; private set; }
        public UpdateManifest Manifest { get; private set; }
        public bool HelperStarted { get; internal set; }
        public bool InstallAuthorized { get; internal set; }

        public PreparedUpdate(string directoryPath, UpdateManifest manifest)
        {
            DirectoryPath = directoryPath;
            Manifest = manifest;
        }
    }

    internal sealed class UpdateResponse : IDisposable
    {
        public int StatusCode;
        public string Location;
        public long ContentLength;
        public Stream Body;
        public Action Close;

        public void Dispose()
        {
            try { if (Body != null) Body.Dispose(); }
            finally { if (Close != null) Close(); }
        }
    }

    internal interface IUpdateTransport
    {
        UpdateResponse Get(Uri uri, System.Threading.CancellationToken token);
    }

    internal interface IUpdateFiles
    {
        bool Exists(string path);
        string Hash(string path);
        Version FileVersion(string path);
        long Length(string path);
        void Copy(string from, string to);
        void Replace(string from, string target, string backup);
        void Move(string from, string to);
        void Delete(string path);
    }

    internal sealed class UpdateProcessIdentity
    {
        public int Id;
        public long StartUtcTicks;
        public string Executable;
    }
}

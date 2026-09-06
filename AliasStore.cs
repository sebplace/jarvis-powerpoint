using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Xml;

namespace JarvisPowerPoint
{
    internal sealed class AliasStore
    {
        private readonly string directory;

        public AliasStore(string dataDirectory)
        {
            if (string.IsNullOrWhiteSpace(dataDirectory))
                throw new ArgumentException("A local data directory is required.", "dataDirectory");
            directory = Path.Combine(Path.GetFullPath(dataDirectory), "Aliases");
        }

        internal static string CanonicalPath(string presentationPath)
        {
            Uri uri;
            if (Uri.TryCreate(presentationPath, UriKind.Absolute, out uri) &&
                (uri.Scheme == Uri.UriSchemeHttps || uri.Scheme == Uri.UriSchemeHttp))
                return uri.AbsoluteUri;
            return Path.GetFullPath(presentationPath).ToUpperInvariant();
        }

        internal static string PathKey(string presentationPath)
        {
            using (SHA256 hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(
                    Encoding.UTF8.GetBytes(CanonicalPath(presentationPath)))).Replace("-", "");
        }

        public List<SlideAlias> Load(string presentationPath)
        {
            string path = StorePath(presentationPath);
            var aliases = new List<SlideAlias>();
            // File.Exists/Directory.Exists hide access failures.
            FileAttributes attributes;
            try { attributes = File.GetAttributes(directory); }
            catch (FileNotFoundException) { return aliases; }
            catch (DirectoryNotFoundException) { return aliases; }
            if ((attributes & FileAttributes.Directory) == 0)
            {
                throw new IOException("The alias directory is a file.");
            }
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = 1024 * 1024
            };
            XmlDocument document = new XmlDocument { XmlResolver = null };
            try
            {
                using (XmlReader reader = XmlReader.Create(path, settings))
                    document.Load(reader);
            }
            catch (FileNotFoundException)
            {
                return aliases;
            }
            XmlElement root = document.DocumentElement;
            if (root == null || root.Name != "slideAliases" ||
                root.GetAttribute("version") != "1" ||
                !string.Equals(root.GetAttribute("presentation"),
                    CanonicalPath(presentationPath), StringComparison.Ordinal))
                throw new InvalidDataException("The alias file does not belong to this presentation.");

            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (XmlNode node in root.ChildNodes)
            {
                XmlElement element = node as XmlElement;
                if (element == null) continue;
                int slideId;
                string name = element.GetAttribute("name");
                if (element.Name != "alias" || string.IsNullOrWhiteSpace(name) ||
                    name.Length > 80 || name != name.Trim() || !names.Add(name) ||
                    !int.TryParse(element.GetAttribute("slideId"), NumberStyles.None,
                        CultureInfo.InvariantCulture, out slideId) || slideId <= 0)
                    throw new InvalidDataException("The alias file contains an invalid alias.");
                aliases.Add(new SlideAlias
                {
                    Name = name,
                    SlideId = slideId,
                    Title = element.GetAttribute("title")
                });
            }
            return aliases;
        }

        public void Save(string presentationPath, List<SlideAlias> aliases)
        {
            Directory.CreateDirectory(directory);
            string path = StorePath(presentationPath);
            string pending = Path.Combine(directory, Guid.NewGuid().ToString("N") + ".new");
            try
            {
                var settings = new XmlWriterSettings
                {
                    Encoding = new UTF8Encoding(false),
                    Indent = true,
                    CloseOutput = false
                };
                using (var stream = new FileStream(pending, FileMode.CreateNew,
                    FileAccess.Write, FileShare.None))
                {
                    using (XmlWriter writer = XmlWriter.Create(stream, settings))
                    {
                        writer.WriteStartElement("slideAliases");
                        writer.WriteAttributeString("version", "1");
                        writer.WriteAttributeString("presentation", CanonicalPath(presentationPath));
                        foreach (SlideAlias alias in aliases)
                        {
                            writer.WriteStartElement("alias");
                            writer.WriteAttributeString("name", alias.Name);
                            writer.WriteAttributeString("slideId", alias.SlideId.ToString(CultureInfo.InvariantCulture));
                            writer.WriteAttributeString("title", alias.Title ?? string.Empty);
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();
                    }
                    stream.Flush(true);
                }
                if (File.Exists(path))
                    File.Replace(pending, path, null);
                else
                    File.Move(pending, path);
            }
            finally
            {
                if (File.Exists(pending)) File.Delete(pending);
            }
        }

        private string StorePath(string presentationPath)
        {
            return Path.Combine(directory, PathKey(presentationPath) + ".xml");
        }
    }
}

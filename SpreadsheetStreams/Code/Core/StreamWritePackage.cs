using SpreadsheetStreams.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

namespace SpreadsheetStreams
{
    internal class PackageWriteStream : IDisposable
    {
        private ZipArchive _ZipArchive = null;
        private List<Relationship> _PackageRelationships = new List<Relationship>();
        private List<ContentType> _ContentTypes = new List<ContentType>();
        private Dictionary<string, List<Relationship>> _PartRelationships = new Dictionary<string, List<Relationship>>();

        internal PackageWriteStream(Stream outputStream)
        {
            _ZipArchive = new ZipArchive(outputStream, ZipArchiveMode.Create);
        }

        #region IDisposable

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_ZipArchive != null)
                {
                    _ZipArchive.Dispose();
                    _ZipArchive = null;
                }
            }
        }

        #endregion

        internal void Close()
        {
            if (_ZipArchive != null)
            {
                _ZipArchive.Dispose();
                _ZipArchive = null;
            }
        }

        internal ZipArchiveEntry CreateStream(string name, CompressionLevel compressionLevel)
        {
            return _ZipArchive.CreateEntry(name.TrimStart('/'), compressionLevel);
        }

        internal void AddPackageRelationship(string target, string type, string id)
        {
            _PackageRelationships.Add(new Relationship { Target = target, Type = type, Id = id });
        }

        internal void AddPartRelationship(string fromPath, string target, string type, string id)
        {
            if (!_PartRelationships.TryGetValue(fromPath, out var rels))
            {
                rels = new List<Relationship>();
                _PartRelationships[fromPath] = rels;
            }

            rels.Add(new Relationship { Target = target, Type = type, Id = id });
        }

        internal void AddContentType(string target, string type)
        {
            _ContentTypes.Add(new ContentType { Target = target, Type = type });
        }

        internal void CommitRelationships(CompressionLevel compressionLevel)
        {
            if (_PackageRelationships != null)
            {
                var pEntry = CreateStream("_rels/.rels", compressionLevel);
                using (var stream = pEntry.Open())
                using (var streamWriter = new StreamWriter(stream))
                {
                    streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                    streamWriter.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                    {
                        foreach (var rel in _PackageRelationships)
                        {
                            streamWriter.Write($"<Relationship Type=\"{XmlHelper.Escape(rel.Type)}\" Target=\"{XmlHelper.Escape(rel.Target)}\"{(string.IsNullOrEmpty(rel.Id) ? "" : $" Id=\"{XmlHelper.Escape(rel.Id)}\"")}/>");
                        }
                    }
                    streamWriter.Write("</Relationships>");
                }
            }

            foreach (var p in _PartRelationships)
            {
                var basePath = Path.GetDirectoryName(p.Key).Replace('\\', '/');
                var relsOwner = Path.GetFileName(p.Key);

                var pEntry = CreateStream($"{basePath}/_rels/{relsOwner}.rels", compressionLevel);
                using (var stream = pEntry.Open())
                using (var streamWriter = new StreamWriter(stream))
                {
                    streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                    streamWriter.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
                    {
                        foreach (var rel in p.Value)
                        {
                            var target = rel.Target;
                            if (target.StartsWith(basePath))
                            {
                                target = target.Remove(0, basePath.Length);
                                while (target.StartsWith("/"))
                                    target = target.Remove(0, 1);
                            }

                            streamWriter.Write($"<Relationship Type=\"{XmlHelper.Escape(rel.Type)}\" Target=\"{XmlHelper.Escape(target)}\"{(string.IsNullOrEmpty(rel.Id) ? "" : $" Id=\"{XmlHelper.Escape(rel.Id)}\"")}/>");
                        }
                    }
                    streamWriter.Write("</Relationships>");
                }
            }
        }

        internal void CommitContentTypes(CompressionLevel compressionLevel)
        {
            if (_ContentTypes != null)
            {
                var ctEntry = CreateStream("[Content_Types].xml", compressionLevel);
                using (var stream = ctEntry.Open())
                using (var streamWriter = new StreamWriter(stream))
                {
                    streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                    streamWriter.Write("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
                    {
                        streamWriter.Write("<Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
                        streamWriter.Write("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");

                        foreach (var ct in _ContentTypes)
                        {
                            streamWriter.Write($"<Override PartName=\"{XmlHelper.Escape(ct.Target)}\" ContentType=\"{XmlHelper.Escape(ct.Type)}\"/>");
                        }
                    }
                    streamWriter.Write("</Types>");
                }
            }
        }

        private class Relationship
        {
            internal string Type;
            internal string Target;
            internal string Id;
        }

        private class ContentType
        {
            internal string Type;
            internal string Target;
        }
    }
}

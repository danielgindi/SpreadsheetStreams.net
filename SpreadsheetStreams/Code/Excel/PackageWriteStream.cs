using SpreadsheetStreams.Util;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace SpreadsheetStreams.Code.Excel
{
    internal class PackageWriteStream : IDisposable
    {
        private XmlWriterHelper _XmlWriterHelper = new XmlWriterHelper();
        private ZipArchive _ZipArchive = null;
        private List<Relationship> _PackageRelationships = new List<Relationship>();
        private List<ContentType> _ContentTypes = new List<ContentType>();
        private Dictionary<string, List<Relationship>> _PartRelationships = new Dictionary<string, List<Relationship>>();
        private ImageStore _ImageStore = new ImageStore();

        internal PackageWriteStream(Stream outputStream, bool leaveOpen)
        {
            _ZipArchive = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen);
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

                if (_XmlWriterHelper != null)
                {
                    _XmlWriterHelper.Dispose();
                    _XmlWriterHelper = null;
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

        internal ZipArchiveEntry CreateEntry(string name, CompressionLevel compressionLevel)
        {
            return _ZipArchive.CreateEntry(name.TrimStart('/'), compressionLevel);
        }

        internal void AddPackageRelationship(string target, string type)
        {
            _PackageRelationships.Add(new Relationship { Target = target, Type = type });
        }

        internal int AddPartRelationship(string fromPath, string target, string type)
        {
            if (!_PartRelationships.TryGetValue(fromPath, out var rels))
            {
                rels = new List<Relationship>();
                _PartRelationships[fromPath] = rels;
            }

            rels.Add(new Relationship { Target = target, Type = type });
            return rels.Count;
        }

        internal void AddContentType(string target, string type)
        {
            _ContentTypes.Add(new ContentType { Target = target, Type = type });
        }

        internal async Task<int> AddImageAsync(Image image, CancellationToken cancellationToken)
        {
            var meta = await _ImageStore.AddImageAsync(image, this, cancellationToken);

            AddPartRelationship(
                fromPath: "xl/richData/richValueRel.xml",
                target: meta.Path,
                type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");

            return meta.Index;
        }

        private const string _PART_RDARRAY_PATH = "/xl/richData/rdarray.xml";
        private const string _PART_RDRICHVALUE_PATH = "/xl/richData/rdrichvalue.xml";
        private const string _PART_RDRICHVALUETYPES_PATH = "/xl/richData/rdRichValueTypes.xml";
        private const string _PART_RDRICHVALUEREL_PATH = "/xl/richData/richValueRel.xml";
        private const string _PART_METADATA_PATH = "/xl/metadata.xml";

        internal async Task CommitRichDataAsync(CompressionLevel compressionLevel)
        {
            if (_ImageStore.ImageList.Count == 0)
                return;

            var imageList = _ImageStore.ImageList;

            var pEntry = CreateEntry(_PART_RDARRAY_PATH, compressionLevel);
            using (var stream = pEntry.Open())
            using (var streamWriter = new StreamWriter(stream))
            {
                await streamWriter.WriteAsync("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>").ConfigureAwait(false);
                await streamWriter.WriteAsync("<arrayData xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2\" count=\"0\"></arrayData>").ConfigureAwait(false);
            }
            AddPartRelationship("/xl/workbook.xml", _PART_RDARRAY_PATH, "http://schemas.microsoft.com/office/2017/06/relationships/rdArray");
            AddContentType(_PART_RDARRAY_PATH, "application/vnd.ms-excel.rdarray+xml");

            var richValueStructureStore = new RichValueStructureStore();
            var imageStructureIndex = richValueStructureStore.AddStructure(
                new RichValueStructure(
                    RichValueStructureType.LocalImage,
                    new List<RichValueStructureKey>
                    {
                         new RichValueStructureKey("_rvRel:LocalImageIdentifier", RichValueDataType.Integer),
                         new RichValueStructureKey("CalcOrigin", RichValueDataType.Integer),
                         //new RichValueStructureKey("Text", RichValueDataType.String),
                    }
                )
            );
            pEntry = CreateEntry(RichValueStructureStore.PART_PATH, compressionLevel);
            using (var stream = pEntry.Open())
            using (var streamWriter = new StreamWriter(stream))
            {
                await richValueStructureStore.WriteXmlToStream(streamWriter);
            }
            AddPartRelationship("/xl/workbook.xml", RichValueStructureStore.PART_PATH, "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure");
            AddContentType(RichValueStructureStore.PART_PATH, "application/vnd.ms-excel.rdrichvaluestructure+xml");

            pEntry = CreateEntry(_PART_RDRICHVALUE_PATH, compressionLevel);
            using (var stream = pEntry.Open())
            using (var streamWriter = new StreamWriter(stream))
            {
                await streamWriter.WriteAsync("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>").ConfigureAwait(false);
                await streamWriter.WriteAsync($"<rvData xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" count=\"{imageList.Count}\">").ConfigureAwait(false);
                foreach (var image in imageList)
                {
                    // For now, support the CalcOrigin=Standalone mode only
                    await streamWriter.WriteAsync($"<rv s=\"{imageStructureIndex}\"><v>{image.Index}</v><v>5</v></rv>");
                }
                await streamWriter.WriteAsync($"</rvData>");
            }
            AddPartRelationship("/xl/workbook.xml", _PART_RDRICHVALUE_PATH, "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue");
            AddContentType(_PART_RDRICHVALUE_PATH, "application/vnd.ms-excel.rdrichvalue+xml");

            pEntry = CreateEntry(_PART_RDRICHVALUEREL_PATH, compressionLevel);
            using (var stream = pEntry.Open())
            using (var streamWriter = new StreamWriter(stream))
            {
                await streamWriter.WriteAsync("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>").ConfigureAwait(false);
                await streamWriter.WriteAsync($"<richValueRels xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">").ConfigureAwait(false);
                foreach (var image in imageList)
                {
                    await streamWriter.WriteAsync($"<rel r:id=\"rId{image.Index + 1}\" />");
                }
                await streamWriter.WriteAsync($"</richValueRels>");
            }
            AddPartRelationship("/xl/workbook.xml", _PART_RDRICHVALUEREL_PATH, "http://schemas.microsoft.com/office/2022/10/relationships/richValueRel");
            AddContentType(_PART_RDRICHVALUEREL_PATH, "application/vnd.ms-excel.richvaluerel+xml");

            pEntry = CreateEntry(_PART_RDRICHVALUETYPES_PATH, compressionLevel);
            using (var stream = pEntry.Open())
            using (var streamWriter = new StreamWriter(stream))
            {
                await streamWriter.WriteAsync("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>").ConfigureAwait(false);
                await streamWriter.WriteAsync($"<rvTypesInfo xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" mc:Ignorable=\"x\">").ConfigureAwait(false);
                await streamWriter.WriteAsync($"<global>");
                await streamWriter.WriteAsync($"<keyFlags>");
                await streamWriter.WriteAsync($"<key name=\"_Self\"><flag name=\"ExcludeFromFile\" value=\"1\" /><flag name=\"ExcludeFromCalcComparison\" value=\"1\" /></key>\r\n<key name=\"_DisplayString\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\" /></key>");
                await streamWriter.WriteAsync($"<key name=\"_Flags\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\" /></key>");
                await streamWriter.WriteAsync($"<key name=\"_Format\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\" /></key>");
                await streamWriter.WriteAsync($"<key name=\"_SubLabel\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\" /></key>");
                await streamWriter.WriteAsync($"<key name=\"_Attribution\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\" /></key>");
                await streamWriter.WriteAsync($"<key name=\"_Icon\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\" /></key>");
                await streamWriter.WriteAsync($"<key name=\"_Display\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\" /></key>");
                await streamWriter.WriteAsync($"<key name=\"_CanonicalPropertyNames\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\" /></key>");
                await streamWriter.WriteAsync($"<key name=\"_ClassificationId\"><flag name=\"ExcludeFromCalcComparison\" value=\"1\" /></key>");
                await streamWriter.WriteAsync($"</keyFlags>");
                await streamWriter.WriteAsync($"</global>");
                await streamWriter.WriteAsync($"<types></types>");
                await streamWriter.WriteAsync($"</rvTypesInfo>");
            }
            AddPartRelationship("/xl/workbook.xml", _PART_RDRICHVALUETYPES_PATH, "http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes");
            AddContentType(_PART_RDRICHVALUETYPES_PATH, "application/vnd.ms-excel.rdrichvaluetypes+xml");

            pEntry = CreateEntry(_PART_METADATA_PATH, compressionLevel);
            using (var stream = pEntry.Open())
            using (var streamWriter = new StreamWriter(stream))
            {
                await streamWriter.WriteAsync("<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?>").ConfigureAwait(false);
                await streamWriter.WriteAsync($"<metadata xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:xlrd=\"http://schemas.microsoft.com/office/spreadsheetml/2017/richdata\" xmlns:xda=\"http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray\">").ConfigureAwait(false);
                await streamWriter.WriteAsync($"<metadataTypes count=\"1\"><metadataType name=\"XLRICHVALUE\" minSupportedVersion=\"120000\"  copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\" /></metadataTypes>");
                await streamWriter.WriteAsync($"<futureMetadata name=\"XLRICHVALUE\" count=\"{imageList.Count}\">");
                foreach (var image in imageList)
                {
                    await streamWriter.WriteAsync($"<bk><extLst><ext uri=\"{{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}}\"><xlrd:rvb i=\"{image.Index}\" /></ext></extLst></bk>");
                }
                await streamWriter.WriteAsync($"</futureMetadata>");
                await streamWriter.WriteAsync($"<valueMetadata count=\"{imageList.Count}\">");
                foreach (var image in imageList)
                {
                    await streamWriter.WriteAsync($"<bk><rc t=\"1\" v=\"{image.Index}\"/></bk>");
                }
                await streamWriter.WriteAsync($"</valueMetadata>");
                await streamWriter.WriteAsync($"</metadata>");
            }
            AddPartRelationship("/xl/workbook.xml", _PART_METADATA_PATH, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata");
            AddContentType(_PART_METADATA_PATH, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml");
        }

        internal async Task CommitRelationshipsAsync(CompressionLevel compressionLevel)
        {
            if (_PackageRelationships != null)
            {
                var pEntry = CreateEntry("_rels/.rels", compressionLevel);
                using (var stream = pEntry.Open())
                using (var streamWriter = new StreamWriter(stream))
                {
                    await streamWriter.WriteAsync("<?xml version=\"1.0\" encoding=\"utf-8\"?>").ConfigureAwait(false);
                    await streamWriter.WriteAsync("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">").ConfigureAwait(false);
                    {
                        int id = 1;
                        foreach (var rel in _PackageRelationships)
                        {
                            await streamWriter.WriteAsync($"<Relationship Type=\"{XmlWriterHelper.EscapeSimpleXmlAttr(rel.Type)}\" Target=\"{XmlWriterHelper.EscapeSimpleXmlAttr(rel.Target)}\" Id=\"rId{id}\"/>").ConfigureAwait(false);
                            id++;
                        }
                    }
                    await streamWriter.WriteAsync("</Relationships>").ConfigureAwait(false);
                }
            }

            foreach (var p in _PartRelationships)
            {
                var basePath = Path.GetDirectoryName(p.Key).Replace('\\', '/');
                var relsOwner = Path.GetFileName(p.Key);

                var pEntry = CreateEntry($"{basePath}/_rels/{relsOwner}.rels", compressionLevel);
                using (var stream = pEntry.Open())
                using (var streamWriter = new StreamWriter(stream))
                {
                    await streamWriter.WriteAsync("<?xml version=\"1.0\" encoding=\"utf-8\"?>").ConfigureAwait(false);
                    await streamWriter.WriteAsync("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">").ConfigureAwait(false);
                    {
                        int id = 1;

                        foreach (var rel in p.Value)
                        {
                            var target = rel.Target;
                            if (target.StartsWith(basePath))
                            {
                                target = target.Remove(0, basePath.Length);
                                while (target.StartsWith("/"))
                                    target = target.Remove(0, 1);
                            }

                            await streamWriter.WriteAsync($"<Relationship Type=\"{XmlWriterHelper.EscapeSimpleXmlAttr(rel.Type)}\" Target=\"{XmlWriterHelper.EscapeSimpleXmlAttr(target)}\" Id=\"rId{id}\"/>").ConfigureAwait(false);
                            id++;
                        }
                    }
                    await streamWriter.WriteAsync("</Relationships>").ConfigureAwait(false);
                }
            }
        }

        internal async Task CommitContentTypesAsync(CompressionLevel compressionLevel)
        {
            if (_ContentTypes != null)
            {
                var ctEntry = CreateEntry("[Content_Types].xml", compressionLevel);
                using (var stream = ctEntry.Open())
                using (var streamWriter = new StreamWriter(stream))
                {
                    await streamWriter.WriteAsync("<?xml version=\"1.0\" encoding=\"utf-8\"?>").ConfigureAwait(false);
                    await streamWriter.WriteAsync("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">").ConfigureAwait(false);
                    {
                        await streamWriter.WriteAsync("<Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>").ConfigureAwait(false);
                        await streamWriter.WriteAsync("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>").ConfigureAwait(false);
                        await streamWriter.WriteAsync("<Default Extension=\"png\" ContentType=\"image/png\"/>").ConfigureAwait(false);
                        await streamWriter.WriteAsync("<Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>").ConfigureAwait(false);
                        await streamWriter.WriteAsync("<Default Extension=\"gif\" ContentType=\"image/gif\"/>").ConfigureAwait(false);
                        await streamWriter.WriteAsync("<Default Extension=\"tif\" ContentType=\"image/tif\"/>").ConfigureAwait(false);
                        await streamWriter.WriteAsync("<Default Extension=\"bmp\" ContentType=\"image/bmp\"/>").ConfigureAwait(false);
                        await streamWriter.WriteAsync("<Default Extension=\"ico\" ContentType=\"image/ico\"/>").ConfigureAwait(false);

                        foreach (var ct in _ContentTypes)
                        {
                            await streamWriter.WriteAsync($"<Override PartName=\"{XmlWriterHelper.EscapeSimpleXmlAttr(ct.Target)}\" ContentType=\"{XmlWriterHelper.EscapeSimpleXmlAttr(ct.Type)}\"/>").ConfigureAwait(false);
                        }
                    }
                    await streamWriter.WriteAsync("</Types>").ConfigureAwait(false);
                }
            }
        }

        private class Relationship
        {
            internal string Type;
            internal string Target;
        }

        private class ContentType
        {
            internal string Type;
            internal string Target;
        }
    }
}

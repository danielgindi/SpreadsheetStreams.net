using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

#nullable enable

namespace SpreadsheetStreams.Code.Excel
{
    internal class ImageStore
    {
        private Dictionary<Image, ImageWithMeta> ImageMetaMap = new Dictionary<Image, ImageWithMeta>();
        private List<ImageWithMeta> _ImageList = new List<ImageWithMeta>();
        private HashSet<string> UsedFileNames = new HashSet<string>();

        public List<ImageWithMeta> ImageList
        {
            get { return _ImageList; }
        }

        internal async Task<ImageWithMeta> AddImageAsync(Image image, PackageWriteStream package, CancellationToken cancellationToken)
        {
            if (ImageMetaMap.TryGetValue(image, out var meta))
            {
                return meta;
            }

            string ext;
            if (image.ContentType == "image/jpeg")
                ext = ".jpg";
            else if (image.ContentType == "image/png")
                ext = ".png";
            else if (image.ContentType == "image/gif")
                ext = ".gif";
            else if (image.ContentType == "image/bmp")
                ext = ".bmp";
            else if (image.ContentType == "image/webp")
                ext = ".webp";
            else if (image.ContentType == "image/tiff")
                ext = ".tif";
            else if (image.ContentType == "image/ico")
                ext = ".ico";
            else throw new System.ArgumentException("`image` must have a supported ContentType");

            var fn = GenerateFileName(ext);
            var path = "xl/media/" + fn;
            meta = new ImageWithMeta(image, "/" + path);

            var entry = package.CreateEntry(path, GetCompressionlevel(ext));
            using var stream = entry.Open();
            if (image.Stream != null)
            {
                await image.Stream.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
            }
            else if (image.Data != null)
            {
                await stream.WriteAsync(image.Data, 0, image.Data.Length, cancellationToken).ConfigureAwait(false);
            }
            else if (image.Path != null)
            {
                using var file = System.IO.File.Open(image.Path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read);
                await file.CopyToAsync(stream, cancellationToken).ConfigureAwait(false);
            }
            else
            {
                throw new System.ArgumentException("The specified Image contains no data");
            }

            var index = ImageList.Count;
            meta.Index = index;
            ImageMetaMap[image] = meta;
            ImageList.Add(meta);

            return meta;
        }

        private string GenerateFileName(string ext)
        {
            string suggested;
            int index = ImageList.Count;
            int extra = 0;

            do
            {
                suggested = "_rv_" + index;

                if (extra > 0)
                    suggested += "_" + extra;

                suggested += ext;

                extra++;
            } while (UsedFileNames.Contains(suggested));
            return suggested;
        }

        private CompressionLevel GetCompressionlevel(string ext)
        {
            switch (ext)
            {
                case ".bmp":
                case ".tif":
                    return CompressionLevel.Optimal;
                default:
                    return CompressionLevel.Fastest;
            }
        }

        internal class ImageWithMeta
        {
            internal Image Image;
            internal string Path;
            internal int Index;

            internal ImageWithMeta(Image image, string path)
            {
                this.Image = image;
                this.Path = path;
                this.Index = -1;
            }
        }
    }
}

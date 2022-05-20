using System;

#nullable enable

namespace SpreadsheetStreams
{
    public class SpreadsheetInfo
    {
        public string? Title { get; set; }

        public string? Subject { get; set; }

        /// <summary>
        /// Semicolon delimited
        /// </summary>
        public string? Author { get; set; }

        public string? Keywords { get; set; }

        public string? Comments { get; set; }

        public string? Status { get; set; }

        public string? Category { get; set; }

        public string? LastModifiedBy { get; set; }

        public DateTime? CreatedOn { get; set; }

        public DateTime? ModifiedOn { get; set; }

        public string? Application { get; set; }

        /// <summary>
        /// In excel: XX.YYYY
        /// </summary>
        public string? AppVersion { get; set; }

        public bool? ScaleCrop { get; set; }

        public string? Manager { get; set; }

        public string? Company { get; set; }

        public bool? LinksUpToDate { get; set; }

        public bool? SharedDoc { get; set; }

        public bool? HyperlinksChanged { get; set; }
    }
}

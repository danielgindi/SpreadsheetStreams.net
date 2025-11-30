using System;
using System.Collections.Generic;

namespace SpreadsheetStreams
{
    public class WorksheetInfo
    {
        internal int Id;
        internal string Path;

        public string Name;
        public List<ColumnInfo> ColumnInfos;

        [Obsolete("Use ColumnInfos instead")]
        public float[] ColumnWidths
        {
            set
            {
                ColumnInfos = new List<ColumnInfo>();

                if (value != null)
                {
                    for (int i = 0; i < value.Length; i++)
                    {
                        ColumnInfos.Add(new ColumnInfo
                        {
                            FromColumn = i,
                            ToColumn = i,
                            Width = value[i],
                        });
                    }
                }
            }
        }

        public float? DefaultRowHeight = 15f;
        public float? DefaultColumnWidth = 10f;
        public bool? RightToLeft = null;
    }
}

using System;
using System.Collections.Generic;
using System.Drawing;

namespace SpreadsheetStreams
{
    public struct Border : IEquatable<Border>
    {
        public BorderPosition Position;
        public Color Color;
        public BorderLineStyle LineStyle;
        public float Weight;

        public Border(
            BorderPosition position,
            Color? color = null,
            BorderLineStyle lineStyle = BorderLineStyle.Continuous,
            float weight = 1.0f)
        {
            this.Position = position;
            this.Color = color ?? Color.Black;
            this.LineStyle = lineStyle;
            this.Weight = weight;
        }

        public bool Equals(Border b)
          => b.Position == Position &&
            b.Color == Color &&
            b.LineStyle == LineStyle &&
            b.Weight == Weight;

        public override bool Equals(object o)
          => o is Border b && b.Equals(this);

        public override int GetHashCode()
        {
            var hashCode = -1582263389;
            hashCode = hashCode * -1521134295 + Position.GetHashCode();
            hashCode = hashCode * -1521134295 + Color.GetHashCode();
            hashCode = hashCode * -1521134295 + LineStyle.GetHashCode();
            hashCode = hashCode * -1521134295 + Weight.GetHashCode();
            return hashCode;
        }

        public static bool operator ==(Border first, Border second) => Equals(first, second);

        public static bool operator !=(Border first, Border second) => !Equals(first, second);
    }
}
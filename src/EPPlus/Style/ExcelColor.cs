/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Jan Källman		                Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Style.XmlAccess;
using System.Drawing;
using System.Globalization;

namespace OfficeOpenXml.Style
{
    /// <summary>
    /// Color for cellstyling
    /// </summary>
    public sealed class ExcelColor: StyleBase, IColor
    {
        eStyleClass _cls;
        StyleBase _parent;
        internal ExcelColor (ExcelStyles styles, OfficeOpenXml.XmlHelper.ChangedEventHandler ChangedEvent, int worksheetID, string address, eStyleClass cls, StyleBase parent) :
            base (styles, ChangedEvent, worksheetID, address)
        {
            _parent = parent;
            _cls = cls;
        }
        /// <summary>
        /// The theme color
        /// </summary>
        public string Theme
        {
            get
            {
                return GetSource ().Theme;
            }
        }
        /// <summary>
        /// The tint value
        /// </summary>
        public decimal Tint
        {
            get
            {
                return GetSource ().Tint;
            }
            set
            {
                if (value > 1 || value < -1)
                {
                    throw (new ArgumentOutOfRangeException ("Value must be between -1 and 1"));
                }
                _ChangedEvent (this, new StyleChangeEventArgs (_cls, eStyleProperty.Tint, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The RGB value
        /// </summary>
        public string Rgb
        {
            get
            {
                return GetSource ().Rgb;
            }
            internal set
            {
                _ChangedEvent (this, new StyleChangeEventArgs (_cls, eStyleProperty.Color, value, _positionID, _address));
            }
        }
        /// <summary>
        /// The indexed color number.
        /// </summary>
        public int Indexed
        {
            get
            {
                return GetSource ().Indexed;
            }
            set
            {
                _ChangedEvent (this, new StyleChangeEventArgs (_cls, eStyleProperty.IndexedColor, value, _positionID, _address));
            }
        }
        /// <summary>
        /// Set the color of the object
        /// </summary>
        /// <param name="color">The color</param>
        public void SetColor (Color color)
        {
            Rgb = color.ToArgb ().ToString ("X");
        }
        /// <summary>
        /// Set the color of the object
        /// </summary>
        /// <param name="alpha">Alpha component value</param>
        /// <param name="red">Red component value</param>
        /// <param name="green">Green component value</param>
        /// <param name="blue">Blue component value</param>
        public void SetColor (int alpha, int red, int green, int blue)
        {
            if (alpha < 0 || red < 0 || green < 0 || blue < 0 ||
               alpha > 255 || red > 255 || green > 255 || blue > 255)
            {
                throw (new ArgumentException ("Argument range must be from 0 to 255"));
            }
            Rgb = alpha.ToString ("X2") + red.ToString ("X2") + green.ToString ("X2") + blue.ToString ("X2");
        }
        internal override string Id
        {
            get
            {
                return Theme + Tint + Rgb + Indexed;
            }
        }
        private ExcelColorXml GetSource ()
        {
            Index = _parent.Index < 0 ? 0 : _parent.Index;
            switch (_cls)
            {
            case eStyleClass.FillBackgroundColor:
                return _styles.Fills[Index].BackgroundColor;
            case eStyleClass.FillPatternColor:
                return _styles.Fills[Index].PatternColor;
            case eStyleClass.Font:
                return _styles.Fonts[Index].Color;
            case eStyleClass.BorderLeft:
                return _styles.Borders[Index].Left.Color;
            case eStyleClass.BorderTop:
                return _styles.Borders[Index].Top.Color;
            case eStyleClass.BorderRight:
                return _styles.Borders[Index].Right.Color;
            case eStyleClass.BorderBottom:
                return _styles.Borders[Index].Bottom.Color;
            case eStyleClass.BorderDiagonal:
                return _styles.Borders[Index].Diagonal.Color;
            default:
                throw (new Exception ("Invalid style-class for Color"));
            }
        }
        internal override void SetIndex (int index)
        {
            _parent.Index = index;
        }
        /// <summary>
        /// Return the RGB value for the Indexed or Tint property
        /// </summary>
        /// <returns>The RGB color starting with a #</returns>
        public string LookupColor ()
        {
            return LookupColor (this);
        }

        private static string NormalizeArgb (string colorValue)
        {
            if (string.IsNullOrEmpty (colorValue))
            {
                return null;
            }

            colorValue = colorValue.TrimStart ('#');
            if (colorValue.Length == 6)
            {
                colorValue = "FF" + colorValue;
            }

            return colorValue.Length == 8 ? "#" + colorValue.ToUpperInvariant () : null;
        }

        private static bool TryParseArgb (string argb, out Color color)
        {
            argb = NormalizeArgb (argb);
            if (!string.IsNullOrEmpty (argb))
            {
                uint value;
                if (uint.TryParse (argb.Substring (1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out value))
                {
                    color = Color.FromArgb (unchecked((int)value));
                    return true;
                }
            }

            color = Color.Empty;
            return false;
        }

        private static string ApplyTintToColor (string argb, decimal tint)
        {
            if (string.IsNullOrEmpty (argb) || tint == 0m)
            {
                return argb;
            }

            Color baseColor;
            if (!TryParseArgb (argb, out baseColor))
            {
                return argb;
            }

            var tinted = ApplyTint (baseColor, (double)tint);
            return "#" + tinted.ToArgb ().ToString ("X8");
        }

        private static Color ApplyTint (Color color, double tint)
        {
            if (tint < -1d)
            {
                tint = -1d;
            }
            else if (tint > 1d)
            {
                tint = 1d;
            }

            double hue;
            double saturation;
            double lightness;
            RgbToHsl (color, out hue, out saturation, out lightness);

            lightness = tint < 0d
                ? lightness * (1d + tint)
                : lightness * (1d - tint) + tint;

            if (lightness < 0d)
            {
                lightness = 0d;
            }
            else if (lightness > 1d)
            {
                lightness = 1d;
            }

            return HslToColor (hue, saturation, lightness, color.A);
        }

        private static void RgbToHsl (Color color, out double hue, out double saturation, out double lightness)
        {
            double r = color.R / 255d;
            double g = color.G / 255d;
            double b = color.B / 255d;
            double max = Math.Max (r, Math.Max (g, b));
            double min = Math.Min (r, Math.Min (g, b));

            hue = 0d;
            saturation = 0d;
            lightness = (max + min) / 2d;

            if (Math.Abs (max - min) < double.Epsilon)
            {
                return;
            }

            double delta = max - min;
            saturation = lightness > 0.5d
                ? delta / (2d - max - min)
                : delta / (max + min);

            if (Math.Abs (max - r) < double.Epsilon)
            {
                hue = (g - b) / delta + (g < b ? 6d : 0d);
            }
            else if (Math.Abs (max - g) < double.Epsilon)
            {
                hue = (b - r) / delta + 2d;
            }
            else
            {
                hue = (r - g) / delta + 4d;
            }

            hue /= 6d;
        }

        private static Color HslToColor (double hue, double saturation, double lightness, int alpha)
        {
            double r;
            double g;
            double b;

            if (Math.Abs (saturation) < double.Epsilon)
            {
                r = g = b = lightness;
            }
            else
            {
                double q = lightness < 0.5d
                    ? lightness * (1d + saturation)
                    : lightness + saturation - lightness * saturation;
                double p = 2d * lightness - q;
                r = HueToRgb (p, q, hue + (1d / 3d));
                g = HueToRgb (p, q, hue);
                b = HueToRgb (p, q, hue - (1d / 3d));
            }

            return Color.FromArgb (
                alpha,
                (int)Math.Round (r * 255d),
                (int)Math.Round (g * 255d),
                (int)Math.Round (b * 255d));
        }

        private static double HueToRgb (double p, double q, double t)
        {
            if (t < 0d)
            {
                t += 1d;
            }
            if (t > 1d)
            {
                t -= 1d;
            }
            if (t < 1d / 6d)
            {
                return p + (q - p) * 6d * t;
            }
            if (t < 1d / 2d)
            {
                return q;
            }
            if (t < 2d / 3d)
            {
                return p + (q - p) * (2d / 3d - t) * 6d;
            }

            return p;
        }

        public string LookupColor (ExcelColor theColor)
        {
            // reference extracted from ECMA-376, Part 4, Section 3.8.26 or 18.8.27 SE Part 1
            string[] rgbLookup =
            {
                "#FF000000", "#FFFFFFFF", "#FFFF0000", "#FF00FF00", "#FF0000FF", "#FFFFFF00", "#FFFF00FF", "#FF00FFFF",
                "#FF000000", "#FFFFFFFF", "#FFFF0000", "#FF00FF00", "#FF0000FF", "#FFFFFF00", "#FFFF00FF", "#FF00FFFF",
                "#FF800000", "#FF008000", "#FF000080", "#FF808000", "#FF800080", "#FF008080", "#FFC0C0C0", "#FF808080",
                "#FF9999FF", "#FF993366", "#FFFFFFCC", "#FFCCFFFF", "#FF660066", "#FFFF8080", "#FF0066CC", "#FFCCCCFF",
                "#FF000080", "#FFFF00FF", "#FFFFFF00", "#FF00FFFF", "#FF800080", "#FF800000", "#FF008080", "#FF0000FF",
                "#FF00CCFF", "#FFCCFFFF", "#FFCCFFCC", "#FFFFFF99", "#FF99CCFF", "#FFFF99CC", "#FFCC99FF", "#FFFFCC99",
                "#FF3366FF", "#FF33CCCC", "#FF99CC00", "#FFFFCC00", "#FFFF9900", "#FFFF6600", "#FF666699", "#FF969696",
                "#FF003366", "#FF339966", "#FF003300", "#FF333300", "#FF993300", "#FF993366", "#FF333399", "#FF333333",
            };

            var src = theColor.GetSource ();
            string translatedRGB = NormalizeArgb (src.Rgb);

            // 1) RGB 明示指定を最優先
            if (string.IsNullOrEmpty (translatedRGB) && !string.IsNullOrEmpty (src.Theme))
            {
                int themeIx;
                if (int.TryParse (src.Theme, out themeIx))
                {
                    translatedRGB = _styles.GetThemeColor (themeIx);
                }
            }

            // 3) indexed は明示設定時のみ使用
            if (string.IsNullOrEmpty (translatedRGB) && src.HasIndexed && src.Indexed >= 0 && src.Indexed < rgbLookup.Length)
            {
                translatedRGB = rgbLookup[src.Indexed];
            }

            // 4) 最後のフォールバック
            if (string.IsNullOrEmpty (translatedRGB))
            {
                translatedRGB = "#FF000000";
            }

            return ApplyTintToColor (translatedRGB, src.Tint);
        }
    }
}

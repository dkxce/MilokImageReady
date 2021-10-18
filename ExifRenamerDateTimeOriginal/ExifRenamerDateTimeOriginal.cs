using System;
using System.Collections.Generic;
using System.Text;

using MilokImageReady;

namespace ExifRenamerDateTimeOriginal
{
    public class ExifRenamerDateTimeOriginal : IExifRenamer
    {
        public string GetNewNameFromExifData(string[] ExifData)
        {
            return ExifData[0].Replace(":", "").Replace("\0", "");
        }

        public string GetTitle()
        {
            return "ExifRenamer DateTimeOriginal";
        }

        public string GetDescription()
        {
            return "Renames File to DateTimeOriginal Exif Value";
        }

        public override string ToString()
        {
            return GetTitle();
        }

        public string[] GetExifKeys()
        {
            return new string[] { "Exif.Photo.DateTimeOriginal" };
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;

namespace MilokImageReady
{
    public interface IExifRenamer
    {
        string[] GetExifKeys();
        string GetNewNameFromExifData(string[] ExifData);
        string GetTitle();
        string GetDescription();
    }
}

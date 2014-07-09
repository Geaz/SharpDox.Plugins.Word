using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharpDox.Plugins.Word
{
    internal static class Helper
    {
        public static string EnsureCopy(string fileToCopy, string destinationFolder)
        {
            if (!Directory.Exists(destinationFolder))
            {
                Directory.CreateDirectory(destinationFolder);
            }

            var destinationFile = Path.Combine(destinationFolder, Path.GetFileName(fileToCopy));
            File.Copy(fileToCopy, destinationFile, true);

            return destinationFile;
        }
    }
}

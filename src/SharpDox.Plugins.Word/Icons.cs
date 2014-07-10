using System.IO;

namespace SharpDox.Plugins.Word
{
    internal static class Icons
    {
        private static string _iconFolder = Path.Combine(Path.GetDirectoryName(typeof(Exporter).Assembly.Location), Templates.Folder, "icons");

        public static string GetIconPath(string type, string accessibility)
        {
            return Path.Combine(_iconFolder, string.Format("{0}_{1}.png", type.ToLower(), accessibility.ToLower()));
        }
    }
}

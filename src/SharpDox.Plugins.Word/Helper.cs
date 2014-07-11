using SharpDox.Model;
using System.IO;
using System.Linq;

namespace SharpDox.Plugins.Word
{
    internal class Helper
    {
        private readonly SDProject _sdProject;

        public Helper(SDProject sdProject)
        {
            _sdProject = sdProject;
        }

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

        public string TransformLinkToken(string linkType, string identifier)
        {
            var link = string.Empty;
            if (linkType == "image")
            {
                link = _sdProject.Images.Single(i => Path.GetFileName(i) == identifier);
            }
            else if (linkType == "namespace")
            {
                //link = string.Format("../{0}/{1}.html", linkType, identifier);
            }
            else if (linkType == "type")
            {
                //var sdType = _sdProject.GetTypeByIdentifier(identifier);
                //if (sdType != null)
                //    link = string.Format("../{0}/{1}.html", "type", sdType.ShortIdentifier);
            }
            else if(linkType == "article")
            {
                //link = string.Format("../{0}/{1}.html", linkType, identifier);
            }
            else // Member
            {
                //var sdMember = _sdProject.GetMemberByIdentifier(identifier);
                //if (sdMember != null)
                //    link = string.Format("../{0}/{1}.html#{2}", "type", sdMember.DeclaringType.ShortIdentifier, sdMember.ShortIdentifier);
            }
            
            return link;
        }
    }
}

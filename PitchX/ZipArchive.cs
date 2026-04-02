using System.IO;

namespace PitchX
{
    internal class ZipArchive
    {
        private FileStream zipToOpen;
        private object read;

        public ZipArchive(FileStream zipToOpen, object read)
        {
            this.zipToOpen = zipToOpen;
            this.read = read;
        }
    }
}
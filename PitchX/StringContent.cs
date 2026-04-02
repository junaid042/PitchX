using System.Text;

namespace PitchX
{
    internal class StringContent
    {
        private string jsonBody;
        private Encoding uTF8;
        private string v;

        public StringContent(string jsonBody, Encoding uTF8, string v)
        {
            this.jsonBody = jsonBody;
            this.uTF8 = uTF8;
            this.v = v;
        }
    }
}
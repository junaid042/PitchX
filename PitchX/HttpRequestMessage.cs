namespace PitchX
{
    internal class HttpRequestMessage
    {
        private object post;
        private string v;

        public HttpRequestMessage(object post, string v)
        {
            this.post = post;
            this.v = v;
        }

        public object Headers { get; internal set; }
    }
}
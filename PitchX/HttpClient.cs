using System;
using System.Threading.Tasks;

namespace PitchX
{
    internal class HttpClient : IDisposable
    {
        private HttpClientHandler handler;

        public HttpClient()
        {
        }

        public HttpClient(HttpClientHandler handler)
        {
            this.handler = handler;
        }

        public Uri BaseAddress { get; internal set; }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        internal void Post(string v, StringContent content)
        {
            throw new NotImplementedException();
        }

        internal object PostAsync(string apiUrl, StringContent content)
        {
            throw new NotImplementedException();
        }

        internal async Task SendAsync(HttpRequestMessage request)
        {
            throw new NotImplementedException();
        }
    }
}
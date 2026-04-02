using System.Net;

namespace PitchX
{
    internal class HttpClientHandler
    {
        public bool UseCookies { get; set; }
        public CookieContainer CookieContainer { get; set; }
    }
}
using System;
using System.Net.Http;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.DL;

namespace JBC.ExploreTheWorld.DL.CountriesNowSpaceApi
{
    /// <summary>
    /// Downloads flag images (Wikimedia PNG thumbnails referenced by the CountriesNow.space
    /// flag data) over HTTP. Works in desktop, server, and browser (WASM) hosts.
    /// </summary>
    public class FlagImageDownload__Repo : FlagImageDownload__Repo__Interface
    {
        // Wikimedia rejects requests without a User-Agent with HTTP 403
        // (https://meta.wikimedia.org/wiki/User-Agent_policy).
        private const string UserAgent = "JBC.ExploreTheWorld/1.0 (https://github.com/JohnBaluka/code-JBC-ExploreTheWorld)";

        private readonly HttpClient _client;

        public FlagImageDownload__Repo() : this(new HttpClient()) { }

        public FlagImageDownload__Repo(HttpClient client)
        {
            _client = client;
            try
            {
                if (_client.DefaultRequestHeaders.UserAgent.Count == 0)
                    _client.DefaultRequestHeaders.UserAgent.ParseAdd(UserAgent);
            }
            catch
            {
                // Browser (WASM) hosts control the User-Agent themselves — fetch strips
                // or forbids the header, and the browser's own UA satisfies Wikimedia.
            }
        }

        public Task<byte[]> DownloadImageAsync(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
                throw new ArgumentException("Image URL must not be empty.", nameof(url));

            return _client.GetByteArrayAsync(url);
        }
    }
}

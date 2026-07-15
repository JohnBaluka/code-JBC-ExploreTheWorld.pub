using System;
using System.IO;
using System.Threading.Tasks;

namespace JBC.ExploreTheWorld.DL
{
    /// <summary>
    /// File-system flag image cache used by desktop and server hosts. Images are stored as
    /// <c>{iso2}.png</c> under <c>%LocalAppData%\JBC\ExploreTheWorld\FlagImages</c> (or a
    /// custom root passed to the constructor), so they survive restarts and work offline.
    /// </summary>
    public class FlagImageStore_FileSystem__Repo : FlagImageStore__Repo__Interface
    {
        private readonly string _rootPath;

        public FlagImageStore_FileSystem__Repo() : this(GetDefaultRootPath()) { }

        public FlagImageStore_FileSystem__Repo(string rootPath)
        {
            _rootPath = rootPath;
        }

        public static string GetDefaultRootPath() =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "JBC", "ExploreTheWorld", "FlagImages");

        public Task<byte[]?> GetImageAsync(string iso2)
        {
            var path = GetPath(iso2);
            return Task.FromResult<byte[]?>(File.Exists(path) ? File.ReadAllBytes(path) : null);
        }

        public Task SaveImageAsync(string iso2, byte[] imageBytes)
        {
            Directory.CreateDirectory(_rootPath);
            File.WriteAllBytes(GetPath(iso2), imageBytes);
            return Task.CompletedTask;
        }

        public Task<string?> GetImageFilePathAsync(string iso2)
        {
            var path = GetPath(iso2);
            return Task.FromResult<string?>(File.Exists(path) ? path : null);
        }

        public Task ClearAsync()
        {
            if (Directory.Exists(_rootPath))
                foreach (var file in Directory.GetFiles(_rootPath, "*.png"))
                    File.Delete(file);
            return Task.CompletedTask;
        }

        private string GetPath(string iso2)
        {
            // ISO2 codes are two letters; sanitize defensively so a malformed code can
            // never escape the cache directory.
            var safe = iso2.Trim().ToUpperInvariant();
            foreach (var c in Path.GetInvalidFileNameChars())
                safe = safe.Replace(c.ToString(), string.Empty);
            return Path.Combine(_rootPath, safe + ".png");
        }
    }
}

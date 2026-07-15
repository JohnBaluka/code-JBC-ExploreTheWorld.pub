using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.CL;
using JBC.ExploreTheWorld.DL;

namespace JBC.ExploreTheWorld.BL
{
    /// <summary>
    /// Resolves country flag images with a cache-first strategy: check the local image
    /// store (file system or browser IndexedDB) → derive the Wikimedia PNG thumbnail URL
    /// from the CountriesNow flag data → download → persist → return.
    /// </summary>
    public class FlagImageManager__Service
    {
        private readonly FlagImageStore__Repo__Interface _store;
        private readonly FlagImageDownload__Repo__Interface _download;
        private readonly CountriesNowSpaceManager__Service _countriesNowManager;

        private Dictionary<string, string>? _flagUrlsByIso2;
        private readonly SemaphoreSlim _flagUrlsLock = new(1, 1);

        public FlagImageManager__Service(
            FlagImageStore__Repo__Interface store,
            FlagImageDownload__Repo__Interface download,
            CountriesNowSpaceManager__Service countriesNowManager)
        {
            _store               = store;
            _download            = download;
            _countriesNowManager = countriesNowManager;
        }

        public async Task<FlagImage_Row?> GetFlagImageAsync(string iso2)
        {
            if (string.IsNullOrWhiteSpace(iso2))
                return null;

            var key = iso2.Trim().ToUpperInvariant();

            var cached = await _store.GetImageAsync(key);
            if (cached is { Length: > 0 })
            {
                return new FlagImage_Row
                {
                    Iso2       = key,
                    ImageBytes = cached,
                    FilePath   = await _store.GetImageFilePathAsync(key),
                    Source     = DataSource_Enum.Database
                };
            }

            var pngUrl = FlagImageUrl_Helper.GetPngThumbnailUrl(await GetFlagUrlAsync(key));
            if (pngUrl == null)
                return null;

            var bytes = await _download.DownloadImageAsync(pngUrl);
            if (bytes.Length == 0)
                return null;

            await _store.SaveImageAsync(key, bytes);

            return new FlagImage_Row
            {
                Iso2       = key,
                ImageBytes = bytes,
                FilePath   = await _store.GetImageFilePathAsync(key),
                Source     = DataSource_Enum.Api
            };
        }

        public async Task<IDictionary<string, FlagImage_Row>> GetFlagImagesAsync(
            IEnumerable<string> iso2Codes, Action<string>? log = null)
        {
            var result    = new Dictionary<string, FlagImage_Row>(StringComparer.OrdinalIgnoreCase);
            var processed = 0;
            var failed    = 0;

            foreach (var iso2 in iso2Codes)
            {
                if (string.IsNullOrWhiteSpace(iso2) || result.ContainsKey(iso2.Trim()))
                    continue;

                try
                {
                    var image = await GetFlagImageAsync(iso2);
                    if (image != null)
                        result[image.Iso2] = image;
                    else
                        failed++;
                }
                catch
                {
                    // A single unreachable image must not fail the whole batch (offline, 404, …).
                    failed++;
                }

                processed++;
                if (processed % 50 == 0)
                    log?.Invoke($"  Resolved {processed} flag images ({result.Count} ok, {failed} unavailable)...");
            }

            log?.Invoke($"  Flag images resolved: {result.Count} ok, {failed} unavailable.");
            return result;
        }

        public Task ClearCacheAsync() => _store.ClearAsync();

        // ── Flag URL lookup (ISO2 → Wikimedia SVG URL, loaded once per manager) ──────

        private async Task<string?> GetFlagUrlAsync(string iso2)
        {
            if (_flagUrlsByIso2 == null)
            {
                await _flagUrlsLock.WaitAsync();
                try
                {
                    if (_flagUrlsByIso2 == null)
                    {
                        var flags = await _countriesNowManager.GetCountryFlagsAsync();
                        var map   = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var flag in flags.Data)
                        {
                            if (!string.IsNullOrEmpty(flag.Iso2) && !string.IsNullOrEmpty(flag.Flag))
                                map[flag.Iso2!] = flag.Flag!;
                        }
                        _flagUrlsByIso2 = map;
                    }
                }
                finally
                {
                    _flagUrlsLock.Release();
                }
            }

            return _flagUrlsByIso2.TryGetValue(iso2, out var url) ? url : null;
        }
    }
}

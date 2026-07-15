using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Primitives;

namespace JBC.ExploreTheWorld.AL.MsOfficeWordVstoAddIn;

// In COM-host (VSTO) context Assembly.GetEntryAssembly() returns null, so the built-in
// StaticWebAssetsLoader silently exits without registering an IFileProvider.
// BlazorWebView then falls back to PhysicalFileProvider(contentRootDir), which cannot
// serve _framework/blazor.webview.js because that file lives in the NuGet package cache,
// not in the local wwwroot/ directory.
//
// This class reads the *.staticwebassets.runtime.json manifest that MSBuild writes next
// to the output assembly and routes each virtual path to the physical file the manifest
// specifies, exactly replicating what StaticWebAssetsLoader does in a normal exe context.
internal sealed class RuntimeManifestFileProvider : IFileProvider
{
    private readonly Dictionary<string, string> _map =
        new(StringComparer.OrdinalIgnoreCase);

    public RuntimeManifestFileProvider(string manifestPath)
    {
        using var doc = JsonDocument.Parse(File.ReadAllBytes(manifestPath));
        var root = doc.RootElement;

        var contentRoots = new List<string>();
        foreach (var cr in root.GetProperty("ContentRoots").EnumerateArray())
            contentRoots.Add(cr.GetString()!);

        WalkTree(root.GetProperty("Root"), string.Empty, contentRoots);
    }

    private void WalkTree(JsonElement node, string prefix, List<string> contentRoots)
    {
        if (!node.TryGetProperty("Children", out var children)
            || children.ValueKind == JsonValueKind.Null)
            return;

        foreach (var kv in children.EnumerateObject())
        {
            var virtualPath = prefix.Length == 0 ? kv.Name : $"{prefix}/{kv.Name}";

            if (kv.Value.TryGetProperty("Asset", out var asset)
                && asset.ValueKind != JsonValueKind.Null)
            {
                var idx     = asset.GetProperty("ContentRootIndex").GetInt32();
                var subPath = asset.GetProperty("SubPath").GetString()!
                                   .Replace('/', Path.DirectorySeparatorChar);
                _map[virtualPath] = Path.Combine(contentRoots[idx], subPath);
            }

            WalkTree(kv.Value, virtualPath, contentRoots);
        }
    }

    public IFileInfo GetFileInfo(string subpath)
    {
        var key = subpath.TrimStart('/').Replace('\\', '/');
        if (_map.TryGetValue(key, out var physical) && File.Exists(physical))
            return new PhysFile(physical);
        return new NotFoundFileInfo(subpath);
    }

    public IDirectoryContents GetDirectoryContents(string subpath)
        => NotFoundDirectoryContents.Singleton;

    public IChangeToken Watch(string filter) => NullChangeToken.Singleton;

    private sealed class PhysFile : IFileInfo
    {
        private readonly FileInfo _fi;
        public PhysFile(string path) => _fi = new FileInfo(path);
        public bool           Exists       => _fi.Exists;
        public long           Length       => _fi.Length;
        public string?        PhysicalPath => _fi.FullName;
        public string         Name         => _fi.Name;
        public DateTimeOffset LastModified => _fi.LastWriteTimeUtc;
        public bool           IsDirectory  => false;
        public Stream         CreateReadStream() => _fi.OpenRead();
    }
}

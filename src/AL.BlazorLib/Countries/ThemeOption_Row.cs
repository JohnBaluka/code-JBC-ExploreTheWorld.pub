using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using JBC.ExploreTheWorld.BL;
using JBC.ExploreTheWorld.DL.CountriesNowSpaceApi;
using JBC.ExploreTheWorld.DL.MsJSInterop.RevealJs;
using Microsoft.AspNetCore.Components;
using Microsoft.JSInterop;

namespace JBC.ExploreTheWorld.AL.BlazorLib.Countries
{
    // ── Display row types ─────────────────────────────────────────────────────
    public record ThemeOption_Row(string Label, string Value);
}

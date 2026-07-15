using System;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice
{
    // Canonical markers for properties a given API (NetOffice, OpenXML, Office.js) cannot provide.
    // Strings use the "**Undefined" marker, enums and plain numerics use -99, booleans and dates use null.
    public static class MsOfficeUndefined
    {
        public const string String = "**Undefined";
        public const int Number = -99;
        public const long NumberLong = -99;

        public static bool IsUndefined(string? value)
        {
            return value == String;
        }

        public static bool IsUndefined(int? value)
        {
            return value == Number;
        }

        public static bool IsUndefined(long? value)
        {
            return value == NumberLong;
        }
    }
}

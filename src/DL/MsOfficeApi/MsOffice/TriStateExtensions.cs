using System;
using System.Collections.Generic;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsOffice
{
    public static class TriStateExtensions
    {

        public static bool IsTrue(this int value)
        {
            TriState_Enum enumValue = (TriState_Enum)value;
            return IsTrue(enumValue);
        }

        public static bool IsTrue(this TriState_Enum value)
        {
            switch (value)
            {
                case TriState_Enum.Undefined:
                case TriState_Enum.msoTriStateToggle:
                case TriState_Enum.msoTriStateMixed:
                case TriState_Enum.msoFalse:
                    return false;
                case TriState_Enum.msoTrue:
                case TriState_Enum.msoCTrue:
                    return true;
                default:
                    throw new Exception("TriStateEnum value not found.");
            }
        }
    }
}

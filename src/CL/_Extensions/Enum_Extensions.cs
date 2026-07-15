using System;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace JBC.ExploreTheWorld.CL
{
    public static class Enum_Extensions
    {

        public static string GetEnumDisplayName(System.Enum en)
        {
            Type type = en.GetType();
            MemberInfo[] memInfo = type.GetMember(en.ToString());
            if (memInfo != null && memInfo.Length > 0)
            {
                object[] attrs = memInfo[0].GetCustomAttributes(typeof(DisplayNameAttribute), false);
                if (attrs != null && attrs.Length > 0)
                    return ((DisplayNameAttribute)attrs[0]).DisplayName;
            }
            return en.ToString();
        }

        public static string GetEnumDescription(System.Enum en)
        {
            Type type = en.GetType();
            MemberInfo[] memInfo = type.GetMember(en.ToString());
            if (memInfo != null && memInfo.Length > 0)
            {
                object[] attrs = memInfo[0].GetCustomAttributes(typeof(DescriptionAttribute), false);
                if (attrs != null && attrs.Length > 0)
                    return ((DescriptionAttribute)attrs[0]).Description;
            }
            return en.ToString();
        }

        public static Guid GetEnumGUID(System.Enum en)
        {
            Type type = en.GetType();
            MemberInfo[] memInfo = type.GetMember(en.ToString());
            if (memInfo != null && memInfo.Length > 0)
            {
                object[] attrs = memInfo[0].GetCustomAttributes(typeof(GuidAttribute), false);
                if (attrs != null && attrs.Length > 0)
                    return new Guid(((GuidAttribute)attrs[0]).Value);
            }
            return Guid.Empty;
        }

        private static Regex UpperCamelCaseRegex = new Regex(@"(?<!^)((?<!\d)\d|(?(?<=[A-Z])[A-Z](?=[a-z])|[A-Z]))", RegexOptions.Compiled);
        public static string AsUpperCamelCaseName(this Enum en)
        {
            return UpperCamelCaseRegex.Replace(en.ToString(), " $1");
        }

        public static List<T> EnumToList<T>()
        {
            Type enumType = typeof(T);

            // Can't use type constraints on value types, so have to do check like this
            if (enumType.BaseType != typeof(Enum))
                throw new ArgumentException("T must be of type System.Enum");

            Array enumValArray = Enum.GetValues(enumType);

            List<T> enumValList = new List<T>(enumValArray.Length);

            foreach (int val in enumValArray)
            {
                enumValList.Add((T)Enum.Parse(enumType, val.ToString()));
            }

            return enumValList;
        }

        public static bool IsSet(this Enum en, Enum matchTo)
        {
            return (Convert.ToUInt32(en) & Convert.ToUInt32(matchTo)) != 0;
        }
    }
}

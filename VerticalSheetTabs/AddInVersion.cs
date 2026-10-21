using System;
using System.Reflection;

namespace VerticalSheetTabs
{
    // The version of this add-in as it is written in the release tag: "v1.1.19". It is the assembly
    // version (AssemblyInfo.cs, the same number as ApplicationVersion in the project and the folder
    // of the installer), which is set when the installer is built, before the tag is given.
    public static class AddInVersion
    {
        public const string Copyright = "Copyright (c) 2026 nk_taka";

        public const string License = "BSD ライセンスに準拠";

        public static string Text
        {
            get
            {
                try
                {
                    Version v = Assembly.GetExecutingAssembly().GetName().Version;
                    return "v" + v.Major + "." + v.Minor + "." + v.Build;
                }
                catch (Exception)
                {
                    return "";
                }
            }
        }
    }
}

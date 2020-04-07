using System;

namespace Подбор_кандидатов
{
    public static class Pathes
    {
        public static string DBPAth = Environment.CurrentDirectory + "\\DBScientificCompany.dbsc";
        public static string DBPAthSostav = Environment.CurrentDirectory + "\\DBScientificCompanySostav.dbsc";
        public static string TempPath = Environment.CurrentDirectory + "\\Temp";

        private static string Directory = "Shablons";
        public static string CurrentDirectory = Environment.CurrentDirectory + "\\" + Directory;
        private static string CreateShablonPath(string FileName) => CurrentDirectory + "\\" + FileName;
        public static string CreateTempPath(string FileName) => TempPath + "\\" + FileName + ".docx";
        
        public static string VedomostShablon = CreateShablonPath("VedomostShablon.docx");
        public static string RaitingShablon = CreateShablonPath("RaitingShablon.docx");
        public static string SpisokShablon = CreateShablonPath("SpisokShablon.docx");
        public static string ListShablon = CreateShablonPath("ListShablon.docx");
    }
}

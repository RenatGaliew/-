using System;

namespace Подбор_кандидатов__выгрузка_данных
{
    public static class Pathes
    {
        public static string TempPath = Environment.CurrentDirectory + "\\Temp";

        private static string Directory = "Shablons";
        public static string DirectoryPDFS = "ToSendPDFS/";
        public static string DirectoryGenerateList = $"{DirectoryPDFS}/GenerateList";
        public static string DownloadsPath = "downloads/";
        public static string DownloadsPathPhotos = "downloads/photos";
        public static string CurrentDirectory = Environment.CurrentDirectory + "\\" + Directory;
        private static string CreateShablonPath(string FileName) => CurrentDirectory + "\\" + FileName;
        public static string CreateTempPath(string FileName) => TempPath + "\\" + FileName + ".docx";
        
        public static string VedomostShablon = CreateShablonPath("VedomostShablon.docx");
        public static string RaitingShablon = CreateShablonPath("RaitingShablon.docx");
        public static string SpisokShablon = CreateShablonPath("SpisokShablon.docx");
        public static string ListShablon = CreateShablonPath("ListShablon.docx");
    }
}

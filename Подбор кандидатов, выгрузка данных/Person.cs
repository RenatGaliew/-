using System;
using System.IO;
using System.Linq;
using System.Text;

namespace Подбор_кандидатов__выгрузка_данных
{
    public class Person
    {
        public event EventHandler IsSendingChanged;
        public event EventHandler IsFilesDownloaded;
        private double[] Koef = { 0.25, 0.15, 0.3, 0.2, 0.5, 0.25, 0.1 };
        public string PDFPath { get; set; }
        public string email { get; set; }
        public Guid ID { get; set; }
        public string URLPhoto { get; set; }
        public string URLSOGLASIE { get; set; }
        public string SurName { get; set; }
        public string Name { get; set; }
        public string Patronomyc { get; set; }
        public string Telefon { get; set; }
        public DateTime Birthday { get; set; }
        public string BirthdayMesto { get; set; }
        public string Country { get; set; }
        public string Addres { get; set; }
        public string AddresRegistry { get; set; }
        public string VK { get; set; }
        public string Health { get; set; }
        public string VUZ { get; set; }
        public string VUZKor { get; set; }
        public string Specialnost { get; set; }
        public string Diplom { get; set; }
        public string VKR { get; set; }
        public double[] Statiy { get; set; }
        public string Statya { get; set; }
        public double[] Prioritet { get; set; }
        public double[] Sience { get; set; }
        public string SienceName { get; set; }
        public double[] SienceStepen { get; set; }
        public double[] Work { get; set; }
        public string WorkName { get; set; }
        public double[] Sport { get; set; }
        public string Napravlenie { get; set; }
        public string Dopusk { get; set; }
        public string Info { get; set; }
        public double SrBall { get; set; }

        public string Family { get; set; }
        public string TattooAndPirsing { get; set; }
        public string HealtHron { get; set; }
        public string SubjectRF { get; set; }

        public double Ball { get; set; }
        public string SportName { get; set; }
        public string Language { get; set; }
        public double Rost { get; set; }
        public string Products { get; set; }
        public double Ves { get; set; }
        public int Soiskatelstvo { get; set; }
        public string Exams { get; set; }
        public bool Reserv { get; set; }
        public bool Magistr { get; set; }
        public bool Bakalavr { get; set; }
        public bool Zachetka { get; set; }
        private bool _isEmailSending;
        private bool _isAllFilesDownloaded;

        public bool IsEmailSending
        {
            get => _isEmailSending;
            set
            {
                _isEmailSending = value;
                IsSendingChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        public string Status => IsEmailSending ? "Отправлено" : "Не отправлено";
        public string Child { get; set; }

        public string SoglasieURLPath { get; set; }
        public string ListSobesURLPath { get; set; }
        public string DiplomURLPath { get; set; }
        public string[] StatyaURLPath { get; set; }
        public string[] OlympURLPath { get; set; }
        public string[] KandidatURLPath { get; set; }
        public string[] WorkURLPath { get; set; }
        public string[] SportURLPath { get; set; }
        public bool IsEtap2 { get; set; }

        public bool IsAllFilesDownloaded
        {
            get => _isAllFilesDownloaded;
            set
            {
                _isAllFilesDownloaded = value;
                IsFilesDownloaded?.Invoke(this, EventArgs.Empty);
            }
        }

        public string PhotoPath { get; set; }

        public string SoglasiePath { get; set; }
        public string ListSobesPath { get; set; }
        public string DiplomPath { get; set; }
        public string[] StatyaPath { get; set; }
        public string[] OlympPath { get; set; }
        public string[] KandidatPath { get; set; }
        public string[] WorkPath { get; set; }
        public string[] SportPath { get; set; }

        public override string ToString()
        {
            string  str = $"{SurName} {Name} {Patronomyc}";
            return str;
        }

        public string ToUpperString()
        {
            return $"{SurName.ToUpper()}{Environment.NewLine}{Name}{Environment.NewLine}{Patronomyc}";
        }

        public Person()
        {
            Statiy = new double[6];
            Sience = new double[7];
            SienceStepen = new double[3];
            Work = new double[3];
            Sport = new double[2];
            Prioritet = new double[2];

            StatyaURLPath = new string[6];
            OlympURLPath = new string[7];
            KandidatURLPath = new string[3];
            WorkURLPath = new string[3];
            SportURLPath = new string[2];
            ID = Guid.NewGuid();
            IsAllFILESDownloaded();
        }

        public double ExecuteBall()
        {
            double summa = 0;
            foreach (var k in Statiy)
            {
                summa += Koef[0] * k;
            }

            summa += SrBall * Koef[1] * (Bakalavr ? 0.8 : Zachetka ? 0.7 : Magistr ? 1 : 1);

            foreach (var k in Prioritet)
            {
                summa += Koef[2] * k;
            }
            foreach (var k in Sience)
            {
                summa += Koef[3] * k;
            }
            foreach (var k in SienceStepen)
            {
                summa += Koef[4] * k;
            }
            foreach (var k in Work)
            {
                summa += Koef[5] * k;
            }
            foreach (var k in Sport)
            {
                summa += Koef[6] * k;
            }
            return summa;
        }

        public string ShortString()
        {
            return $"{SurName} {Name[0]}. {Patronomyc[0]}.";
        }
        public string FileNameDOCX()
        {
            return $"{SurName}_{Name}_{Patronomyc}.docx";
        }

        public string FileNamePDF()
        {
            return $"{SurName}_{Name}_{Patronomyc}.pdf";
        }

        public string ShortStringNewLine()
        {
            return $"{SurName} {Name[0]}. {Patronomyc[0]}.";
        }

        public string LongString()
        {
            return $"{SurName} {Name} {Patronomyc}";
        }

        public bool OK()
        {
            return SurName != "" &&
                   Name != "" &&
                   Patronomyc != "" &&
                   Telefon != "" &&
                   Birthday != null &&
                   BirthdayMesto != "" &&
                   Country != "" &&
                   Addres != "" &&
                   AddresRegistry != "" &&
                   VK != "" &&
                   Health != "" &&
                   VUZ != "" &&
                   VUZKor != "" &&
                   Specialnost != "" &&
                   Diplom != "" &&
                   VKR != "" &&
                   Prioritet.Sum() != 0 &&
                   Statya != "" &&
                   SienceName != "" &&
                   WorkName != "" &&
                   Napravlenie != "" &&
                   Dopusk != "" &&
                   SrBall != 0 &&
                   SportName != "" &&
                   Language != "" &&
                   Rost != 0 &&
                   Products != "" &&
                   Ves != 0 &&
                   Family != "" &&
                   TattooAndPirsing != "" &&
                   HealtHron != "" &&
                   SubjectRF != "";
        }

        public void IsAllFILESDownloaded()
        {
            var tmp = true;
            if (!string.IsNullOrEmpty(SoglasiePath))
            {
                if (!File.Exists(SoglasiePath))
                    tmp = false;
            }
            if (tmp && !string.IsNullOrEmpty(ListSobesPath))
            {
                if (!File.Exists(ListSobesPath))
                    tmp = false;
            }
            if (tmp && !string.IsNullOrEmpty(DiplomPath))
            {
                if (!File.Exists(DiplomPath))
                    tmp = false;
            }

            foreach (var s in StatyaPath)
            {
                if (!tmp || string.IsNullOrEmpty(s)) continue;
                if (!File.Exists(s))
                    tmp = false;
            }
            foreach (var s in OlympPath)
            {
                if (!tmp || string.IsNullOrEmpty(s)) continue;
                if (!File.Exists(s))
                    tmp = false;
            }
            foreach (var s in KandidatPath)
            {
                if (!tmp || string.IsNullOrEmpty(s)) continue;
                if (!File.Exists(s))
                    tmp = false;
            }
            foreach (var s in WorkPath)
            {
                if (!tmp || string.IsNullOrEmpty(s)) continue;
                if (!File.Exists(s))
                    tmp = false;
            }
            foreach (var s in SportPath)
            {
                if (!tmp || string.IsNullOrEmpty(s)) continue;
                if (!File.Exists(s))
                    tmp = false;
            }

            IsAllFilesDownloaded = tmp;
        }
    }

    public class FileToDownload
    {
        public Google.Apis.Drive.v2.Data.File File { get; set; }

        public FileToDownload(Google.Apis.Drive.v2.Data.File file)
        {
            File = file;
        }
        public override string ToString()
        {
            return $"{File.Title}({File.ModifiedDate})";
        }
    }
}            
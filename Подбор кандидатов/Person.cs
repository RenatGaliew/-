using System;
using System.Linq;
using Newtonsoft.Json.Serialization;

namespace Подбор_кандидатов
{
    public class Person
    {
        public double[] Koef = { 0.25, 0.15, 0.3, 0.2, 0.5, 0.25, 0.1 };
        public string base64SoglasieString { get; set; }
        public string email { get; set; }
        public string photoPath { get; set; }
        public string base64PhotoString { get; set; }
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

        public override string ToString()
        {
            return $"{SurName} {Name} {Patronomyc}";
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
            ID = Guid.NewGuid();
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

        public string ShortStringNewLine()
        {
            return $"{SurName} {Name[0]}. {Patronomyc[0]}.";
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
    }
}

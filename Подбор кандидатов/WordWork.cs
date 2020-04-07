using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace Подбор_кандидатов
{
    public class WordWork
    {
        private Word.Application wordapp;
        private Word.Documents worddocuments;
        private Word.Document worddocument;

        private string FileName { get; set; }

        private void InitializeShablon(string fileName, string fileNameTMP)
        {
            File.Copy(fileName, fileNameTMP, true);
            wordapp = new Word.Application();
            wordapp.Visible = false;
            Object filename = fileNameTMP;
            Object confirmConversions = true;
            Object readOnly = false;
            Object addToRecentFiles = false;
            Object passwordDocument = Type.Missing;
            Object passwordTemplate = Type.Missing;
            Object revert = false;
            Object writePasswordDocument = Type.Missing;
            Object writePasswordTemplate = Type.Missing;
            Object format = Type.Missing;
            Object encoding = Type.Missing;
            Object oVisible = Type.Missing;
            Object openAndRepair = Type.Missing;
            Object documentDirection = Type.Missing;
            Object noEncodingDialog = false;
            Object xmlTransform = Type.Missing;
            worddocument = wordapp.Documents.Open(ref filename,
                ref confirmConversions, ref readOnly, ref addToRecentFiles,
                ref passwordDocument, ref passwordTemplate, ref revert,
                ref writePasswordDocument, ref writePasswordTemplate,
                ref format, ref encoding, ref oVisible,
                ref openAndRepair, ref documentDirection, ref noEncodingDialog, ref xmlTransform);
        }

        private void SaveAndQuit()
        {
            worddocument.Save();
            wordapp.Quit();
        }

        private void InitializeShablonToShow(string fileName)
        {
            wordapp = new Word.Application();
            wordapp.Visible = true;
            Object filename = fileName;
            Object confirmConversions = true;
            Object readOnly = false;
            Object addToRecentFiles = false;
            Object passwordDocument = Type.Missing;
            Object passwordTemplate = Type.Missing;
            Object revert = false;
            Object writePasswordDocument = Type.Missing;
            Object writePasswordTemplate = Type.Missing;
            Object format = Type.Missing;
            Object encoding = Type.Missing;
            Object oVisible = Type.Missing;
            Object openAndRepair = Type.Missing;
            Object documentDirection = Type.Missing;
            Object noEncodingDialog = false;
            Object xmlTransform = Type.Missing;
            worddocument = wordapp.Documents.Open(ref filename,
                ref confirmConversions, ref readOnly, ref addToRecentFiles,
                ref passwordDocument, ref passwordTemplate, ref revert,
                ref writePasswordDocument, ref writePasswordTemplate,
                ref format, ref encoding, ref oVisible,
                ref openAndRepair, ref documentDirection, ref noEncodingDialog, ref xmlTransform);
        }

        public void VedomostStart(List<Person> persons)
        {
            FileName = Pathes.CreateTempPath(Guid.NewGuid().ToString());
            InitializeShablon(Pathes.VedomostShablon, FileName);

            #region Заполняет таблицу

            Word.Table wordtable = worddocument.Tables[1];

            int rowsCount = wordtable.Rows.Count;
            if (rowsCount == 2)
            {
                //добавление новых строк в таблицу
                var wordcellrange = wordtable.Rows.Last.Cells[1].Range;
                wordcellrange.Text = "1.";
                wordcellrange = wordtable.Rows.Last.Cells[2].Range;
                wordcellrange.Text = persons[0].ToString();

                wordcellrange = wordtable.Rows.Last.Cells[3].Range;
                wordcellrange.Text = persons[0].Statiy.Sum().ToString();

                wordcellrange = wordtable.Rows.Last.Cells[4].Range;
                wordcellrange.Text = persons[0].Koef[0].ToString();

                wordcellrange = wordtable.Rows.Last.Cells[5].Range;
                wordcellrange.Text = persons[0].SrBall.ToString();

                wordcellrange = wordtable.Rows.Last.Cells[6].Range;
                wordcellrange.Text = persons[0].Koef[1].ToString();

                wordcellrange = wordtable.Rows.Last.Cells[7].Range;
                wordcellrange.Text = persons[0].Prioritet.Sum().ToString();

                wordcellrange = wordtable.Rows.Last.Cells[8].Range;
                wordcellrange.Text = persons[0].Koef[2].ToString();

                wordcellrange = wordtable.Rows.Last.Cells[9].Range;
                wordcellrange.Text = persons[0].Sience.Sum().ToString();

                wordcellrange = wordtable.Rows.Last.Cells[10].Range;
                wordcellrange.Text = persons[0].Koef[3].ToString();

                wordcellrange = wordtable.Rows.Last.Cells[11].Range;
                wordcellrange.Text = persons[0].SienceStepen.Sum().ToString();

                wordcellrange = wordtable.Rows.Last.Cells[12].Range;
                wordcellrange.Text = persons[0].Koef[4].ToString();

                wordcellrange = wordtable.Rows.Last.Cells[13].Range;
                wordcellrange.Text = persons[0].Work.Sum().ToString();

                wordcellrange = wordtable.Rows.Last.Cells[14].Range;
                wordcellrange.Text = persons[0].Koef[5].ToString();

                wordcellrange = wordtable.Rows.Last.Cells[15].Range;
                wordcellrange.Text = persons[0].Sport.Sum().ToString();

                wordcellrange = wordtable.Rows.Last.Cells[16].Range;
                wordcellrange.Text = persons[0].Koef[6].ToString();

                wordcellrange = wordtable.Rows.Last.Cells[17].Range;
                wordcellrange.Text = persons[0].Ball.ToString();

                for (int i = 1; i < persons.Count; i++)
                {
                    wordtable.Rows.Add();
                    wordcellrange = wordtable.Rows.Last.Cells[1].Range;
                    wordcellrange.Text = i + 1 + ".";
                    wordcellrange = wordtable.Rows.Last.Cells[2].Range;
                    wordcellrange.Text = persons[i].ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[3].Range;
                    wordcellrange.Text = persons[i].Statiy.Sum().ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[4].Range;
                    wordcellrange.Text = persons[i].Koef[0].ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[5].Range;
                    wordcellrange.Text = persons[i].SrBall.ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[6].Range;
                    wordcellrange.Text = persons[i].Koef[1].ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[7].Range;
                    wordcellrange.Text = persons[i].Prioritet.Sum().ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[8].Range;
                    wordcellrange.Text = persons[i].Koef[2].ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[9].Range;
                    wordcellrange.Text = persons[i].Sience.Sum().ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[10].Range;
                    wordcellrange.Text = persons[i].Koef[3].ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[11].Range;
                    wordcellrange.Text = persons[i].SienceStepen.Sum().ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[12].Range;
                    wordcellrange.Text = persons[i].Koef[4].ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[13].Range;
                    wordcellrange.Text = persons[i].Work.Sum().ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[14].Range;
                    wordcellrange.Text = persons[i].Koef[5].ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[15].Range;
                    wordcellrange.Text = persons[i].Sport.Sum().ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[16].Range;
                    wordcellrange.Text = persons[i].Koef[6].ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[17].Range;
                    wordcellrange.Text = persons[i].Ball.ToString();
                }

                ///////////////////////////////
            }
            #endregion
        }

        public void RaitingStart(List<Person> persons, int max)
        {
            FileName = Pathes.CreateTempPath(Guid.NewGuid().ToString());
            InitializeShablon(Pathes.RaitingShablon, FileName);

            #region Заполняет таблицу

            Word.Table wordtable = worddocument.Tables[1];

            int rowsCount = wordtable.Rows.Count;
            if (rowsCount == 2)
            {
                //добавление новых строк в таблицу
                var wordcellrange = wordtable.Rows.Last.Cells[1].Range;
                wordcellrange.Text = "1.";
                wordcellrange = wordtable.Rows.Last.Cells[2].Range;
                wordcellrange.Text = persons[0].ToUpperString();// + Environment.NewLine + (persons[0].Reserv ? "(Резерв)" : "");

                wordcellrange = wordtable.Rows.Last.Cells[3].Range;
                wordcellrange.Text = persons[0].Birthday.ToShortDateString();

                wordcellrange = wordtable.Rows.Last.Cells[4].Range;
                wordcellrange.Text = persons[0].VK;

                wordcellrange = wordtable.Rows.Last.Cells[5].Range;
                wordcellrange.Text = persons[0].VUZ;

                wordcellrange = wordtable.Rows.Last.Cells[6].Range;
                wordcellrange.Text = persons[0].Specialnost;

                wordcellrange = wordtable.Rows.Last.Cells[7].Range;
                wordcellrange.Text = persons[0].Napravlenie;

                wordcellrange = wordtable.Rows.Last.Cells[8].Range;
                wordcellrange.Text = persons[0].SrBall.ToString();

                wordcellrange = wordtable.Rows.Last.Cells[9].Range;
                wordcellrange.Text = persons[0].Ball.ToString();

                for (int i = 1; i < persons.Count; i++)
                {
                    if (i == max-1)
                    {
                        wordtable.Rows.Add();
                        object begCell = wordtable.Cell(max + 1, 1).Range.Start;
                        object endCell = wordtable.Cell(max + 1, 9).Range.End;
                        wordcellrange = worddocument.Range(ref begCell, ref endCell);
                        wordcellrange.Select();
                        wordapp.Selection.Cells.Merge();
                    }
                    wordtable.Rows.Add();
                    wordcellrange = wordtable.Rows.Last.Cells[1].Range;
                    wordcellrange.Text = i + 1 + ".";
                    wordcellrange = wordtable.Rows.Last.Cells[2].Range;
                    wordcellrange.Text = persons[i].ToUpperString();// + Environment.NewLine + (persons[i].Reserv ? "(Резерв)" : "");

                    wordcellrange = wordtable.Rows.Last.Cells[3].Range;
                    wordcellrange.Text = persons[i].Birthday.ToShortDateString();

                    wordcellrange = wordtable.Rows.Last.Cells[4].Range;
                    wordcellrange.Text = persons[i].VK;

                    wordcellrange = wordtable.Rows.Last.Cells[5].Range;
                    wordcellrange.Text = persons[i].VUZ;

                    wordcellrange = wordtable.Rows.Last.Cells[6].Range;
                    wordcellrange.Text = persons[i].Specialnost;

                    wordcellrange = wordtable.Rows.Last.Cells[7].Range;
                    wordcellrange.Text = persons[i].Napravlenie;

                    wordcellrange = wordtable.Rows.Last.Cells[8].Range;
                    wordcellrange.Text = persons[i].SrBall.ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[9].Range;
                    wordcellrange.Text = persons[i].Ball.ToString();

                }

                ///////////////////////////////
            }
            #endregion
            
        }

        public void SpisokStart(List<Person> persons)
        {
            FileName = Pathes.CreateTempPath(Guid.NewGuid().ToString());
            InitializeShablon(Pathes.SpisokShablon, FileName);

            #region Заполняет таблицу

            Word.Table wordtable = worddocument.Tables[1];

            int rowsCount = wordtable.Rows.Count;
            if (rowsCount == 2)
            {
                //добавление новых строк в таблицу
                var wordcellrange = wordtable.Rows.Last.Cells[1].Range;
                wordcellrange.Text = "1.";
                wordcellrange = wordtable.Rows.Last.Cells[2].Range;
                wordcellrange.Text = persons[0].ShortString();

                wordcellrange = wordtable.Rows.Last.Cells[3].Range;
                wordcellrange.Text = persons[0].Birthday.Year.ToString();

                wordcellrange = wordtable.Rows.Last.Cells[4].Range;
                wordcellrange.Text = persons[0].SubjectRF;

                wordcellrange = wordtable.Rows.Last.Cells[5].Range;
                wordcellrange.Text = persons[0].SrBall.ToString();

                wordcellrange = wordtable.Rows.Last.Cells[6].Range;
                wordcellrange.Text = persons[0].Ball.ToString();

                for (int i = 1; i < persons.Count; i++)
                {
                    wordtable.Rows.Add();
                    wordcellrange = wordtable.Rows.Last.Cells[1].Range;
                    wordcellrange.Text = i + 1 + ".";
                    wordcellrange = wordtable.Rows.Last.Cells[2].Range;
                    wordcellrange.Text = persons[i].ShortString();

                    wordcellrange = wordtable.Rows.Last.Cells[3].Range;
                    wordcellrange.Text = persons[i].Birthday.Year.ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[4].Range;
                    wordcellrange.Text = persons[i].SubjectRF;

                    wordcellrange = wordtable.Rows.Last.Cells[5].Range;
                    wordcellrange.Text = persons[i].SrBall.ToString();

                    wordcellrange = wordtable.Rows.Last.Cells[6].Range;
                    wordcellrange.Text = persons[i].Ball.ToString();
                }

                ///////////////////////////////
            }
            # endregion
        }
        
        public void ListStart(List<Person> persons)
        {
            foreach (var person in persons)
            {
                FileName = Pathes.CreateTempPath(person.ShortString());
                InitializeShablon(Pathes.ListShablon, FileName);
                
                Word.Table wordtable = worddocument.Tables[2];

                var wordcellrange = wordtable.Cell(2, 3).Range;
                string fio = person.ToString();
                fio += Environment.NewLine;
                fio += person.Telefon;
                wordcellrange.Text = fio;
                string row3 = person.Birthday.ToLongDateString();
                row3 += Environment.NewLine + person.BirthdayMesto;
                wordcellrange = wordtable.Cell(3, 3).Range;
                wordcellrange.Text = row3;
                wordcellrange = wordtable.Cell(4, 3).Range;
                wordcellrange.Text = person.Country;

                wordcellrange = wordtable.Cell(5, 3).Range;
                wordcellrange.Text = person.VK;

                wordcellrange = wordtable.Cell(6, 3).Range;
                wordcellrange.Text = person.Health;

                wordcellrange = wordtable.Cell(7, 3).Range;
                wordcellrange.Text = person.VUZ;

                wordcellrange = wordtable.Cell(8, 3).Range;
                wordcellrange.Text = person.Specialnost;

                wordcellrange = wordtable.Cell(9, 3).Range;
                wordcellrange.Text = person.Diplom;

                wordcellrange = wordtable.Cell(10, 3).Range;
                wordcellrange.Text = person.VKR;

                wordcellrange = wordtable.Cell(11, 3).Range;
                wordcellrange.Text = person.Statya;

                wordcellrange = wordtable.Cell(12, 3).Range;
                wordcellrange.Text = person.SienceName;

                wordcellrange = wordtable.Cell(13, 3).Range;
                wordcellrange.Text = person.Soiskatelstvo == 0 ? "Нет" : "Да" + ", " + person.Exams;

                wordcellrange = wordtable.Cell(14, 3).Range;
                wordcellrange.Text = person.SportName;

                wordcellrange = wordtable.Cell(15, 3).Range;
                wordcellrange.Text = person.Napravlenie;

                wordcellrange = wordtable.Cell(16, 3).Range;
                wordcellrange.Text = person.Dopusk;

                wordcellrange = wordtable.Cell(17, 3).Range;
                wordcellrange.Text = person.Info;

                SaveAndQuit();
            }
        }

        public void ToVisible()
        {
            SaveAndQuit();
            InitializeShablonToShow(FileName);
        }

        public void Close()
        {
            wordapp = null;
            worddocument = null;
            GC.Collect();
            Thread.Sleep(1000);
            GC.Collect();
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Подбор_кандидатов__выгрузка_данных
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
        private string SaveAndQuitPDF(Person person)
        {
            var tmp = $"{Environment.CurrentDirectory}\\{Pathes.DirectoryGenerateList}\\{person.FileNamePDF()}";
            worddocument.SaveAs(tmp, Word.WdSaveFormat.wdFormatPDF);
            wordapp.Quit();
            return tmp;
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
        
        public string ListStart(Person person)
        {
            FileName = $"{Environment.CurrentDirectory}\\{Pathes.DirectoryGenerateList}\\{person.FileNameDOCX()}";
            InitializeShablon(Pathes.ListShablon, FileName);
                
            Word.Table wordtable = worddocument.Tables[2];

            var wordcellrange = wordtable.Cell(2, 3).Range;
            string fio = person.LongString();
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

            return SaveAndQuitPDF(person);
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

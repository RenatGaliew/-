using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Drive.v2;
using Google.Apis.Drive.v2.Data;
using Newtonsoft.Json;
using File = System.IO.File;

namespace Подбор_кандидатов__выгрузка_данных
{
    public partial class Form1 : Form
    {
        private static string XLSXPATH;
        private static string XLSXPATHPhotos;
        private static string TextFileDB = "TextFileDB.txt";

        private static MailService Services { get; set; }
        private static InternetChecker IC { get; set; }
        private WordWork word = new WordWork();
        private AccessExcel _excel;
        
        private List<Person> Persons;
        private static bool IsEnternetExist { get; set; }

        public Form1()
        {
            Persons = new List<Person>();
            InitializeComponent();
        }

        private void Tmp(Action action)
        {
            new Task(() =>
            {
                var taskTmp = new Task(action);
                taskTmp.Start();
                taskTmp.Wait();
            }).Start();
        }
        
        private void PopulateTreeView()
        {
            FilesResource.ListRequest list = Services.DriveService.Files.List();
            list.OrderBy = "createdDate";
            list.MaxResults = 100000;

            FileList filesFeed = list.Execute();
            var files = filesFeed.Items;
            spreadsLB.Items.Clear();
            sheetsLB.Items.Clear();
            foreach (var file in files)
            {
                if (file.MimeType == "application/vnd.google-apps.spreadsheet")
                {
                    spreadsLB.Items.Add(new FileToDownload(file));
                    sheetsLB.Items.Add(new FileToDownload(file));
                }
            }
        }

        private void btnConnect_Click(object sender, EventArgs e)
        {
            if (IsEnternetExist)
            {
                Tmp(() =>
                {
                    Services = new MailService();
                    BeginInvoke(new MethodInvoker(PopulateTreeView));
                });
            }
        }

        private void updateBTN_Click(object sender, EventArgs e)
        {
            if (IsEnternetExist)
            {
                Tmp(() =>
                {
                    Services = new MailService();
                    BeginInvoke(new MethodInvoker(PopulateTreeView));
                });
            }
        }

        private void downloadBTN_Click(object sender, EventArgs e)
        {
            try
            {
                var file = (FileToDownload)sheetsLB.SelectedItem;
                string filenameXLSX = $"{Pathes.DownloadsPath}{file.File.Title}.xlsx";
                var link = file.File.ExportLinks.Values.First(s => s.Contains("xlsx"));
                WebClient myWebClient = new WebClient();
                myWebClient.DownloadFile(link, filenameXLSX);
                XLSXPATHPhotos = filenameXLSX;
                ReadExcelPhotos(XLSXPATHPhotos);
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }

        private void ReadExcelPhotos(string pathFile)
        {
            _excel = new AccessExcel();
            _excel.DoAccess(pathFile);
            photosListView.Items.Clear();

            int count = _excel.MaxRows();
            for (int i = 2; i < count; i++)
            {
                int j = 2;
                
                string email = _excel.ReadCellString(i, j++);
                Person p = Persons.FirstOrDefault(person => person.email == email);
                if (p == null)
                {
                    continue;
                }

                p.SoglasieURLPath = _excel.ReadCellString(i, j++);
                p.ListSobesURLPath = _excel.ReadCellString(i, j++);
                p.DiplomURLPath = _excel.ReadCellString(i, j++);
                p.StatyaURLPath[0] = _excel.ReadCellString(i, j++);
                p.StatyaURLPath[1] = _excel.ReadCellString(i, j++);
                p.StatyaURLPath[2] = _excel.ReadCellString(i, j++);
                p.StatyaURLPath[3] = _excel.ReadCellString(i, j++);
                p.StatyaURLPath[4] = _excel.ReadCellString(i, j++);
                p.StatyaURLPath[5] = _excel.ReadCellString(i, j++);
                p.OlympURLPath[0] = _excel.ReadCellString(i, j++);
                p.OlympURLPath[1] = _excel.ReadCellString(i, j++);
                p.OlympURLPath[2] = _excel.ReadCellString(i, j++);
                p.OlympURLPath[3] = _excel.ReadCellString(i, j++);
                p.OlympURLPath[4] = _excel.ReadCellString(i, j++);
                p.OlympURLPath[5] = _excel.ReadCellString(i, j++);
                p.OlympURLPath[6] = _excel.ReadCellString(i, j++);
                p.KandidatURLPath[0] = _excel.ReadCellString(i, j++);
                p.KandidatURLPath[1] = _excel.ReadCellString(i, j++);
                p.KandidatURLPath[2] = _excel.ReadCellString(i, j++);
                p.WorkURLPath[0] = _excel.ReadCellString(i, j++);
                p.WorkURLPath[1] = _excel.ReadCellString(i, j++);
                p.WorkURLPath[2] = _excel.ReadCellString(i, j++);
                p.SportURLPath[0] = _excel.ReadCellString(i, j++);
                p.SportURLPath[1] = _excel.ReadCellString(i, j++);
                p.IsEtap2 = true;
            }

            photosListView.Items.Clear();
            foreach (var person in Persons)
            {
                if (person.IsEtap2)
                {
                    AddPersonInListPhotos(person);
                }

                word.Close();
            }

            _excel.FinishAccess();
        }

        private void DownloadPDF(string inputPath, string outputPath)
        {
            Services.DownloadFile(inputPath, outputPath);
        }
        
        private void ReadExcel(string pathFile)
        {
            _excel = new AccessExcel();
            _excel.DoAccess(pathFile);
            emailsListView.Items.Clear();

            int count = _excel.MaxRows();
            for (int i = 2; i < count; i++)
            {
                int j = 2;

                #region PersonReadExcel
                Person p = new Person();
                p.email = _excel.ReadCellString(i, j++);
                if (!Persons.Exists(person => person.email == p.email))
                {
                    p.URLPhoto = _excel.ReadCellString(i, j++);
                    p.SurName = _excel.ReadCellString(i, j++);
                    p.Name = _excel.ReadCellString(i, j++);
                    p.Patronomyc = _excel.ReadCellString(i, j++);
                    p.Telefon = _excel.ReadCellString(i, j++);
                    p.Birthday = _excel.ReadDate(i, j++);
                    p.BirthdayMesto = _excel.ReadCellString(i, j++);
                    p.Country = _excel.ReadCellString(i, j++);
                    p.Addres = _excel.ReadCellString(i, j++);
                    p.AddresRegistry = _excel.ReadCellString(i, j++);
                    p.SubjectRF = _excel.ReadCellString(i, j++);
                    p.VK = _excel.ReadCellString(i, j++);
                    p.Health = _excel.ReadCellString(i, j++);
                    p.HealtHron = _excel.ReadCellString(i, j++);
                    p.VUZ = _excel.ReadCellString(i, j++);
                    p.VUZKor = _excel.ReadCellString(i, j++);
                    p.Specialnost = _excel.ReadCellString(i, j++);
                    p.Diplom = _excel.ReadCellString(i, j++);
                    p.SrBall = _excel.ReadCellDouble(i, j++);
                    p.VKR = _excel.ReadCellString(i, j++);
                    p.Soiskatelstvo = _excel.ReadCellString(i, j++) == "Да" ? 5 : 0;
                    p.Exams = _excel.ReadCellString(i, j++);
                    p.Statiy[0] = _excel.ReadCellString(i, j++) == "Есть" ? 5 : 0;
                    p.Statiy[1] = _excel.ReadCellString(i, j++) == "Есть" ? 4 : 0;
                    p.Statiy[2] = _excel.ReadCellString(i, j++) == "Есть" ? 3 : 0;
                    p.Statiy[3] = _excel.ReadCellString(i, j++) == "Есть" ? 1 : 0;
                    p.Statiy[4] = _excel.ReadCellString(i, j++) == "Есть" ? 1 : 0;
                    p.Statiy[5] = _excel.ReadCellString(i, j++) == "Есть" ? 0.5 : 0;
                    p.Sience[0] = _excel.ReadCellString(i, j++) == "Есть" ? 4 : 0;
                    p.Sience[1] = _excel.ReadCellString(i, j++) == "Есть" ? 4 : 0;
                    p.Sience[2] = _excel.ReadCellString(i, j++) == "Есть" ? 3 : 0;
                    p.Sience[3] = _excel.ReadCellString(i, j++) == "Есть" ? 3 : 0;
                    p.Sience[4] = _excel.ReadCellString(i, j++) == "Есть" ? 3 : 0;
                    p.Sience[5] = _excel.ReadCellString(i, j++) == "Есть" ? 2 : 0;
                    p.Sience[6] = _excel.ReadCellString(i, j++) == "Есть" ? 1 : 0;
                    p.SienceName = _excel.ReadCellString(i, j++);
                    p.SienceStepen[0] = _excel.ReadCellString(i, j++) == "Есть" ? 3 : 0;
                    p.SienceStepen[1] = _excel.ReadCellString(i, j++) == "Есть" ? 6 : 0;
                    p.SienceStepen[2] = _excel.ReadCellString(i, j++) == "Есть" ? 8 : 0;
                    p.Work[0] = _excel.ReadCellString(i, j++) == "Есть" ? 2 : 0;
                    p.Work[1] = _excel.ReadCellString(i, j++) == "Есть" ? 4 : 0;
                    p.Work[2] = _excel.ReadCellString(i, j++) == "Есть" ? 6 : 0;
                    p.Sport[0] = _excel.ReadCellString(i, j++) == "Есть" ? 4 : 0;
                    p.Sport[1] = _excel.ReadCellString(i, j++) == "Есть" ? 2 : 0;
                    p.Language = _excel.ReadCellString(i, j++);
                    p.Products = _excel.ReadCellString(i, j++);
                    p.Napravlenie = _excel.ReadCellString(i, j++);
                    p.Dopusk = _excel.ReadCellString(i, j++);
                    p.TattooAndPirsing = _excel.ReadCellString(i, j++);
                    p.Rost = _excel.ReadCellDouble(i, j++);
                    p.Ves = _excel.ReadCellDouble(i, j++);
                    p.Family = _excel.ReadCellString(i, j++);
                    p.Child = _excel.ReadCellString(i, j++);
                    p.Info = _excel.ReadCellString(i, j++);

                    p.Prioritet[0] = _excel.ReadCellString(i, 100) == "Есть" ? 3 : 0;
                    p.Prioritet[1] = _excel.ReadCellString(i, 101) == "Есть" ? 1 : 0;
                    p.Reserv = _excel.ReadCellString(i, 102) == "Резерв";
                    p.Magistr = _excel.ReadCellString(i, 103) == "Магистр";
                    p.Bakalavr = _excel.ReadCellString(i, 103) == "Бакаларв";
                    p.Zachetka = _excel.ReadCellString(i, 103) == "Зачетка";
                    p.Ball = p.ExecuteBall();

                    #endregion

                    emailsListView.Items.Add(p.email, p.ToString(), p.Status);
                    Persons.Add(p);
                }
            }

            emailsListView.Items.Clear();
            foreach (var person in Persons)
            {
                AddPersonInList(emailsListView, person);
                if (string.IsNullOrEmpty(person.PDFPath) && !File.Exists(person.PDFPath))
                {
                    person.PDFPath = word.ListStart(person);
                }

                word.Close();
            }

            _excel.FinishAccess();
        }

        private void sendMessage_Click(object sender, EventArgs e)
        {
            var list = Persons.Where(person => !person.IsEmailSending).ToList();

            foreach (var person in list)
            {
                statusLBL.Text = "Отправляется";
                person.IsSendingChanged += PersonOnIsSendingChanged;
                SendMessage(person);
            }
        }

        private void PersonOnIsSendingChanged(object sender, EventArgs eventArgs)
        {
            foreach (ListViewItem item in emailsListView.Items)
            {
                if (item.Tag is Person person)
                {
                    item.BackColor = person.IsEmailSending ? Color.LightGreen : Color.Pink;
                }
            }
        }

        private void SendMessage(Person person)
        {
            #region sendMessage

            Tmp(() =>
            {
                bool IsSending = Services.SendMail(person);
                person.IsEmailSending = IsSending;
            });

            #endregion
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer", $"{Environment.CurrentDirectory}\\{Pathes.DirectoryPDFS}");
        }

        private void IfNotDirectoryCreate(string path)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
        }

        private void InitializeAllDirectories()
        {
            IfNotDirectoryCreate(Pathes.DirectoryPDFS);
            IfNotDirectoryCreate(Pathes.DownloadsPath);
            IfNotDirectoryCreate(Pathes.DownloadsPathPhotos);
            IfNotDirectoryCreate(Pathes.DirectoryGenerateList);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            string jsonString = JsonConvert.SerializeObject(Persons);
            File.WriteAllText(TextFileDB,jsonString);
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            if (File.Exists(TextFileDB))
            {
                Persons = JsonConvert.DeserializeObject<List<Person>>(File.ReadAllText(TextFileDB));
                foreach (var person in Persons)
                {
                    person.IsSendingChanged += PersonOnIsSendingChanged;
                    person.IsFilesDownloaded += PersonOnIsFilesDownloaded;
                    AddPersonInList(emailsListView, person);
                }
                foreach (var person in Persons)
                {
                    AddPersonInListPhotos(person);
                    person.IsAllFILESDownloaded();
                }
            }

            Tmp(() =>
            {
                IC = new InternetChecker();
                IC.StatusChanging += IcOnStatusChanging;
                IC.Start();
            });

            InitializeAllDirectories();
            InitializePDFFiles();
        }

        private void InitializePDFFiles()
        {
            var files = Directory.GetFiles(Pathes.DirectoryPDFS);
            foreach (string file in files)
            {
                if (file.Split('/').Length == 2)
                    pdfList.Items.Add(file.Split('/')[1]);
            }
        }

        private void IcOnStatusChanging(object sender, InternetStatus internetStatus)
        {
            switch (internetStatus)
            {
                case InternetStatus.YesInternet:
                    BeginInvoke(new MethodInvoker(delegate
                    {
                        statusLabel.Text = "Готово";
                        IsEnternetExist = true;
                        btnConnect.Enabled = true;
                        updateBTN.Enabled = true;
                    }));
                    break;
                case InternetStatus.Searching:
                    BeginInvoke(new MethodInvoker(delegate
                    {
                        statusLabel.Text = "Поиск интернета";
                        IsEnternetExist = false;
                        btnConnect.Enabled = false;
                        updateBTN.Enabled = false;
                    }));
                    break;
                case InternetStatus.SearchingAgain:
                    BeginInvoke(new MethodInvoker(delegate
                    {
                        statusLabel.Text = "Переподключение";
                        IsEnternetExist = false;
                        btnConnect.Enabled = false;
                        updateBTN.Enabled = false;
                    }));
                    break;
                case InternetStatus.NoInternet:
                    BeginInvoke(new MethodInvoker(delegate
                    {
                        statusLabel.Text = "Нет интернета";
                        IsEnternetExist = false;
                        btnConnect.Enabled = false;
                        updateBTN.Enabled = false;
                    }));
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(internetStatus), internetStatus, null);
            } 
        }

        private void PersonOnIsFilesDownloaded(object sender, EventArgs eventArgs)
        {
            foreach (ListViewItem item in photosListView.Items)
            {
                if (item.Tag is Person person)
                {
                    item.BackColor = person.IsEmailSending ? Color.LightGreen : Color.Pink;
                }
            }
        }

        private void AddPersonInList(ListView lv, Person person)
        {
            var row = new[] { person.email, person.ToString(), person.Status };
            var lvi = new ListViewItem(row)
            {
                Tag = person,
                Checked = person.IsEmailSending,
            };
            lv.Items.Add(lvi);
        }

        private void AddPersonInListPhotos(Person person)
        {
            var row = new[] { person.ToString() };
            var lvi = new ListViewItem(row)
            {
                Tag = person,
                Checked = person.IsEmailSending,
            };
            photosListView.Items.Add(lvi);
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem selectedItem in emailsListView.SelectedItems)
            {
                if (selectedItem.Tag is Person person)
                    SendMessage(person);
            }
        }

        private void openPDFBTN_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem selectedItem in emailsListView.SelectedItems)
            {
                if (selectedItem.Tag is Person person)
                    System.Diagnostics.Process.Start(person.PDFPath);
            }
        }

        private void dwnloadBTN_Click(object sender, EventArgs e)
        {
            try
            {
                var file = (FileToDownload)spreadsLB.SelectedItem;
                string filenameXLSX = $"{Pathes.DownloadsPath}{file.File.Title}.xlsx";
                var link = file.File.ExportLinks.Values.First(s => s.Contains("xlsx"));
                WebClient myWebClient = new WebClient();
                myWebClient.DownloadFile(link, filenameXLSX);
                XLSXPATH = filenameXLSX;
                ReadExcel(XLSXPATH);
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dwnloadFilesBTN_Click(object sender, EventArgs e)
        {
            foreach (var p in Persons)
            {
                var path = $"{Pathes.DownloadsPath}{p.email}";
                if (p.IsEtap2)
                {
                    #region DOWNLOAD_FILES
                    if (!string.IsNullOrWhiteSpace(p.URLPhoto))
                    {
                        string filenamePhoto = $"{path}/{p.SurName}_{p.Name}_{p.Patronomyc}.png";
                        IfNotDirectoryCreate(path);
                        Services.DownloadFile(p.URLPhoto, filenamePhoto);
                        p.PhotoPath = filenamePhoto;

                        using (Image image = Image.FromFile(filenamePhoto))
                        {
                            using (MemoryStream m = new MemoryStream())
                            {
                                image.Save(m, image.RawFormat);
                            }
                        }
                    }

                    if (!string.IsNullOrEmpty(p.SoglasieURLPath))
                    {
                        string filename = $"{path}/{p.SurName}_{p.Name}_{p.Patronomyc}_СОГЛАСИЕ.pdf";
                        DownloadPDF(p.SoglasieURLPath, filename);
                        p.SoglasiePath = filename;
                    }

                    if (!string.IsNullOrEmpty(p.DiplomURLPath))
                    {
                        string filename = $"{path}/{p.SurName}_{p.Name}_{p.Patronomyc}_ДИПЛОМ.pdf";
                        DownloadPDF(p.DiplomURLPath, filename);
                        p.DiplomPath = filename;
                    }

                    if (!string.IsNullOrEmpty(p.ListSobesURLPath))
                    {
                        string filename = $"{path}/{p.SurName}_{p.Name}_{p.Patronomyc}_ЛИСТ_БЕСЕДЫ.pdf";
                        DownloadPDF(p.ListSobesURLPath, filename);
                        p.ListSobesPath = filename;
                    }
                    int k = 1;
                    foreach (string s in p.StatyaURLPath)
                    {
                        if (!string.IsNullOrEmpty(s))
                        {
                            string filename = $"{path}/{p.SurName}_{p.Name}_{p.Patronomyc}_Статья_№{k}.pdf";
                            DownloadPDF(s, filename);
                            p.StatyaPath[k - 1] = filename;
                            k++;
                        }
                    }

                    k = 1;
                    foreach (string s in p.OlympURLPath)
                    {
                        if (!string.IsNullOrEmpty(s))
                        {
                            string filename = $"{path}/{p.SurName}_{p.Name}_{p.Patronomyc}_Олимпиада_№{k}.pdf";
                            DownloadPDF(s, filename);
                            p.OlympPath[k - 1] = filename;
                            k++;
                        }
                    }

                    k = 1;
                    foreach (string s in p.SportURLPath)
                    {
                        if (!string.IsNullOrEmpty(s))
                        {
                            string filename = $"{path}/{p.SurName}_{p.Name}_{p.Patronomyc}_Спорт_№{k}.pdf";
                            DownloadPDF(s, filename);
                            p.SportPath[k - 1] = filename;
                            k++;
                        }
                    }

                    k = 1;
                    foreach (string s in p.KandidatURLPath)
                    {
                        if (!string.IsNullOrEmpty(s))
                        {
                            string filename = $"{path}/{p.SurName}_{p.Name}_{p.Patronomyc}_Кандидат_№{k}.pdf";
                            DownloadPDF(s, filename);
                            p.KandidatPath[k - 1] = filename;
                            k++;
                        }
                    }

                    k = 1;
                    foreach (string s in p.WorkURLPath)
                    {
                        if (!string.IsNullOrEmpty(s))
                        {
                            string filename = $"{path}/{p.SurName}_{p.Name}_{p.Patronomyc}_РАБОТА_№{k}.pdf";
                            DownloadPDF(s, filename);
                            p.WorkPath[k - 1] = filename;
                            k++;
                        }
                    }
                    #endregion

                    p.IsAllFILESDownloaded();
                }
            }
        }
    }
}


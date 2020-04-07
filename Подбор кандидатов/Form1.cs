using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Auth.OAuth2.Flows;
using Google.Apis.Auth.OAuth2.Responses;
using Google.Apis.Drive.v2;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using static System.Char;

namespace Подбор_кандидатов
{
    public partial class Form1 : Form
    {
        private AccessExcel _excel;
        private int Max = 100;
        private WordWork word = new WordWork();
        private List<Person> Persons;
        private Person SelectedPerson { get; set; }
        private string PathFile { get; set; }
        private ErrorProvider errorProvider1;
        private static string code = "qrwhu_5RKGUpBuc6joycjiO6";
        private static string idClient = "client_secret_572291057451-2mai9prgom68ufefoj5p0709sqji23s1.apps.googleusercontent.com.json";

        public Form1()
        {
            InitializeComponent();
            Persons = new List<Person>();
            changeBTN.Enabled = listBox1.SelectedItem != null;
            vesTB.TextChanged += TextBoxOnValidated;
            rostTB.TextChanged += TextBoxOnValidated;
            srBallTB.TextChanged += TextBoxOnValidated;
            errorProvider1 = new ErrorProvider();
            errorProvider1.SetIconAlignment(vesTB, ErrorIconAlignment.MiddleRight);
            errorProvider1.SetIconAlignment(rostTB, ErrorIconAlignment.MiddleRight);
            errorProvider1.SetIconAlignment(srBallTB, ErrorIconAlignment.MiddleRight);

            var controls = groupBox1.Controls.OfType<Control>().ToArray();
            foreach(var control in controls)
            {
                if(control is TextBox tb)
                {
                    tb.TextChanged += TbOnTextChanged;
                }
                if (control is CheckedListBox lb)
                {
                    lb.ItemCheck += LbOnItemCheck;
                }
            }


        }

        private static void SaveStream(System.IO.MemoryStream stream, string saveTo)
        {
            using (System.IO.FileStream file = new System.IO.FileStream(saveTo, System.IO.FileMode.Create, System.IO.FileAccess.Write))
            {
                stream.WriteTo(file);
            }
        }

        private static void DownloadFile(DriveService service, Google.Apis.Drive.v2.Data.File file, string saveTo)
        {
            var request = service.Files.Get(file.Id);
            var stream = new System.IO.MemoryStream();

            // Add a handler which will be notified on progress changes.
            // It will notify on each chunk download and when the
            // download is completed or failed.
            request.MediaDownloader.ProgressChanged += (Google.Apis.Download.IDownloadProgress progress) =>
            {
                switch (progress.Status)
                {
                    case Google.Apis.Download.DownloadStatus.Downloading:
                    {
                        Console.WriteLine(progress.BytesDownloaded);
                        break;
                    }
                    case Google.Apis.Download.DownloadStatus.Completed:
                    {
                        Console.WriteLine("Download complete.");
                        SaveStream(stream, saveTo);
                        break;
                    }
                    case Google.Apis.Download.DownloadStatus.Failed:
                    {
                        Console.WriteLine("Download failed.");
                        break;
                    }
                }
            };
            request.Download(stream);

        }

        /// <summary>
        /// This method requests Authentcation from a user using Oauth2.  
        /// Credentials are stored in System.Environment.SpecialFolder.Personal
        /// Documentation https://developers.google.com/accounts/docs/OAuth2
        /// </summary>
        /// <param name="clientSecretJson">PathFile to the client secret json file from Google Developers console.</param>
        /// <param name="userName">Identifying string for the user who is being authentcated.</param>
        /// <returns>DriveService used to make requests against the Drive API</returns>
        public static DriveService AuthenticateOauth(string clientSecretJson, string userName)
        {
            try
            {
                if (string.IsNullOrEmpty(userName))
                    throw new ArgumentNullException("userName");
                if (string.IsNullOrEmpty(clientSecretJson))
                    throw new ArgumentNullException("clientSecretJson");
                if (!File.Exists(clientSecretJson))
                    throw new Exception("clientSecretJson file does not exist.");

                // These are the scopes of permissions you need. It is best to request only what you need and not all of them
                string[] scopes = new string[] { DriveService.Scope.DriveReadonly };         	//View the files in your Google Drive                                                 
               

                UserCredential credential;

                using (var stream = new FileStream(clientSecretJson, FileMode.Open, FileAccess.Read))
                {
                    string credPath = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".credentials/", System.Reflection.Assembly.GetExecutingAssembly().GetName().Name);

                    // Requesting Authentication or loading previously stored authentication for userName
                    var secret = GoogleClientSecrets.Load(stream).Secrets;
                    secret.ClientSecret = code;
                    //secret.ClientId = idClient;
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(secret,
                                                                             scopes,
                                                                             userName,
                                                                             CancellationToken.None,
                                                                             new FileDataStore(credPath, true)).Result;
                }

                // Create Drive API service.
                return new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Drive Oauth2 Authentication Sample"
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine("Create Oauth2 account DriveService failed" + ex.Message);
                throw new Exception("CreateServiceAccountDriveFailed", ex);
            }
        }

        private void LbOnItemCheck(object sender, ItemCheckEventArgs itemCheckEventArgs)
        {
            Proverks();
        }

        private void TbOnTextChanged(object sender, EventArgs eventArgs)
        {
            Proverks();
        }

        private void Proverks()
        {
            if (SelectedPerson.OK())
            {
                SaveBTN.BackColor = Color.PaleGreen;
            }
            else
                SaveBTN.BackColor = Color.PaleVioletRed;

            var controls = groupBox1.Controls.OfType<Control>().ToArray();
            foreach (var control in controls)
            {
                if (control is TextBox tb)
                {
                    if (tb.Text == "")
                        tb.BackColor = Color.PaleVioletRed;
                    else
                    {
                        tb.BackColor = SystemColors.Control;
                    }
                }
            }
            if (sportNameTB.Text == "")
                sportNameTB.BackColor = Color.PaleVioletRed;
            else
                sportNameTB.BackColor = SystemColors.Control;
        }

        private void TextBoxOnValidated(object sender, EventArgs eventArgs)
        {
            if(sender is TextBox tb)
                if (!IsNum(tb.Text))
                {
                    errorProvider1.SetError(tb, "Введите корректное число!");
                    tb.Select();
                }
                else
                {
                    if (double.Parse(tb.Text) <= 0)
                    {
                        errorProvider1.SetError(tb, "Введите корректное число!");
                        tb.Select();
                    }
                    else
                    {
                        errorProvider1.SetError(tb, string.Empty);
                    }
                }
        }

        /// <summary>
        /// метод при изменении выделенного кандидата
        /// </summary>
        private void listBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            if ((Person) listBox1.SelectedItem != null)
            {
                changeBTN.Enabled = true;
                SelectedPerson = Persons.First(x => x.ID == ((Person) listBox1.SelectedItem).ID);
                ball.Text = SelectedPerson.Ball.ToString();
                telefonTB.Text = SelectedPerson.Telefon;
                birthdayMestoTB.Text = SelectedPerson.BirthdayMesto;
                vuzTB.Text = SelectedPerson.VUZ;
                specialnostTB.Text = SelectedPerson.Specialnost;
                diplomTB.Text = SelectedPerson.Diplom;
                vkrTB.Text = SelectedPerson.VKR;
                soiscatelstvoCB.Checked = SelectedPerson.Soiskatelstvo != 0;
                examsTB.Text = SelectedPerson.Exams;
                napravlenieTB.Text = SelectedPerson.Napravlenie;
                dopuskTB.Text = SelectedPerson.Dopusk;
                srBallTB.Text = SelectedPerson.SrBall.ToString();
                familyTB.Text = SelectedPerson.Family;
                tattooAndPirsingTB.Text = SelectedPerson.TattooAndPirsing;
                healthHronTB.Text = SelectedPerson.HealtHron;
                subjectRFTB.Text = SelectedPerson.SubjectRF;
                languageTB.Text = SelectedPerson.Language;
                productsTB.Text = SelectedPerson.Products;
                rostTB.Text = SelectedPerson.Rost.ToString();
                vesTB.Text = SelectedPerson.Ves.ToString();
                statiyNameTB.Text = SelectedPerson.Statya;
                sienceNameTB.Text = SelectedPerson.SienceName;
                workNameTB.Text = SelectedPerson.WorkName;
                sportNameTB.Text = SelectedPerson.SportName;
                countryTB.Text = SelectedPerson.Country;
                surnameTB.Text = SelectedPerson.SurName;
                nameTB.Text = SelectedPerson.Name;
                patronomycTB.Text = SelectedPerson.Patronomyc;
                adressTB.Text = SelectedPerson.Addres;
                adresRegTB.Text = SelectedPerson.AddresRegistry;
                vuzKor.Text = SelectedPerson.VUZKor;
                vkTB.Text = SelectedPerson.VK;
                healthTB.Text = SelectedPerson.Health;
                infoTB.Text = SelectedPerson.Info;
                dateTimePicker1.Value = SelectedPerson.Birthday;
                reservCB.Checked = SelectedPerson.Reserv;
                magCB.Checked = SelectedPerson.Magistr;
                BakCB.Checked = SelectedPerson.Bakalavr;
                ZachetCB.Checked = SelectedPerson.Zachetka;
                //pictureBox1.Image = Image.FromFile(SelectedPerson.photoPath);
                
                for (int i = 0; i < SelectedPerson.Statiy.Length; i++)
                {
                    StatiyCLB.SetItemChecked(i, SelectedPerson.Statiy[i] != 0);
                }

                for (int i = 0; i < SelectedPerson.Sience.Length; i++)
                {
                    OlympCLB.SetItemChecked(i, SelectedPerson.Sience[i] != 0);
                }

                for (int i = 0; i < SelectedPerson.Work.Length; i++)
                {
                    WorkCLB.SetItemChecked(i, SelectedPerson.Work[i] != 0);
                }

                for (int i = 0; i < SelectedPerson.Sport.Length; i++)
                {
                    SportCLB.SetItemChecked(i, SelectedPerson.Sport[i] != 0);
                }

                for (int i = 0; i < SelectedPerson.SienceStepen.Length; i++)
                {
                    KandCLB.SetItemChecked(i, SelectedPerson.SienceStepen[i] != 0);
                }
                for (int i = 0; i < SelectedPerson.Prioritet.Length; i++)
                {
                    PrioritetCLB.SetItemChecked(i, SelectedPerson.Prioritet[i] != 0);
                }
            }
        }

        /// <summary>
        /// метод для задачи
        /// </summary>
        private void Tmp(Action action)
        {
            new Task(() =>
            {
                var taskTmp = new Task(action);
                taskTmp.Start();
                taskTmp.Wait();
                word.ToVisible();
                word.Close();
                BeginInvoke(new MethodInvoker(delegate
                {
                    //гиф
                }));
            }).Start();
        }

        /// <summary>
        /// Нажатие кнопки Сохранить
        /// </summary>
        private void button5_Click(object sender, EventArgs e)
        {
            SelectedPerson.SurName = surnameTB.Text;
            SelectedPerson.Name = nameTB.Text;
            SelectedPerson.Patronomyc = patronomycTB.Text;
            SelectedPerson.Telefon = telefonTB.Text;
            SelectedPerson.Birthday = dateTimePicker1.Value;
            SelectedPerson.BirthdayMesto = birthdayMestoTB.Text;
            SelectedPerson.Country = countryTB.Text;
            SelectedPerson.Addres = adressTB.Text;
            SelectedPerson.AddresRegistry = adresRegTB.Text;
            SelectedPerson.SubjectRF = subjectRFTB.Text;
            SelectedPerson.VK = vkTB.Text;
            SelectedPerson.Health = healthTB.Text;
            SelectedPerson.HealtHron = healthHronTB.Text;
            SelectedPerson.VUZ = vuzTB.Text;
            SelectedPerson.VUZKor = vuzKor.Text;
            SelectedPerson.Specialnost = specialnostTB.Text;
            SelectedPerson.Diplom = diplomTB.Text;
            SelectedPerson.SrBall = double.Parse(srBallTB.Text);
            SelectedPerson.VKR = vkrTB.Text;
            SelectedPerson.Soiskatelstvo = soiscatelstvoCB.Checked ? 5 : 0;
            SelectedPerson.Exams = examsTB.Text;
            SelectedPerson.Statiy[0] = StatiyCLB.GetItemChecked(0) ? 5 : 0;
            SelectedPerson.Statiy[1] = StatiyCLB.GetItemChecked(1) ? 4 : 0;
            SelectedPerson.Statiy[2] = StatiyCLB.GetItemChecked(2) ? 3 : 0;
            SelectedPerson.Statiy[3] = StatiyCLB.GetItemChecked(3) ? 1 : 0;
            SelectedPerson.Statiy[4] = StatiyCLB.GetItemChecked(4) ? 1 : 0;
            SelectedPerson.Statiy[5] = StatiyCLB.GetItemChecked(5) ? 0.5 : 0;
            SelectedPerson.Statya = statiyNameTB.Text;
            SelectedPerson.Sience[0] = OlympCLB.GetItemChecked(0) ? 4 : 0;
            SelectedPerson.Sience[1] = OlympCLB.GetItemChecked(1) ? 4 : 0;
            SelectedPerson.Sience[2] = OlympCLB.GetItemChecked(2) ? 3 : 0;
            SelectedPerson.Sience[3] = OlympCLB.GetItemChecked(3) ? 3 : 0;
            SelectedPerson.Sience[4] = OlympCLB.GetItemChecked(4) ? 3 : 0;
            SelectedPerson.Sience[5] = OlympCLB.GetItemChecked(5) ? 2 : 0;
            SelectedPerson.Sience[6] = OlympCLB.GetItemChecked(6) ? 1 : 0;
            SelectedPerson.SienceName = sienceNameTB.Text;
            SelectedPerson.SienceStepen[0] = KandCLB.GetItemChecked(0) ? 3 : 0;
            SelectedPerson.SienceStepen[1] = KandCLB.GetItemChecked(1) ? 6 : 0;
            SelectedPerson.SienceStepen[2] = KandCLB.GetItemChecked(2) ? 8 : 0;
            SelectedPerson.Work[0] = WorkCLB.GetItemChecked(0) ? 2 : 0;
            SelectedPerson.Work[1] = WorkCLB.GetItemChecked(1) ? 4 : 0;
            SelectedPerson.Work[2] = WorkCLB.GetItemChecked(2) ? 6 : 0;
            SelectedPerson.WorkName = workNameTB.Text;
            SelectedPerson.Sport[0] = SportCLB.GetItemChecked(0) ? 4 : 0;
            SelectedPerson.Sport[1] = SportCLB.GetItemChecked(1) ? 2 : 0;
            SelectedPerson.Prioritet[0] = PrioritetCLB.GetItemChecked(0) ? 3 : 0;
            SelectedPerson.Prioritet[1] = PrioritetCLB.GetItemChecked(1) ? 1 : 0;
            SelectedPerson.SportName = sportNameTB.Text;
            SelectedPerson.Language = languageTB.Text;
            SelectedPerson.Products = productsTB.Text;
            SelectedPerson.Napravlenie = napravlenieTB.Text;
            SelectedPerson.Dopusk = dopuskTB.Text;
            SelectedPerson.TattooAndPirsing = tattooAndPirsingTB.Text;
            SelectedPerson.Rost = double.TryParse(rostTB.Text, out double result1) ? result1 : SelectedPerson.Rost;
            SelectedPerson.Ves = double.TryParse(vesTB.Text, out double result2) ? result2 : SelectedPerson.Ves;
            SelectedPerson.Family = familyTB.Text;
            SelectedPerson.Info = infoTB.Text;
            SelectedPerson.Reserv = reservCB.Checked;
            SelectedPerson.Magistr = magCB.Checked;
            SelectedPerson.Zachetka = ZachetCB.Checked;
            SelectedPerson.Bakalavr = BakCB.Checked;
            SelectedPerson.Ball = SelectedPerson.ExecuteBall();
            ball.Text = SelectedPerson.Ball.ToString();

            /*if (!SelectedPerson.OK())
            {
                MessageBox.Show("Введите все данные", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }*/

            _excel = new AccessExcel();
            _excel.DoAccess(PathFile);

            groupBox1.Enabled = false;
            groupBox2.Enabled = false;
            listBox1.Enabled = true;
            changeBTN.Enabled = true;

            Persons.Sort(new NaturalStringComparer());
            Persons.Sort(new BallComparer());
            listBox1.Items.Clear();
            foreach (var person in Persons)
            {
                listBox1.Items.Add(person);
            }

            listBox1.SelectedItem = listBox1.Items.OfType<Person>().First(x => x.ID == SelectedPerson.ID);

            for (int i = 0; i < Persons.Count; i++)
            {
                _excel.WriteCell<string>(1, 1, "Дата изменения");
                int iRow = i + 2;
                var p = Persons[i];
                int j = 1;
                _excel.WriteCell<string>(iRow, j++, DateTime.Now.ToString("G"));
                _excel.WriteCell<string>(iRow, j++, p.URLPhoto);
                _excel.WriteCell<string>(iRow, j++, p.URLPhoto);
                _excel.WriteCell<string>(iRow, j++, p.URLSOGLASIE);
                _excel.WriteCell<string>(iRow, j++, p.SurName);
                _excel.WriteCell<string>(iRow, j++, p.Name);
                _excel.WriteCell<string>(iRow, j++, p.Patronomyc);
                _excel.WriteCell<string>(iRow, j++, p.Telefon);
                _excel.WriteCell<string>(iRow, j++, p.Birthday.ToShortDateString());
                _excel.WriteCell<string>(iRow, j++, p.BirthdayMesto);
                _excel.WriteCell<string>(iRow, j++, p.Country);
                _excel.WriteCell<string>(iRow, j++, p.Addres);
                _excel.WriteCell<string>(iRow, j++, p.AddresRegistry);
                _excel.WriteCell<string>(iRow, j++, p.SubjectRF);
                _excel.WriteCell<string>(iRow, j++, p.VK);
                _excel.WriteCell<string>(iRow, j++, p.Health);
                _excel.WriteCell<string>(iRow, j++, p.HealtHron);
                _excel.WriteCell<string>(iRow, j++, p.VUZ);
                _excel.WriteCell<string>(iRow, j++, p.VUZKor);
                _excel.WriteCell<string>(iRow, j++, p.Specialnost);
                _excel.WriteCell<string>(iRow, j++, p.Diplom);
                _excel.WriteCell<double>(iRow, j++, p.SrBall);
                _excel.WriteCell<string>(iRow, j++, p.VKR);
                _excel.WriteCell<string>(iRow, j++, p.Soiskatelstvo == 0 ? "Нет": "Да");
                _excel.WriteCell<string>(iRow, j++, p.Exams);
                _excel.WriteCell<string>(iRow, j++, p.Statiy[0] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Statiy[1] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Statiy[2] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Statiy[3] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Statiy[4] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Statiy[5] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Statya);
                _excel.WriteCell<string>(iRow, j++, p.Sience[0] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Sience[1] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Sience[2] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Sience[3] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Sience[4] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Sience[5] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Sience[6] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.SienceName);
                _excel.WriteCell<string>(iRow, j++, p.SienceStepen[0] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.SienceStepen[1] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.SienceStepen[2] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Work[0] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Work[1] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Work[2] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.WorkName);
                _excel.WriteCell<string>(iRow, j++, p.Sport[0] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Sport[1] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.SportName);
                _excel.WriteCell<string>(iRow, j++, p.Language);
                _excel.WriteCell<string>(iRow, j++, p.Products);
                _excel.WriteCell<string>(iRow, j++, p.Napravlenie);
                _excel.WriteCell<string>(iRow, j++, p.Dopusk);
                _excel.WriteCell<string>(iRow, j++, p.TattooAndPirsing);
                _excel.WriteCell<double>(iRow, j++, p.Rost);
                _excel.WriteCell<double>(iRow, j++, p.Ves);
                _excel.WriteCell<string>(iRow, j++, p.Family);
                _excel.WriteCell<string>(iRow, j++, p.Info);
                _excel.WriteCell<string>(iRow, 100, p.Prioritet[0] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, 101, p.Prioritet[1] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, 102, p.Reserv ? "Резерв" : "Основа");
                _excel.WriteCell<string>(iRow, 103, p.Magistr ? "Магистр" : p.Bakalavr ? "Бакалавр" : p.Zachetka ? "Зачетка" : "");
            }
            _excel.FinishAccess();
        }
        
        /// <summary>
        /// Нажатие кнопки изменить
        /// </summary>
        private void button6_Click(object sender, EventArgs e)
        {
            SelectedPerson = (Person)listBox1.SelectedItem;
            groupBox1.Enabled = true;
            groupBox2.Enabled = true;
            listBox1.Enabled = false;
            changeBTN.Enabled = false;
        }

        /// <summary>
        /// Нажатие кнопки Считать
        /// </summary>
        private void считатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void ReadExcel()
        {
            _excel = new AccessExcel();
            _excel.DoAccess(PathFile);
            Persons.Clear();
            listBox1.Items.Clear();

            int count = _excel.MaxRows();
            for (int i = 2; i < count; i++)
            {
                int j = 3;
                Person p = new Person();
                p.URLPhoto = _excel.ReadCell<string>(i, j++);
                p.URLSOGLASIE = _excel.ReadCell<string>(i, j++);
                p.SurName = _excel.ReadCell<string>(i, j++);
                p.Name = _excel.ReadCell<string>(i, j++);
                p.Patronomyc = _excel.ReadCell<string>(i, j++);
                p.Telefon = _excel.ReadCell<string>(i, j++);
                p.Birthday = _excel.ReadDate(i, j++);
                p.BirthdayMesto = _excel.ReadCell<string>(i, j++);
                p.Country = _excel.ReadCell<string>(i, j++);
                p.Addres = _excel.ReadCell<string>(i, j++);
                p.AddresRegistry = _excel.ReadCell<string>(i, j++);
                p.SubjectRF = _excel.ReadCell<string>(i, j++);
                p.VK = _excel.ReadCell<string>(i, j++);
                p.Health = _excel.ReadCell<string>(i, j++);
                p.HealtHron = _excel.ReadCell<string>(i, j++);
                p.VUZ = _excel.ReadCell<string>(i, j++);
                p.VUZKor = _excel.ReadCell<string>(i, j++);
                p.Specialnost = _excel.ReadCell<string>(i, j++);
                p.Diplom = _excel.ReadCell<string>(i, j++);
                p.SrBall = _excel.ReadCellDouble(i, j++);
                p.VKR = _excel.ReadCell<string>(i, j++);
                p.Soiskatelstvo = _excel.ReadCell<string>(i, j++) == "Да" ? 5 : 0;
                p.Exams = _excel.ReadCell<string>(i, j++);
                p.Statiy[0] = _excel.ReadCell<string>(i, j++) == "Есть" ? 5 : 0;
                p.Statiy[1] = _excel.ReadCell<string>(i, j++) == "Есть" ? 4 : 0;
                p.Statiy[2] = _excel.ReadCell<string>(i, j++) == "Есть" ? 3 : 0;
                p.Statiy[3] = _excel.ReadCell<string>(i, j++) == "Есть" ? 1 : 0;
                p.Statiy[4] = _excel.ReadCell<string>(i, j++) == "Есть" ? 1 : 0;
                p.Statiy[5] = _excel.ReadCell<string>(i, j++) == "Есть" ? 0.5 : 0;
                p.Statya = _excel.ReadCell<string>(i, j++);
                p.Sience[0] = _excel.ReadCell<string>(i, j++) == "Есть" ? 4 : 0;
                p.Sience[1] = _excel.ReadCell<string>(i, j++) == "Есть" ? 4 : 0;
                p.Sience[2] = _excel.ReadCell<string>(i, j++) == "Есть" ? 3 : 0;
                p.Sience[3] = _excel.ReadCell<string>(i, j++) == "Есть" ? 3 : 0;
                p.Sience[4] = _excel.ReadCell<string>(i, j++) == "Есть" ? 3 : 0;
                p.Sience[5] = _excel.ReadCell<string>(i, j++) == "Есть" ? 2 : 0;
                p.Sience[6] = _excel.ReadCell<string>(i, j++) == "Есть" ? 1 : 0;
                p.SienceName = _excel.ReadCell<string>(i, j++);
                p.SienceStepen[0] = _excel.ReadCell<string>(i, j++) == "Есть" ? 3 : 0;
                p.SienceStepen[1] = _excel.ReadCell<string>(i, j++) == "Есть" ? 6 : 0;
                p.SienceStepen[2] = _excel.ReadCell<string>(i, j++) == "Есть" ? 8 : 0;
                p.Work[0] = _excel.ReadCell<string>(i, j++) == "Есть" ? 2 : 0;
                p.Work[1] = _excel.ReadCell<string>(i, j++) == "Есть" ? 4 : 0;
                p.Work[2] = _excel.ReadCell<string>(i, j++) == "Есть" ? 6 : 0;
                p.WorkName = _excel.ReadCell<string>(i, j++);
                p.Sport[0] = _excel.ReadCell<string>(i, j++) == "Есть" ? 4 : 0;
                p.Sport[1] = _excel.ReadCell<string>(i, j++) == "Есть" ? 2 : 0;
                p.SportName = _excel.ReadCell<string>(i, j++);
                p.Language = _excel.ReadCell<string>(i, j++);
                p.Products = _excel.ReadCell<string>(i, j++);
                p.Napravlenie = _excel.ReadCell<string>(i, j++);
                p.Dopusk = _excel.ReadCell<string>(i, j++);
                p.TattooAndPirsing = _excel.ReadCell<string>(i, j++);
                p.Rost = _excel.ReadCellDouble(i, j++);
                p.Ves = _excel.ReadCellDouble(i, j++);
                p.Family = _excel.ReadCell<string>(i, j++);
                p.Info = _excel.ReadCell<string>(i, j++);

                p.Prioritet[0] = _excel.ReadCell<string>(i, 100) == "Есть" ? 3 : 0;
                p.Prioritet[1] = _excel.ReadCell<string>(i, 101) == "Есть" ? 1 : 0;
                p.Reserv = _excel.ReadCell<string>(i, 102) == "Резерв";
                p.Magistr = _excel.ReadCell<string>(i, 103) == "Магистр";
                p.Bakalavr = _excel.ReadCell<string>(i, 103) == "Бакаларв";
                p.Zachetka = _excel.ReadCell<string>(i, 103) == "Зачетка";
                p.Ball = p.ExecuteBall();
                listBox1.Items.Add(p);
               /* if (!string.IsNullOrWhiteSpace(p.URLPhoto))
                {
                    string id = p.URLPhoto.Split('=')[1];
                    DriveService service = AuthenticateOauth(idClient, "nr7vas");
                    Google.Apis.Drive.v2.Data.File file = new Google.Apis.Drive.v2.Data.File();
                    file.Id = id;
                    if (!Directory.Exists("images"))
                        Directory.CreateDirectory("images");
                    if (!Directory.Exists("images/photos"))
                        Directory.CreateDirectory("images/photos");
                    if (!Directory.Exists("images/soglasie"))
                        Directory.CreateDirectory("images/soglasie");
                    string filenamePhoto =
                        $"images/photos/{p.SurName}_{p.Name}_{p.Patronomyc}.png";
                    string filenameSoglasie =
                        $"images/soglasie/{p.SurName}_{p.Name}_{p.Patronomyc}_СОГЛАСИЕ.pdf";
                    DownloadFile(service, file, filenamePhoto);

                    using (Image image = Image.FromFile(filenamePhoto))
                    {
                        using (MemoryStream m = new MemoryStream())
                        {
                            image.Save(m, image.RawFormat);
                            byte[] imageBytes = m.ToArray();
                            
                            p.base64PhotoString = Convert.ToBase64String(imageBytes);
                            p.photoPath = filenamePhoto;
                        }
                    }

                    string idSoglasie = p.URLSOGLASIE.Split('=')[1];
                    file = new Google.Apis.Drive.v2.Data.File();
                    file.Id = idSoglasie;
                    DownloadFile(service, file, filenameSoglasie);
                    using (Image image = Image.FromFile(filenameSoglasie))
                    {
                        using (MemoryStream m = new MemoryStream())
                        {
                            image.Save(m, image.RawFormat);
                            byte[] imageBytes = m.ToArray();

                            // Convert byte[] to Base64 String
                            p.base64SoglasieString = Convert.ToBase64String(imageBytes);
                       }
                    }
                }*/
                Persons.Add(p);
            }

            Persons.Sort(new NaturalStringComparer());
            Persons.Sort(new BallComparer());
            listBox1.Items.Clear();
            foreach (var person in Persons)
            {
                listBox1.Items.Add(person);
            }

            _excel.FinishAccess();
        }

        /// <summary>
        /// Кнопка напечатать ведомость
        /// </summary>
        private void ведомостьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if(Persons.Count != 0)
            Tmp(() => {
                word.VedomostStart(Persons);
                word.ToVisible();
                word.Close();
            });
            else
            {
                MessageBox.Show("Список пуст", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Кнопка напечатать список рейтинговый
        /// </summary>
        private void списокToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Persons.Count != 0)
            {
                Persons.Sort(new NaturalStringComparer());
                Persons.Sort(new BallComparer());

                Tmp(() =>
                {
                    word.RaitingStart(Persons, Max);
                    word.ToVisible();
                    word.Close();
                });
            }
            else
            {
                MessageBox.Show("Список пуст", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Кнопка напечатать список
        /// </summary>
        private void списокToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (Persons.Count != 0)
            {
                Persons.Sort(new NaturalStringComparer());
                Persons.Sort(new BallComparer());

                Tmp(() =>
                {
                    word.SpisokStart(Persons);
                    word.ToVisible();
                    word.Close();
                });
            }
            else
            {
                MessageBox.Show("Список пуст", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void листСобеседованияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Persons.Count != 0)
            {
                Tmp(() =>
                {
                    word.ListStart(Persons);
                    System.Diagnostics.Process.Start("explorer.exe", Pathes.TempPath);
                    word.Close();
                });
            }
            else
            {
                MessageBox.Show("Список пуст", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        static bool IsNum(string s)
        {
            if (s.Count(x => x == ',') > 1)
                return false;
            else
            {
                if (s.Count(x => x == ',') == 1)
                {
                    //double 
                    foreach (char c in s)
                    {
                        if (c != ',')
                        if (!IsDigit(c)) return false;
                    }
                }
                else
                {
                    foreach (char c in s)
                    {
                        if (c == ',')
                            if (!IsDigit(c)) return false;
                    }
                }
            }
           
            return true;
        }

        private void CancelBTN_Click(object sender, EventArgs e)
        {
            groupBox1.Enabled = false;
            groupBox2.Enabled = false;
            listBox1.Enabled = true;
            changeBTN.Enabled = true;

            Persons.Sort(new NaturalStringComparer());
            Persons.Sort(new BallComparer());
            listBox1.Items.Clear();
            foreach (var person in Persons)
            {
                listBox1.Items.Add(person);
            }

            listBox1.SelectedItem = listBox1.Items.OfType<Person>().First(x => x.ID == SelectedPerson.ID);
        }

        private void файлToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = false,
                Filter = "Excel files (*.xlsx)|*.xlsx",
                InitialDirectory = "c:\\",
                FilterIndex = 1,
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                PathFile = openFileDialog.FileName;
                ReadExcel();
                this.Text = $"Подбор кандидатов ({PathFile})";
            }
        }

        private void тестовыйФайлToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PathFile = "Шаблон анкеты (Ответы).xlsx";
            ReadExcel();
            this.Text = $"Подбор кандидатов ({PathFile})";
        }

        private void сохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            string filename = saveFileDialog1.FileName;

            _excel = new AccessExcel();
            _excel.DoAccess(filename);

            groupBox1.Enabled = false;
            groupBox2.Enabled = false;
            listBox1.Enabled = true;
            changeBTN.Enabled = true;

            Persons.Sort(new NaturalStringComparer());
            Persons.Sort(new BallComparer());
            listBox1.Items.Clear();
            foreach (var person in Persons)
            {
                listBox1.Items.Add(person);
            }

            listBox1.SelectedItem = listBox1.Items.OfType<Person>().First(x => x.ID == SelectedPerson.ID);

            for (int i = 0; i < Persons.Count; i++)
            {
                _excel.WriteCell<string>(1, 1, "Дата изменения");
                int iRow = i + 2;
                var p = Persons[i];
                int j = 1;
                _excel.WriteCell<string>(iRow, j++, DateTime.Now.ToString("G"));
                _excel.WriteCell<string>(iRow, j++, p.email);
                _excel.WriteCell<string>(iRow, j++, p.base64PhotoString);
                _excel.WriteCell<string>(iRow, j++, p.base64SoglasieString);
                _excel.WriteCell<string>(iRow, j++, p.SurName);
                _excel.WriteCell<string>(iRow, j++, p.Name);
                _excel.WriteCell<string>(iRow, j++, p.Patronomyc);
                _excel.WriteCell<string>(iRow, j++, p.Telefon);
                _excel.WriteCell<string>(iRow, j++, p.Birthday.ToShortDateString());
                _excel.WriteCell<string>(iRow, j++, p.BirthdayMesto);
                _excel.WriteCell<string>(iRow, j++, p.Country);
                _excel.WriteCell<string>(iRow, j++, p.Addres);
                _excel.WriteCell<string>(iRow, j++, p.AddresRegistry);
                _excel.WriteCell<string>(iRow, j++, p.SubjectRF);
                _excel.WriteCell<string>(iRow, j++, p.VK);
                _excel.WriteCell<string>(iRow, j++, p.Health);
                _excel.WriteCell<string>(iRow, j++, p.HealtHron);
                _excel.WriteCell<string>(iRow, j++, p.VUZ);
                _excel.WriteCell<string>(iRow, j++, p.VUZKor);
                _excel.WriteCell<string>(iRow, j++, p.Specialnost);
                _excel.WriteCell<string>(iRow, j++, p.Diplom);
                _excel.WriteCell<double>(iRow, j++, p.SrBall);
                _excel.WriteCell<string>(iRow, j++, p.VKR);
                _excel.WriteCell<string>(iRow, j++, p.Soiskatelstvo == 0 ? "Нет" : "Да");
                _excel.WriteCell<string>(iRow, j++, p.Exams);
                _excel.WriteCell<string>(iRow, j++, p.Statiy[0] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Statiy[1] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Statiy[2] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Statiy[3] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Statiy[4] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Statiy[5] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Statya);
                _excel.WriteCell<string>(iRow, j++, p.Sience[0] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Sience[1] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Sience[2] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Sience[3] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Sience[4] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Sience[5] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Sience[6] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.SienceName);
                _excel.WriteCell<string>(iRow, j++, p.SienceStepen[0] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.SienceStepen[1] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.SienceStepen[2] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Work[0] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Work[1] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Work[2] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.WorkName);
                _excel.WriteCell<string>(iRow, j++, p.Sport[0] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.Sport[1] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, j++, p.SportName);
                _excel.WriteCell<string>(iRow, j++, p.Language);
                _excel.WriteCell<string>(iRow, j++, p.Products);
                _excel.WriteCell<string>(iRow, j++, p.Napravlenie);
                _excel.WriteCell<string>(iRow, j++, p.Dopusk);
                _excel.WriteCell<string>(iRow, j++, p.TattooAndPirsing);
                _excel.WriteCell<double>(iRow, j++, p.Rost);
                _excel.WriteCell<double>(iRow, j++, p.Ves);
                _excel.WriteCell<string>(iRow, j++, p.Family);
                _excel.WriteCell<string>(iRow, j++, p.Info);
                _excel.WriteCell<string>(iRow, 100, p.Prioritet[0] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, 101, p.Prioritet[1] == 0 ? "Нет" : "Есть");
                _excel.WriteCell<string>(iRow, 102, p.Reserv ? "Резерв" : "Основа");
                _excel.WriteCell<string>(iRow, 103, p.Magistr ? "Магистр" : p.Bakalavr ? "Бакалавр" : p.Zachetka ? "Зачетка" : "");
            }
            _excel.FinishAccess();
        }
    }

    [SuppressUnmanagedCodeSecurity]
    internal static class NativeMethods
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        public static extern int StrCmpLogicalW(string psz1, string psz2);
    }

    public sealed class NaturalStringComparer : IComparer<Person>
    {
        public int Compare(Person a, Person b)
        {
            return NativeMethods.StrCmpLogicalW(a.ToString(), b.ToString());
        }
    }

    public sealed class BallComparer : IComparer<Person>
    {
        public int Compare(Person a, Person b)
        {
            if (a.Ball < b.Ball)
                return 1;
            if (a.Ball == b.Ball)
            {
                if (a.SrBall > b.SrBall)
                {
                    return -1;
                }
                else
                {
                    return 1;
                }
            }

            return -1;
        }
    }

    
}

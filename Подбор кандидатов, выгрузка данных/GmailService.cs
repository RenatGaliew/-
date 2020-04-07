using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Download;
using Google.Apis.Drive.v2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using MimeKit;
using SendRequest = Google.Apis.Gmail.v1.UsersResource.MessagesResource.SendRequest;

namespace Подбор_кандидатов__выгрузка_данных
{
    public class MailService
    {
        private static string EMAIL_NR = "nr7vas@gmail.com";
        private static string code = "qrwhu_5RKGUpBuc6joycjiO6";
        private static string idClient = "client_secret_572291057451-2mai9prgom68ufefoj5p0709sqji23s1.apps.googleusercontent.com.json";
        private static string idClientMail = "savvy-eye-272217-dc7d321d0f67.json";
        private static string ServiceMail = "nr7vas@savvy-eye-272217.iam.gserviceaccount.com";
        private static string username = "nr7vas";
        private static string usernameEmail = "nr7vas@gmail.com";

        private static string FileSogl = $"{Pathes.DirectoryPDFS}/СОГЛАСИЕ.pdf";

        public DriveService DriveService { get; set; }
        public SheetsService SheetService { get; set; }
        public GmailService GMailService { get; set; }

        public bool OkServices => DriveService != null && SheetService != null && GMailService != null;

        public MailService()
        {
            Tmp(() =>
            {
                var auth = AuthenticateOauth(idClient, username);
                DriveService = auth.Item1;
                SheetService = auth.Item2;
                GMailService = AuthenticateMail(idClient);
            });

            
        }

        private static Tuple<DriveService, SheetsService> AuthenticateOauth(string clientSecretJson, string userName)
        {
            try
            {
                if (String.IsNullOrEmpty(userName))
                    throw new ArgumentNullException("userName");
                if (String.IsNullOrEmpty(clientSecretJson))
                    throw new ArgumentNullException("clientSecretJson");
                if (!File.Exists(clientSecretJson))
                    throw new Exception("clientSecretJson file does not exist.");

                // These are the scopes of permissions you need. It is best to request only what you need and not all of them
                string[] scopes = { DriveService.Scope.DriveReadonly };//View the files in your Google Drive                                                 


                UserCredential credential;

                using (var stream = new FileStream(clientSecretJson, FileMode.Open, FileAccess.Read))
                {
                    string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".credentials/", Assembly.GetExecutingAssembly().GetName().Name);

                    // Requesting Authentication or loading previously stored authentication for userName
                    var secret = GoogleClientSecrets.Load(stream).Secrets;
                    secret.ClientSecret = code;
                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(secret,
                                                                             scopes,
                                                                             userName,
                                                                             CancellationToken.None,
                                                                             new FileDataStore(credPath, true)).Result;
                }

                var drive = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Drive Oauth2 Authentication Sample"
                });
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Drive Oauth2 Authentication Sample"
                });
                return new Tuple<DriveService, SheetsService>(drive, service);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Create Oauth2 account DriveService failed" + ex.Message);
                throw new Exception("CreateServiceAccountDriveFailed", ex);
            }
        }

        private static GmailService AuthenticateMail(string clientSecretJson)
        {
            try
            {
                string[] scopes = { GmailService.Scope.GmailSend };

                UserCredential credential;

                using (var stream = new FileStream(clientSecretJson, FileMode.Open, FileAccess.Read))
                {
                    string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".credentials/", Assembly.GetExecutingAssembly().GetName().Name);

                    var secret = GoogleClientSecrets.Load(stream).Secrets;
                    secret.ClientSecret = code;

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        secret,
                        scopes,
                        EMAIL_NR,
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                }
                var service = new GmailService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Drive Oauth2 Authentication Sample",
                });
                return service;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Create Oauth2 account DriveService failed" + ex.Message);
                throw new Exception("CreateServiceAccountDriveFailed", ex);
            }
        }

        public bool SendMail(Person person)
        {
            var builder = new BodyBuilder();
            var orden1 = builder.LinkedResources.Add("message/orden1.png");
            orden1.ContentId = "orden1";
            orden1.IsAttachment = false;
            var orden2 = builder.LinkedResources.Add("message/orden2.png");
            orden2.ContentId = "orden2";
            orden2.IsAttachment = false;
            var orden3 = builder.LinkedResources.Add("message/orden3.png");
            orden3.ContentId = "orden3";
            orden3.IsAttachment = false;
            var emblema = builder.LinkedResources.Add("message/emblema.png");
            emblema.ContentId = "emblema";
            emblema.IsAttachment = false;

            builder.HtmlBody = File.ReadAllText("../../body.html", Encoding.UTF8);
            builder.Attachments.Add(person.PDFPath);
            builder.Attachments.Add(FileSogl);
            var messageMime = new MimeMessage { Body = builder.ToMessageBody() };
            messageMime.From.Add(new MailboxAddress("Научная рота ВАС", EMAIL_NR));
            messageMime.To.Add(new MailboxAddress(person.Name, person.email));
            messageMime.Subject = "Заявка в научную роту ВАС";

            string tmp1 = messageMime.ToString();
            tmp1 = tmp1.Replace("&FIO", $"Уважаемый {person.Name} {person.Patronomyc}!");
            tmp1 = tmp1.Replace("&TEXT1", @"Поздравляем Вас с прохождением первого этапа отборочной 
комиссии для прохождения военной службы по призыву в Военной академи связи имени Маршала Советского Союза 
С.М. Будённого города Санкт-Петербург!");
            tmp1 = tmp1.Replace("&TEXT2", @"Для продолжения участия в отборе просим Вас перейти по ссылке 
и заполнить вторую форму, касающуюся Ваших достижений в научной деятельности:");
            tmp1 = tmp1.Replace("&btnTEXT", @"Подтверждающие документы призывника");
            tmp1 = tmp1.Replace("&TEXT3",
                @"Кроме того, просим Вас ознакомиться с документами прикрепленными к данному письму.");
            tmp1 = tmp1.Replace("&TEXT4", @"С уважением,");
            tmp1 = tmp1.Replace("&TEXT5", @"Члены отборочной комиссии Военной академии связи!");

            var gmailMessage = new Message
            {
                Raw = Encode(tmp1)
            };

            SendRequest request = GMailService.Users.Messages.Send(gmailMessage, EMAIL_NR);
            try
            {
                request.Execute();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string Encode(string text)
        {
            byte[] bytesUTF8 = Encoding.UTF8.GetBytes(text);

            return Convert.ToBase64String(bytesUTF8)
                .Replace('+', '-')
                .Replace('/', '_')
                .Replace("=", "");
        }

        public void DownloadFile(string fileToDownload, string saveTo)
        {
            var tmp = fileToDownload.Split('=');
            string idFile = "";
            if (tmp.Length == 2)
            {
                idFile = tmp[1];
            }

            if (File.Exists(saveTo))
                return;
            var request = DriveService.Files.Get(idFile);
            var stream = new MemoryStream();

            // Add a handler which will be notified on progress changes.
            // It will notify on each chunk download and when the
            // download is completed or failed.
            request.MediaDownloader.ProgressChanged += (IDownloadProgress progress) =>
            {
                switch (progress.Status)
                {
                    case DownloadStatus.Downloading:
                    {
                        Console.WriteLine(progress.BytesDownloaded);
                        break;
                    }
                    case DownloadStatus.Completed:
                    {
                        Console.WriteLine("Download complete.");
                        SaveStream(stream, saveTo);
                        break;
                    }
                    case DownloadStatus.Failed:
                    {
                        Console.WriteLine("Download failed.");
                        break;
                    }
                }
            };
            request.Download(stream);

        }

        private static void SaveStream(System.IO.MemoryStream stream, string saveTo)
        {
            using (System.IO.FileStream file = new System.IO.FileStream(saveTo, System.IO.FileMode.Create, System.IO.FileAccess.Write))
            {
                stream.WriteTo(file);
            }
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
    }
}

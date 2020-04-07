using System;
using System.Threading;

namespace Подбор_кандидатов__выгрузка_данных
{
    public enum InternetStatus
    {
        YesInternet,
        Searching,
        SearchingAgain,
        NoInternet
    }
    public class InternetChecker
    {
        private InternetStatus _status;
        public event EventHandler<InternetStatus> StatusChanging;

        public InternetStatus Status
        {
            get => _status;
            set
            {
                _status = value;
                StatusChanging?.Invoke(this, value);
            }
        }

        public void Start()
        {
            while (true)
            {
                if (ConnectivityChecker.CheckInternet() != ConnectivityChecker.ConnectionStatus.Connected)
                {
                    Status = InternetStatus.SearchingAgain;
                    Thread.Sleep(5000);
                    Status = InternetStatus.Searching;
                }
                else
                    Status = InternetStatus.YesInternet;

                if (Status != InternetStatus.YesInternet)
                    Status = InternetStatus.NoInternet;
                Thread.Sleep(5000);
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;


namespace EReport
{
    class ViewModel : INotifyPropertyChanged
    {
        String[] To_address = new String[50];
        String[] CC_address = new String[100];
        String[] Bcc_address = new String[20];
        DBWork DBW = new DBWork();

        private String _logviewer = "";

        double IDDincoming_Diff_actual = 0;
        double IDDincoming_acceptance_diff = 0.3; //%      
        public string IDDAcceptanceString = "";

        public double General_acceptance_diff = 1.0; //%  
        public bool ShouldEmailSendorNot = true;

        public String User_EmailID = "";
        public String Password_String = "";

        public String SMTP_Client = "";
        public int SMTP_Port = 587;

        List<String> Filename = new List<string>();

        public String Destination_Excel_Path = "D:\\Excel";

        public int SubtractiveDataDay = 1;


        public String To_String = "";
        public String CC_String = "";
        public String BCC_String = "";
        public String MailSubject = "";
        public String MailBody = "";


        public String LogViewer
        {
            get { return _logviewer; }
            set
            {
                _logviewer = value;
                // Call OnPropertyChanged whenever the property is updated
                OnPropertyChanged("LogViewer");
            }
        }

        public ViewModel()
        {
            this.PropertyChanged += ViewModel_PropertyChanged;
            DBW.PropertyChanged += DBW_PropertyChanged;

            if (!Directory.Exists(Destination_Excel_Path))
            {
                Directory.CreateDirectory(Destination_Excel_Path);
            }
        }

        private void DBW_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "LogViewer")
            {
                this.LogViewer = DBW.LogViewer;
            }
            if (e.PropertyName == "FileLogger")
            {
                Write_logFile(DBW.FileLogger);
            }
        }

        private void ViewModel_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {

        }



        protected void OnPropertyChanged(string data)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(data));
            }
        }


        public async void TaskHandle(String _generator, bool error_checkOrNot)
        {
            if (_generator == "AUTO" && error_checkOrNot)
            {
                try
                {
                    if (IDDAcceptanceString == "")
                    {
                        IDDincoming_acceptance_diff = 0.3;
                    }
                    else
                    {
                        IDDincoming_acceptance_diff = Convert.ToDouble(IDDAcceptanceString);
                    }
                }
                catch (Exception ex)
                {
                    LogViewer = ex.Message;
                    IDDincoming_acceptance_diff = 0.3;
                }


                IDDincoming_Diff_actual = await IDDIncomingDifferenceCheck();
                if (IDDincoming_Diff_actual <= IDDincoming_acceptance_diff)
                {
                    CallAsyncTasks();
                }
                else
                {
                    LogViewer = "Event triggered, but traffic data is abnormal. Suspended email sending. Difference is " + IDDincoming_Diff_actual.ToString("#0.00") + "%, which is greater than " + IDDincoming_acceptance_diff.ToString("#0.00") + "%.";
                    Write_logFile(LogViewer);
                }
            }
            else
            {
                CallAsyncTasks();
            }
        }

        private async void CallAsyncTasks()
        {
            string YYYY = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).Year.ToString();
            string yy = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("yy");
            string MMMM = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("MMMM");
            string MMM = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("MMM");
            string dd = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("dd");

            string _dir = Destination_Excel_Path + "\\" + YYYY + "\\" + MMMM + "\\" + dd + "-" + MMM + "-" + yy;
            if (!Directory.Exists(_dir))
            {
                Directory.CreateDirectory(_dir);
            }

            LogViewer = "Accessing remote database, please wait.... ... .. .";
            List<Task<String>> tasklist = new List<Task<String>>();
            //tasklist.Add(TaskHandleAsyncIDDIncoming(_dir));
            //tasklist.Add(TaskHandleAsyncIDDOutgoing(_dir));
            //tasklist.Add(TaskHandleAsyncANS(_dir));
            tasklist.Add(TaskHandleAsyncICX(_dir));
            string[] _filename = await Task.WhenAll(tasklist);
            
            Filename.Clear();
            for (int i = 0; i < tasklist.Count; i++)
            {
                Filename.Add(_filename[i]);
            }
            
            tasklist.Clear();

            await TryToMailAsync();
        }

        private Task TryToMailAsync()
        {
            return Task.Run(() => TryToMail());
        }

        public bool MailingProgess = false;

        private async Task TryToMail()
        {
            bool replymail = false;
            MailingProgess = true;
            int count = 0;
            while (replymail == false)
            {
                if (!ShouldEmailSendorNot)
                {
                    LogViewer = "Email suspended and auto sending has been stopped.";
                    MailingProgess = false;
                    break;
                }

                replymail = await MailReportAsync();

                if(replymail)
                {
                    MailingProgess = false;
                }
                count++;
                if (replymail == false && ShouldEmailSendorNot)
                {
                    if (count == 5)
                    {
                        LogViewer = "Message sending failed, retried 5 times.";
                        Write_logFile(LogViewer);
                        MessageBox.Show("[EReport]: Message sending failed, retried 5 times.", "EReport", MessageBoxButton.OK, MessageBoxImage.Error);
                        MailingProgess = false;
                        break;
                    }
                    Thread.Sleep(60000); //1 min
                }
            }
        }

        private Task<bool> MailReportAsync()
        {
            LogViewer = "Sending mail, please wait.... ... .. .";
            return Task.Run(() => MailReport());
        }

        private bool MailReport()
        {
            bool reply;
            MailMessage mail = new MailMessage();

            System.Net.Mail.Attachment _attachment = null;

            try
            {
                To_address = To_String.Split(',');
                for (int i = 0; i < To_address.Length; i++)
                {
                    if (To_address[i] != "")
                        mail.To.Add(To_address[i]);
                }


                if (CC_String != "")
                {
                    CC_address = CC_String.Split(',');
                    for (int i = 0; i < CC_address.Length; i++)
                    {
                        if (CC_address[i] != "")
                            mail.CC.Add(CC_address[i]);
                    }
                }

                if (BCC_String != "")
                {
                    Bcc_address = BCC_String.Split(',');
                    for (int i = 0; i < Bcc_address.Length; i++)
                    {
                        if (Bcc_address[i] != "")
                            mail.Bcc.Add(Bcc_address[i]);
                    }
                }

                String _mailbody = "";

                _mailbody = "Traffic date: " + DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToShortDateString() + Environment.NewLine + Environment.NewLine;

                _mailbody += MailBody; // body from window

                _mailbody += "\n\n" + Signature;

                mail.Body = _mailbody;

                if (MailSubject == "")
                {
                    mail.Subject = "Traffic date: " + DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToShortDateString();
                }
                else
                {
                    mail.Subject = MailSubject + ", Traffic date: " + DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToShortDateString();
                }

                mail.From = new MailAddress(User_EmailID);




                if (Filename.Count > 0)
                {
                    for (int i = 0; i < Filename.Count; i++)
                    {
                        _attachment = new System.Net.Mail.Attachment(Filename[i]);
                        mail.Attachments.Add(_attachment);
                    }
                }


                var client = new SmtpClient(SMTP_Client, SMTP_Port)
                {
                    Credentials = new NetworkCredential(User_EmailID, Password_String),
                    EnableSsl = true

                };

                client.Send(mail);

                LogViewer = "Mail successfully sent for date: " + DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToShortDateString() + ".";
                Write_logFile(LogViewer);
                reply = true;
            }
            catch (Exception ex)
            {
                LogViewer = ex.Message;
                Write_logFile(ex.Message);
                reply = false;
            }

            finally
            {
                if (_attachment != null)
                    _attachment.Dispose();

                if (mail != null)
                    mail.Dispose();
            }

            return reply;
        }

        public string Signature = "[It is autogenerated mail from Auto Report Generator application. Please do not reply with this mail.]\n\n{Powered by: Md. Rakib Subaid\nManager, Billing System\nIT & Billing, BTCL\nSher-E-Bangla Nagar, Dhaka}";

        private Task<String> TaskHandleAsyncIDDIncoming(string _dir)
        {
            return Task.Run(() =>
            {
                String _file = "";
                while (_file == "")
                {
                    _file = DBW.QueryDatabaseforIDDIncoming(SubtractiveDataDay, _dir, General_acceptance_diff);
                }
                return _file;
            });
        }

        private Task <String> TaskHandleAsyncIDDOutgoing(string _dir)
        {
            return Task.Run(() =>
            {
                String _file = "";
                while (_file == "")
                {
                    _file = DBW.QueryDatabaseforIDDOutgoing(SubtractiveDataDay, _dir, General_acceptance_diff);
                }
                return _file;
            });
        }

        //private Task<String> TaskHandleAsyncANSIncoming(string _dir)
        //{
        //    return Task.Run(() =>
        //    {
        //        String _file = "";
        //        while (_file == "")
        //        {
        //            _file = DBW.QueryDatabaseforANSLocalIncoming(SubtractiveDataDay, _dir, General_acceptance_diff);
        //        }
        //        return _file;
        //    });
        //}

        private Task<String> TaskHandleAsyncANS(string _dir)
        {
            return Task.Run(() =>
            {
                String _file = "";
                while (_file == "")
                {
                    _file = DBW.QueryDatabaseforANSLocal(SubtractiveDataDay, _dir, General_acceptance_diff);
                }
                return _file;
            });
        }

        private Task<String> TaskHandleAsyncICX(string _dir)
        {
            return Task.Run(() =>
            {
                String _file = "";
                while (_file == "")
                {
                    _file = DBW.QueryDatabaseforICX(SubtractiveDataDay, _dir);
                }
                return _file;
            });
        }


        private Object logFileLock = new Object();
        public void Write_logFile(String str)///////////////////////////////////////////////////////////////////////////////////////////////////
        {
            try
            {
                lock (logFileLock)
                {
                    // Write the string to a file.
                    //Logfile = new System.IO.StreamWriter(@"c:\\Users\\Public\\Echo_Log.txt", true);
                    System.IO.StreamWriter Logfile = new System.IO.StreamWriter(@"c:\\Users\\Public\\EReport_" + DateTime.Now.Year.ToString() + ".log", true);
                    Logfile.WriteLine(DateTime.Now.ToString() + ":- " + str);
                    Logfile.Close();
                }
            }
            catch (Exception ex)
            {
                this.LogViewer = ex.Message + " <" + ex.GetType().ToString() + ">";
            }
        }

        public Task<Double> IDDIncomingDifferenceCheck()
        {
            return Task.Run(() => DBW.IDDIncomingDifferenceCheck(SubtractiveDataDay));
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }
}
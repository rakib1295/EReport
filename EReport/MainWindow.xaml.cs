using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Deployment.Application;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Media;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace EReport
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        String EventTime = "";        
        bool error_checkOrNot = true;

        ViewModel VM = new ViewModel();

        System.Windows.Forms.NotifyIcon IconInstance = new System.Windows.Forms.NotifyIcon();
        public MainWindow()
        {
            InitializeComponent();

            timerforPopup.Interval = TimeSpan.FromSeconds(5);
            timerforPopup.Tick += timer_TickForPopup;

            DispatcherTimerClock();

            Stream iconStream = Application.GetResourceStream(new Uri("pack://application:,,,/Images/image_0Iw_icon.ico")).Stream;
            IconInstance.Icon = new System.Drawing.Icon(iconStream);

            VM.PropertyChanged += View_PropertyChanged;
            IconInstance.Visible = true;
            this.Closing += MainWindow_Closing;
            Application.Current.MainWindow.Loaded += MainWindow_Loaded;

            IconInstance.Text = "EReport";
            SignatureBody.Text = VM.Signature;

            IconInstance.DoubleClick +=
                delegate (object sender, EventArgs args)
                {
                    this.Show();
                    this.Activate();
                    this.WindowState = WindowState.Maximized;
                };

#if !DEBUG
            versionNumber.Text = "Version: " + ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4);
#endif
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            if (MessageBox.Show("Do you want to close the application?", "EReport", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.No)
            {
                e.Cancel = true;
                
            }
            else
            {
                Properties.Settings.Default.Subject = Sub.Text;
                Properties.Settings.Default.To = To.Text;
                Properties.Settings.Default.CC = CC.Text;
                Properties.Settings.Default.Bcc = Bcc.Text;
                Properties.Settings.Default.Body = Body.Text;
                Properties.Settings.Default.ActionTime = Action_time_textbox.Text;
                Properties.Settings.Default.Email = user_email.Text;
                Properties.Settings.Default.Password = acc_psw.Password;
                Properties.Settings.Default.SMTPClient = SMTP_Client.Text;
                Properties.Settings.Default.SMTPPort = SMTP_Port.Text;
                Properties.Settings.Default.CheckError = (bool)IDD_Diff_Checkbox.IsChecked;
                Properties.Settings.Default.IDDInErrorLimit = IDD_in_Error_percentage_textbox.Text;
                Properties.Settings.Default.GeneralErrorLimit = General_percentage_textbox.Text;

                Properties.Settings.Default.Save();
                IconInstance.Dispose();
            }
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            Sub.Text = Properties.Settings.Default.Subject;
            To.Text = Properties.Settings.Default.To;
            CC.Text = Properties.Settings.Default.CC;
            Bcc.Text = Properties.Settings.Default.Bcc;
            Body.Text = Properties.Settings.Default.Body;
            Action_time_textbox.Text = Properties.Settings.Default.ActionTime;
            user_email.Text = Properties.Settings.Default.Email;
            acc_psw.Password = Properties.Settings.Default.Password;
            SMTP_Client.Text = Properties.Settings.Default.SMTPClient;
            SMTP_Port.Text = Properties.Settings.Default.SMTPPort;
            IDD_Diff_Checkbox.IsChecked = Properties.Settings.Default.CheckError;
            IDD_in_Error_percentage_textbox.Text = Properties.Settings.Default.IDDInErrorLimit;
            General_percentage_textbox.Text = Properties.Settings.Default.GeneralErrorLimit;
        }

        protected override void OnStateChanged(EventArgs e)
        {
            if (WindowState == System.Windows.WindowState.Minimized)
            {
                this.Hide();
            }

            base.OnStateChanged(e);
        }

        private void View_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "LogViewer")
            {
                Show_LogTextblock(VM.LogViewer);
            }
        }

        private Object thisLock = new Object();

        void Show_LogTextblock(String str)
        {
            try
            {
                lock (thisLock)
                {

                    Dispatcher.BeginInvoke((Action)(() =>
                    {
                        log_textblock.Text = log_textblock.Text + "# " + DateTime.Now.ToString() + ":- " + str + "\n";
                        _scrollbar_log.ScrollToBottom();
                    }));
                }
            }
            catch (Exception ex)
            {
                VM.Write_logFile(ex.Message + " <" + ex.GetType().ToString() + ">");
            }
        }


        void DispatcherTimerClock()
        {
            DispatcherTimer timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(1);
            timer.Tick += timer_Tick;
            timer.Start();
        }

        void timer_Tick(object sender, EventArgs e)
        {
            string CurrentTime;
            CurrentTime = DateTime.Now.ToLongTimeString();

            Dispatcher.BeginInvoke((Action)(() =>
            {
                clock_textblock.Text = CurrentTime; //time showing
            }));

            if (CurrentTime == "12:00:00 AM") //###################################################CONSIDER ALWAYS########################################################
            {
                //VM.SubtractiveDataDay = 1;
                //_date_picker.IsEnabled = false;
                //calender_btn.Content = "Enable Calender";
                //_date_picker.SelectedDate = DateTime.Today.Subtract(TimeSpan.FromDays(1));

                Dispatcher.BeginInvoke((Action)(() =>
                {
                    log_textblock.Text = "";
                }));

                //VM.Write_logFile("Successfully cleared previous day status.");
            }

            if (CurrentTime == EventTime && VM.ShouldEmailSendorNot == true) //time for event fire
            {
                _date_picker.IsEnabled = false;
                calender_btn.Content = "Enable Calender";
                _date_picker.SelectedDate = DateTime.Today.Subtract(TimeSpan.FromDays(1));
                VM.SubtractiveDataDay = 1;

                if (To.Text != "")
                {
                    VM.TaskHandle("AUTO", error_checkOrNot);
                }
                else
                {
                    Show_LogTextblock("Please select at least one recepient at 'To'");
                    MessageBox.Show("Please give at least one recepient at 'To'", "EReport", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        private void SelectFolder_function_Click_1(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog dlg = new System.Windows.Forms.FolderBrowserDialog();
            //Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension
            //dlg.DefaultExt = ".xls";
            //dlg.Filter = "Excel Worksheets|*.xls;*.xlsx";

            // Display OpenFileDialog by calling ShowDialog method
            System.Windows.Forms.DialogResult result = dlg.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                // Open document
                string pathname = dlg.SelectedPath;
                VM.Destination_Excel_Path = pathname;
                Show_LogTextblock("Excel files will be saved to: " + VM.Destination_Excel_Path);
                VM.Write_logFile("Excel files will be saved to: " + VM.Destination_Excel_Path);
            }
        }

        private void OpenFolder_function_Click(object sender, RoutedEventArgs e)
        {
            string _dir = VM.Destination_Excel_Path + "\\" + DateTime.Today.Subtract(TimeSpan.FromDays(VM.SubtractiveDataDay)).Year.ToString() + "\\" +
                    DateTime.Today.Subtract(TimeSpan.FromDays(VM.SubtractiveDataDay)).ToString("MMMM") + "\\" + DateTime.Today.Subtract(TimeSpan.FromDays(VM.SubtractiveDataDay)).ToShortDateString();
            if (!Directory.Exists(_dir))
            {
                Process.Start(VM.Destination_Excel_Path);
            }
            else
            {
                Process.Start(_dir);
            }            
        }

        private void exit_function_Click_1(object sender, RoutedEventArgs e)
        {
            Close();
        }



        private void Settings_function_Click_1(object sender, RoutedEventArgs e)
        {
            Popup_Settings.IsOpen = true;
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
                e.Handled = true;
            }
            catch (Exception ex)
            {
                Show_LogTextblock(ex.Message);
                VM.Write_logFile(ex.Message);
            }
        }

        private void _date_picker_SelectedDateChanged_1(object sender, SelectionChangedEventArgs e)
        {
            var diff = DateTime.Today.Subtract(_date_picker.SelectedDate.Value);
            VM.SubtractiveDataDay = diff.Days;
            if (VM.SubtractiveDataDay > 1)
                Show_LogTextblock("Selected date is " + VM.SubtractiveDataDay + " days before today.");
            else if (VM.SubtractiveDataDay == 0)
                Show_LogTextblock("Selected day is today.");
        }

        private void calender_btn_Click_1(object sender, RoutedEventArgs e)
        {
            if (_date_picker.IsEnabled == false)
            {
                _date_picker.IsEnabled = true;
                calender_btn.Content = "Disable Calender";
                Show_LogTextblock("Calender is enabled.");
                _date_picker.DisplayDateEnd = DateTime.Today;
            }
            else
            {
                _date_picker.IsEnabled = false;
                Show_LogTextblock("Calender is disabled.");
                calender_btn.Content = "Enable Calender";
                VM.SubtractiveDataDay = 1;
                Show_LogTextblock("Selected day is yesterday.");
                _date_picker.SelectedDate = DateTime.Today.Subtract(TimeSpan.FromDays(1));                
            }
        }


        private void To_TextChanged(object sender, TextChangedEventArgs e)
        {
            VM.To_String = To.Text;
        }


        private void CC_TextChanged(object sender, TextChangedEventArgs e)
        {
            VM.CC_String = CC.Text;
        }

        private void BCC_TextChanged(object sender, TextChangedEventArgs e)
        {
            VM.BCC_String = Bcc.Text;
        }

        private void Body_TextChanged(object sender, TextChangedEventArgs e)
        {
            VM.MailBody = Body.Text;
        }


        private void Sub_TextChanged(object sender, TextChangedEventArgs e)
        {
            VM.MailSubject = Sub.Text;
        }

        private void Alarm_TextBox_TextChanged_1(object sender, TextChangedEventArgs e)
        {
            EventTime = Action_time_textbox.Text;
        }

        private void Send_btn_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Do you really want to send mail for traffic date: " + DateTime.Today.Subtract(TimeSpan.FromDays(VM.SubtractiveDataDay)).ToShortDateString() + "?",
                "EReport", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                if (VM.ShouldEmailSendorNot)
                {
                    if (To.Text != "")
                    {
                        VM.TaskHandle("MANUAL", false);
                    }
                    else
                    {
                        Show_LogTextblock("Please give at least one recepient at 'To'");
                        MessageBox.Show("Please give at least one recepient at 'To'", "EReport", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                else
                {
                    Show_LogTextblock("Please start email first by clicking 'Start Email'");
                    MessageBox.Show("Please start email first by clicking 'Start Email'", "EReport", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        private void Clear_btn_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Do you really want to clear data?", "EReport", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                To.Text = "";
                CC.Text = "";
                Bcc.Text = "";
                Body.Text = "";
                Sub.Text = "";
                log_textblock.Text = "";
            }
        }

        private void Stop_btn_Click(object sender, RoutedEventArgs e)
        {
            if (VM.ShouldEmailSendorNot)
            {
                VM.ShouldEmailSendorNot = false;
                StopMail_btn.Content = "Start Email";
                if (!VM.MailingProgess)
                {
                    Show_LogTextblock("Auto Email now stopped.");
                }
                else
                {
                    Show_LogTextblock("Please wait to suspend current mail.... ... .. .");
                }
            }
            else
            {
                VM.ShouldEmailSendorNot = true;
                StopMail_btn.Content = "Stop Email";
                Show_LogTextblock("Auto Email started again.");
            }
        }

        //DBWork dbw = new DBWork();
        private void IDD_Diff_btn_Click(object sender, RoutedEventArgs e)
        {
            
            VM.IDDIncomingDifferenceCheck();
        }

        private void To_MouseEnter_1(object sender, MouseEventArgs e)
        {
            Popup_To_textblock.Text = "Please use comma ',' to separate each email address.";
            Popup_To.IsOpen = true;
            timerforPopup.Start();
        }

        private void To_MouseLeave_1(object sender, MouseEventArgs e)
        {
            Popup_To.IsOpen = false;
            timerforPopup.Stop();
        }


        DispatcherTimer timerforPopup = new DispatcherTimer();

        private void timer_TickForPopup(object sender, EventArgs e)
        {
            timerforPopup.Stop();
            Popup_To.IsOpen = false;
            Popup_OpenFolder.IsOpen = false;
        }

        private void acc_psw_PasswordChanged_1(object sender, RoutedEventArgs e)
        {
            VM.Password_String = acc_psw.Password;
        }

        private void user_name_TextChanged_1(object sender, TextChangedEventArgs e)
        {
            VM.User_EmailID = user_email.Text;
        }

        private void SMTP_Client_TextChanged(object sender, TextChangedEventArgs e)
        {
            VM.SMTP_Client = SMTP_Client.Text;
        }

        private void SMTP_Port_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (SMTP_Port.Text == "")
                    VM.SMTP_Port = 587;
                else
                    VM.SMTP_Port = Convert.ToInt32(this.SMTP_Port.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " <" + ex.GetType().ToString() + ">", "EReport", MessageBoxButton.OK, MessageBoxImage.Error);
                VM.Write_logFile(ex.Message);
            }
        }

        private void Settings_OK_btn_Click_1(object sender, RoutedEventArgs e)
        {
            Popup_Settings.IsOpen = false;
        }

        private void IDD_Diff_Checkbox_Checked_1(object sender, RoutedEventArgs e)
        {
            error_checkOrNot = true;
            if(IDD_in_Error_percentage_textbox != null)
                IDD_in_Error_percentage_textbox.IsEnabled = true;
            Show_LogTextblock("Application will check IGW error.");
        }

        private void IDD_Diff_Checkbox_Unchecked_1(object sender, RoutedEventArgs e)
        {
            error_checkOrNot = false;
            IDD_in_Error_percentage_textbox.IsEnabled = false;
            Show_LogTextblock("Application will not check IGW error.");
        }

        private void Error_percentage_textbox_TextChanged_1(object sender, TextChangedEventArgs e)
        {
            VM.IDDAcceptanceString = IDD_in_Error_percentage_textbox.Text;
        }

        private void General_percentage_textbox_TextChanged_1(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (General_percentage_textbox.Text == "")
                    VM.General_acceptance_diff = 1.0;
                else
                    VM.General_acceptance_diff = Convert.ToDouble(General_percentage_textbox.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + " <" + ex.GetType().ToString() + ">", "EReport", MessageBoxButton.OK, MessageBoxImage.Error);
                VM.Write_logFile(ex.Message);
            }
        }

        private void Instructions_MouseEnter_1(object sender, MouseEventArgs e)
        {
            _InstructRun1.Text = "Instructions of using this app:";
            _InstructRun2.Text = Show_Instructions();
            Popup_Instruct.IsOpen = true;
        }

        private void Instructions_MouseLeave_1(object sender, MouseEventArgs e)
        {
            Popup_Instruct.IsOpen = false;
        }



        private string Show_Instructions()
        {
            return
                "\n  1. Select folder from File menu to save excel files. By Default it will save to 'D:\\Excel'" +
                "\n  2. Enter email info of sender email ID from settings." +
                "\n  3. Enter error limit if needed from settings." +
                "\n  4. Click 'Send Mail' button if you want to send mail manually." +
                "\n  5. Adjust the time of action if needed." +
                "\n  6. Each log data will be saved to this directory:- C:\\Users\\Public\\EReport_Log_" + DateTime.Now.Year.ToString() + ".txt." +
                "\n  7. You can disable sending mail by clicking the button 'Stop Mail'." +
                "\n  8. You can check the ICX vs. IGW incoming traffic difference by clicking 'IDD Difference' button.";
        }

        private void OpenFolder_function_MouseEnter(object sender, MouseEventArgs e)
        {
            Popup_OpenFolder_textblock.Text = "Current directory: " + VM.Destination_Excel_Path;
            Popup_OpenFolder.IsOpen = true;
            timerforPopup.Start();
        }

        private void OpenFolder_function_MouseLeave(object sender, MouseEventArgs e)
        {
            Popup_OpenFolder.IsOpen = false;
            timerforPopup.Stop();
        }
    }    
}

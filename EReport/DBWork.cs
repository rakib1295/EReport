using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Oracle.ManagedDataAccess.Client;
using System.Windows;
using System.ComponentModel;
using System.IO;

namespace EReport
{
    class DBWork : INotifyPropertyChanged
    {
        string OracldbLive = "Data Source=(DESCRIPTION =" +
        "(ADDRESS = (PROTOCOL = TCP)(HOST = 10.10.10.9)(PORT = 1521))" +
        "(CONNECT_DATA =" +
        "(SERVER = DEDICATED)" +
        "(SERVICE_NAME = ora11g)));" +
        "User Id= amc1;Password=amc1;";

        ExcelWork EW = new ExcelWork();
        private String _logviewer = "";
        private String _fileLogger = "";
        public String FileLogger
        {
            get { return _fileLogger; }
            set
            {
                _fileLogger = value;
                // Call OnPropertyChanged whenever the property is updated
                OnPropertyChanged("FileLogger");
            }
        }

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

        protected void OnPropertyChanged(string data)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(data));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public String QueryDatabaseforIDDIncoming(int SubtractiveDataDay, String _folderPath, double IDDincoming_acceptance_diff)
        {
            String _filename = "";
            String operator_type = "";
            try
            {
                OracleDataAdapter da;
                OracleConnection conn = new OracleConnection(OracldbLive);  // C#
                conn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.CommandType = CommandType.Text;
                cmd.Connection = conn;

                //////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "select t.TRUNKOUT_OPERATOR, sum(t.CDR_AMOUNT), TO_CHAR(sum(t.DURATION)/60) from cdr_inter_icx_stat t where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') " +
                    " and t.partition_day =  TO_CHAR((sysdate-  " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE = '11' group by t.TRUNKOUT_OPERATOR order by t.TRUNKOUT_OPERATOR";

                DataSet ds1;
                ds1 = new DataSet();
                da = new OracleDataAdapter(cmd);
                da.Fill(ds1);

                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
                Excel.Worksheet xlWorkSheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet1.Name = "Local Operator Wise";

                int rowNum = 1, rowNum1 = 1, rowNum2 = 1;
                int colNum = 1;
                string name = "";
                operator_type = "Called Operators";

                name = "ICX (Overseas to local operator)";
                rowNum1 = EW.ExcelPlot_with_Compare(ref xlWorkSheet1, ref ds1, colNum, name, operator_type);

                ds1.Dispose();
                da.Dispose();

                ////////////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = cmd.CommandText = "select t.CALLED_OPERATOR, sum(t.CDR_AMOUNT), sum(t.DURATION_FLOAT) from cdr_inter_itx_stat t where t.billingcycle = TO_CHAR((sysdate- "
                    + SubtractiveDataDay + "),'yyyymm') and t.partition_day =  TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE in ('23','24','25') group by t.CALLED_OPERATOR order by  t.CALLED_OPERATOR";

                ds1 = new DataSet();
                da = new OracleDataAdapter(cmd);
                da.Fill(ds1);

                name = "IGW (Overseas to local operator)";

                //rowNum++;
                //rowNum++;
                rowNum2 = EW.ExcelPlot_with_Compare(ref xlWorkSheet1, ref ds1, colNum + 5, name, operator_type);

                ds1.Dispose();
                da.Dispose();

                rowNum = EW.Difference_Entry_in_Excel(ref xlWorkSheet1, rowNum1, rowNum2, IDDincoming_acceptance_diff, "LEFT");

                /////////////////////////////////////////////////////////////////////////////////////////////
                //cmd.CommandText = cmd.CommandText = "select t.CALLED_OPERATOR, sum(t.DURATION_FLOAT) from cdr_inter_itx_stat t where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') " +
                //    " and t.partition_day =  TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE in ('23','24','25') and t.RESERVE3 = 'BTCL_IGW1' group by t.CALLED_OPERATOR order by t.CALLED_OPERATOR";
                //ds1 = new DataSet();
                //da = new OracleDataAdapter(cmd);
                //da.Fill(ds1);

                //name = "From IGW1 (Overseas to local operator)";

                //rowNum++;
                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_without_Compare(ref xlWorkSheet1, ref ds1, rowNum, name, operator_type);


                //ds1.Dispose();
                //da.Dispose();


                ////////////////////////////////////////////////////////////////////////////////////////////////
                //cmd.CommandText = cmd.CommandText = "select t.CALLED_OPERATOR, sum(t.DURATION_FLOAT) from cdr_inter_itx_stat t where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') " +
                //    " and t.partition_day =  TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE in ('23','24','25') and t.RESERVE3 = 'ITX7' group by t.CALLED_OPERATOR order by t.CALLED_OPERATOR";
                //ds1 = new DataSet();
                //da = new OracleDataAdapter(cmd);
                //da.Fill(ds1);

                //name = "From ITX7 (Overseas to local operator)";

                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_without_Compare(ref xlWorkSheet1, ref ds1, rowNum, name, operator_type);


                //ds1.Dispose();
                //da.Dispose();

                LogViewer = "IDD In: Successfully created first sheet.";
                /////////////////////////////////first sheet completed//////////////////////////////////////


                Excel.Worksheet xlWorkSheet2 = xlWorkBook.Sheets.Add(misValue, xlWorkSheet1, 1, misValue);
                xlWorkSheet2.Name = "International Carrier Wise";

                /////////////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "select t.TRUNKIN_OPERATOR, sum(t.cdr_amount), sum(t.DURATION_FLOAT), p.network_type from cdr_inter_itx_stat t, prm3.ent_inter_operator_info p "
                    + " where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') and t.partition_day = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE in ('23', '24', '25') and t.TRUNKIN_OPERATOR = p.partner_name "
                    + " group by t.TRUNKIN_OPERATOR,p.network_type order by p.network_type desc, sum(t.DURATION_FLOAT) desc, t.TRUNKIN_OPERATOR";
                ds1 = new DataSet();
                da = new OracleDataAdapter(cmd);
                da.Fill(ds1);

                name = "From IGW (From carriers to IGW)";
                operator_type = "Calling Operators";
                rowNum = 1;
                rowNum = EW.ExcelPlot_without_Compare(ref xlWorkSheet2, ref ds1, rowNum, name, operator_type, "Operator Trype");


                ds1.Dispose();
                da.Dispose();


                /////////////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "select p.network_type, sum(t.cdr_amount), sum(t.DURATION_FLOAT) from cdr_inter_itx_stat t, prm3.ent_inter_operator_info p " +
                    " where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') and t.partition_day = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE in ('23', '24', '25') and t.TRUNKIN_OPERATOR = p.partner_name " +
                    " group by p.network_type order by p.network_type desc, sum(t.DURATION_FLOAT) desc";
                ds1 = new DataSet();
                da = new OracleDataAdapter(cmd);
                da.Fill(ds1);

                name = "";
                operator_type = "Carrier Type";

                rowNum++;
                rowNum++;
                rowNum = EW.ExcelPlot_without_Compare(ref xlWorkSheet2, ref ds1, rowNum, name, operator_type, "");


                ds1.Dispose();
                da.Dispose();


                ///////////////////////////////////////////////////////////////////////////////////////////
                //cmd.CommandText = cmd.CommandText = "select t.TRUNKIN_OPERATOR, sum(t.DURATION_FLOAT) from cdr_inter_itx_stat t where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') " +
                //    "and t.partition_day =  TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE in ('23','24','25') and t.RESERVE3 = 'BTCL_IGW1' group by t.TRUNKIN_OPERATOR order by t.TRUNKIN_OPERATOR";
                //ds1 = new DataSet();
                //da = new OracleDataAdapter(cmd);
                //da.Fill(ds1);

                //name = "From IGW1 (From carriers to IGW)";

                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_without_Compare(ref xlWorkSheet2, ref ds1, rowNum, name, operator_type);


                //ds1.Dispose();
                //da.Dispose();

                //cmd.CommandText = cmd.CommandText = "select t.TRUNKIN_OPERATOR, sum(t.DURATION_FLOAT) from cdr_inter_itx_stat t where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') " +
                //    "and t.partition_day =  TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE in ('23','24','25') and t.RESERVE3 = 'ITX7' group by t.TRUNKIN_OPERATOR order by t.TRUNKIN_OPERATOR";
                //ds1 = new DataSet();
                //da = new OracleDataAdapter(cmd);
                //da.Fill(ds1);

                //name = "From ITX7 (From carriers to IGW)";

                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_without_Compare(ref xlWorkSheet2, ref ds1, rowNum, name, operator_type);


                //ds1.Dispose();
                //da.Dispose();

                LogViewer = "IDD In: Successfully created second sheet.";
                /////////////////////////////////second sheet completed//////////////////////////////////////

                Excel.Worksheet xlWorkSheet3 = xlWorkBook.Sheets.Add(misValue, xlWorkSheet2, 1, misValue);
                xlWorkSheet3.Name = "Hour Wise for Local Operator";


                ///////////////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "select * from (select p.DURATION, p.HOURLY, p.TRUNKOUT_OPERATOR from cdr_inter_icx_d_stat p " +
                    " where p.PARTITION_DAY = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and p.TRANSIT_TYPE = '11' and p.BILLINGCYCLE = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm')) " +
                    " pivot( " +
                    " sum(DURATION) " +
                    " for HOURLY in ('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23') " +
                    " ) piv";

                da = new OracleDataAdapter(cmd);
                ds1 = new DataSet();
                da.Fill(ds1);

                rowNum = 1;
                name = "From ICX (Overseas to local operator)";
                operator_type = "Called Operators";

                rowNum = EW.ExcelPlot_Hourly_Sheet_IDD(ref xlWorkSheet3, ref ds1, rowNum, name, 60, operator_type);

                ds1.Dispose();
                da.Dispose();
                LogViewer = "IDD In: Completed query of hourly ICX data for local operator.";


                //////////////////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "select * from (select p.DURATION_FLOAT, p.HOURLY, p.CALLED_OPERATOR from cdr_inter_itx_d_stat p " +
                    "where p.PARTITION_DAY = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and p.TRANSIT_TYPE in ('23','24','25') and p.BILLINGCYCLE = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm')) " +
                    "pivot(" +
                    " sum(DURATION_FLOAT)" +
                    " for HOURLY in ('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23') " +
                    " ) piv";

                da = new OracleDataAdapter(cmd);
                ds1 = new DataSet();
                da.Fill(ds1);

                name = "From IGW (Overseas to local operator)";
                rowNum++;
                rowNum++;
                rowNum++;
                rowNum = EW.ExcelPlot_Hourly_Sheet_IDD(ref xlWorkSheet3, ref ds1, rowNum, name, 1, operator_type);

                ds1.Dispose();
                da.Dispose();
                LogViewer = "IDD In: Completed query of hourly IGW-Total data for local operator.";



                //////////////////////////////////////////////////////////////////////////////////////////////////
                //cmd.CommandText = "select * from (select p.DURATION_FLOAT, p.HOURLY, p.CALLED_OPERATOR from cdr_inter_itx_d_stat p " +
                //    "where p.PARTITION_DAY = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and p.TRANSIT_TYPE in ('23','24','25') and p.SWITCH_ID = 'BTCL_IGW1' and p.BILLINGCYCLE = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm')) " +
                //    "pivot(" +
                //    " sum(DURATION_FLOAT)" +
                //    " for HOURLY in ('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23') " +
                //    " ) piv";

                //da = new OracleDataAdapter(cmd);
                //ds1 = new DataSet();
                //da.Fill(ds1);

                //name = "From IGW1 (Overseas to local operator)";
                //rowNum++;
                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_Hourly_Sheet_IDD(ref xlWorkSheet3, ref ds1, rowNum, name, 1, operator_type);

                //ds1.Dispose();
                //da.Dispose();
                //LogViewer = "IDD In: Completed query of hourly IGW1 data for local operator.";


                ///////////////////////////////////////////////////////////////////////////////////////////////////
                //cmd.CommandText = "select * from (select p.DURATION_FLOAT, p.HOURLY, p.CALLED_OPERATOR from cdr_inter_itx_d_stat p " +
                //    "where p.PARTITION_DAY = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and p.TRANSIT_TYPE in ('23','24','25') and p.SWITCH_ID = 'ITX7' and p.BILLINGCYCLE = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm')) " +
                //    "pivot(" +
                //    " sum(DURATION_FLOAT)" +
                //    " for HOURLY in ('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23') " +
                //    " ) piv";

                //da = new OracleDataAdapter(cmd);
                //ds1 = new DataSet();
                //da.Fill(ds1);

                //name = "From ITX7 (Overseas to local operator)";
                //rowNum++;
                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_Hourly_Sheet_IDD(ref xlWorkSheet3, ref ds1, rowNum, name, 1, operator_type);

                //ds1.Dispose();
                //da.Dispose();
                //LogViewer = "IDD In: Completed query of hourly ITX7 data for local operator.";

                LogViewer = "IDD In: Successfully created third sheet.";

                ((Excel._Worksheet)xlWorkSheet3).Activate();
                xlWorkSheet3.Application.ActiveWindow.SplitColumn = 2;
                xlWorkSheet3.Application.ActiveWindow.FreezePanes = true;

                /////////////////////////////////third sheet completed//////////////////////////////////////



                Excel.Worksheet xlWorkSheet4 = xlWorkBook.Sheets.Add(misValue, xlWorkSheet3, 1, misValue);
                xlWorkSheet4.Name = "Hour Wise for Carrier";


                ///////////////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "select * from (select p.DURATION_FLOAT, p.HOURLY, p.TRUNKIN_OPERATOR from cdr_inter_itx_d_stat p" +
                    " where p.PARTITION_DAY = TO_CHAR((sysdate- " + SubtractiveDataDay + " ),'dd') and p.TRANSIT_TYPE in ('23','24','25') and p.BILLINGCYCLE = TO_CHAR((sysdate- " + SubtractiveDataDay + " ),'yyyymm')) " +
                    " pivot(" +
                    " sum(DURATION_FLOAT) " +
                    " for HOURLY in ('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23') " +
                    " ) piv";

                da = new OracleDataAdapter(cmd);
                ds1 = new DataSet();
                da.Fill(ds1);

                rowNum = 1;
                name = "From IGW (From carriers to IGW)";
                operator_type = "Calling Operators";

                rowNum = EW.ExcelPlot_Hourly_Sheet_IDD(ref xlWorkSheet4, ref ds1, rowNum, name, 1, operator_type);

                ds1.Dispose();
                da.Dispose();
                LogViewer = "IDD In: Completed query of hourly IGW-Total data for carrier.";



                //////////////////////////////////////////////////////////////////////////////////////////////////
                //cmd.CommandText = "select * from (select p.DURATION_FLOAT, p.HOURLY, p.TRUNKIN_OPERATOR from cdr_inter_itx_d_stat p" +
                //    " where p.PARTITION_DAY = TO_CHAR((sysdate- " + SubtractiveDataDay + " ),'dd') and p.TRANSIT_TYPE in ('23','24','25')  and p.BILLINGCYCLE = TO_CHAR((sysdate- " + SubtractiveDataDay + " ),'yyyymm') and p.SWITCH_ID = 'BTCL_IGW1')" +
                //    " pivot(" +
                //    " sum(DURATION_FLOAT) " +
                //    " for HOURLY in ('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23') " +
                //    " ) piv";

                //da = new OracleDataAdapter(cmd);
                //ds1 = new DataSet();
                //da.Fill(ds1);

                //name = "From IGW1 (From carriers to IGW)";
                //rowNum++;
                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_Hourly_Sheet_IDD(ref xlWorkSheet4, ref ds1, rowNum, name, 1, operator_type);

                //ds1.Dispose();
                //da.Dispose();
                //LogViewer = "IDD In: Completed query of hourly IGW1 data for carrier.";


                ///////////////////////////////////////////////////////////////////////////////////////////////////
                //cmd.CommandText = "select * from (select p.DURATION_FLOAT, p.HOURLY, p.TRUNKIN_OPERATOR from cdr_inter_itx_d_stat p" +
                //    " where p.PARTITION_DAY = TO_CHAR((sysdate- " + SubtractiveDataDay + " ),'dd') and p.TRANSIT_TYPE in ('23','24','25') and p.BILLINGCYCLE = TO_CHAR((sysdate- " + SubtractiveDataDay + " ),'yyyymm') and p.SWITCH_ID = 'ITX7')" +
                //    " pivot(" +
                //    " sum(DURATION_FLOAT) " +
                //    " for HOURLY in ('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23') " +
                //    " ) piv";

                //da = new OracleDataAdapter(cmd);
                //ds1 = new DataSet();
                //da.Fill(ds1);

                //name = "From ITX7 (From carriers to IGW)";
                //rowNum++;
                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_Hourly_Sheet_IDD(ref xlWorkSheet4, ref ds1, rowNum, name, 1, operator_type);

                //ds1.Dispose();
                //da.Dispose();
                //LogViewer = "IDD In: Completed query of hourly ITX7 data for carrier.";

                LogViewer = "IDD In: Successfully created fourth sheet.";


                ((Excel._Worksheet)xlWorkSheet4).Activate();
                xlWorkSheet4.Application.ActiveWindow.SplitColumn = 2;
                xlWorkSheet4.Application.ActiveWindow.FreezePanes = true;

                /////////////////////////////////fourth sheet completed//////////////////////////////////////


                ((Excel._Worksheet)xlWorkSheet1).Activate();
                xlApp.DisplayAlerts = false;

                string yy = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("yy");
                string MMM = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("MMM");
                string dd = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("dd");

                _filename = _folderPath + "\\IDD_In_" + dd + "-" + MMM + "-" + yy /*DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToShortDateString()*/ + ".xls";
                xlWorkBook.SaveAs(_filename, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                LogViewer = "Successfully created 1st excel file for IDD incoming traffic data.";


                cmd.Dispose();
                conn.Dispose();
                releaseObject(xlWorkSheet1);
                releaseObject(xlWorkSheet2);
                releaseObject(xlWorkSheet3);
                releaseObject(xlWorkSheet4);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                LogViewer = "Exception in IDD incoming file creation: " + ex.Message;
                FileLogger = LogViewer;
                MessageBox.Show(LogViewer, "EReport", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return _filename;            
        }


        public String QueryDatabaseforIDDOutgoing(int SubtractiveDataDay, String _folderPath, double General_acceptance_diff)
        {
            String _filename = "";
            String operator_type = "";
            try
            {
                OracleDataAdapter da;
                OracleConnection conn = new OracleConnection(OracldbLive);  // C#
                conn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.CommandType = CommandType.Text;
                cmd.Connection = conn;

                //////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "select t.TRUNKIN_OPERATOR, sum(t.cdr_amount), TO_CHAR(sum(t.DURATION)/60) from cdr_inter_icx_stat t where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') " +
                    " and t.partition_day =  TO_CHAR((sysdate-  " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE = '12' group by t.TRUNKIN_OPERATOR order by t.TRUNKIN_OPERATOR";

                DataSet ds1;
                ds1 = new DataSet();
                da = new OracleDataAdapter(cmd);
                da.Fill(ds1);

                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
                Excel.Worksheet xlWorkSheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet1.Name = "Local Operator Wise";

                int rowNum = 1, rowNum1 = 1, rowNum2 = 1;
                string name = "";

                name = "ICX (From local operators to overseas)";
                operator_type = "Calling Operators";

                rowNum1 = EW.ExcelPlot_with_Compare(ref xlWorkSheet1, ref ds1, 1, name, operator_type);

                ds1.Dispose();
                da.Dispose();

                ////////////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = cmd.CommandText = "select t.CALLING_OPERATOR, sum(t.cdr_amount), sum(t.DURATION_FLOAT) from cdr_inter_itx_stat t where t.billingcycle = TO_CHAR((sysdate- "
                    + SubtractiveDataDay + "),'yyyymm') and t.partition_day =  TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE in ('20','21','22') group by t.CALLING_OPERATOR order by  t.CALLING_OPERATOR";

                ds1 = new DataSet();
                da = new OracleDataAdapter(cmd);
                da.Fill(ds1);

                name = "IGW (From local operators to overseas)";

                rowNum++;
                rowNum++;
                rowNum2 = EW.ExcelPlot_with_Compare(ref xlWorkSheet1, ref ds1, 6, name, operator_type);

                ds1.Dispose();
                da.Dispose();


                rowNum = EW.Difference_Entry_in_Excel(ref xlWorkSheet1, rowNum1, rowNum2, General_acceptance_diff, "RIGHT");

                /////////////////////////////////////////////////////////////////////////////////////////////
                //cmd.CommandText = cmd.CommandText = "select t.CALLING_OPERATOR, sum(t.DURATION_FLOAT) from cdr_inter_itx_stat t where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') " +
                //    " and t.partition_day =  TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE in ('20','21','22') and t.RESERVE3 = 'BTCL_IGW1' group by t.CALLING_OPERATOR order by t.CALLING_OPERATOR";
                //ds1 = new DataSet();
                //da = new OracleDataAdapter(cmd);
                //da.Fill(ds1);

                //name = "From IGW1 (From local operators to overseas)";

                //rowNum++;
                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_without_Compare(ref xlWorkSheet1, ref ds1, rowNum, name, operator_type);


                //ds1.Dispose();
                //da.Dispose();


                ////////////////////////////////////////////////////////////////////////////////////////////////
                //cmd.CommandText = cmd.CommandText = "select t.CALLING_OPERATOR, sum(t.DURATION_FLOAT) from cdr_inter_itx_stat t where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') " +
                //    " and t.partition_day =  TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE in ('20','21','22') and t.RESERVE3 = 'ITX7' group by t.CALLING_OPERATOR order by t.CALLING_OPERATOR";
                //ds1 = new DataSet();
                //da = new OracleDataAdapter(cmd);
                //da.Fill(ds1);

                //name = "From ITX7 (From local operators to overseas)";

                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_without_Compare(ref xlWorkSheet1, ref ds1, rowNum, name, operator_type);


                //ds1.Dispose();
                //da.Dispose();

                LogViewer = "IDD Out: Successfully created first sheet.";
                /////////////////////////////////first sheet completed//////////////////////////////////////


                Excel.Worksheet xlWorkSheet2 = xlWorkBook.Sheets.Add(misValue, xlWorkSheet1, 1, misValue);
                xlWorkSheet2.Name = "International Carrier Wise";


                /////////////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = cmd.CommandText = "select t.TRUNKOUT_OPERATOR, sum(t.cdr_amount), sum(t.DURATION_FLOAT) from cdr_inter_itx_stat t where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') " +
                    "and t.partition_day =  TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE in ('20','21','22') group by t.TRUNKOUT_OPERATOR order by t.TRUNKOUT_OPERATOR";
                ds1 = new DataSet();
                da = new OracleDataAdapter(cmd);
                da.Fill(ds1);

                name = "From IGW (IGW to carriers)";
                operator_type = "Called Operators";

                rowNum = 1;
                rowNum = EW.ExcelPlot_without_Compare(ref xlWorkSheet2, ref ds1, rowNum, name, operator_type, "");


                ds1.Dispose();
                da.Dispose();


                ///////////////////////////////////////////////////////////////////////////////////////////
                //cmd.CommandText = cmd.CommandText = "select t.TRUNKOUT_OPERATOR, sum(t.DURATION_FLOAT) from cdr_inter_itx_stat t where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') " +
                //    "and t.partition_day =  TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE in ('20','21','22')  and t.RESERVE3 = 'BTCL_IGW1' group by t.TRUNKOUT_OPERATOR order by t.TRUNKOUT_OPERATOR";
                //ds1 = new DataSet();
                //da = new OracleDataAdapter(cmd);
                //da.Fill(ds1);

                //name = "From IGW1 (IGW to carriers)";

                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_without_Compare(ref xlWorkSheet2, ref ds1, rowNum, name, operator_type);


                //ds1.Dispose();
                //da.Dispose();

                //cmd.CommandText = cmd.CommandText = "select t.TRUNKOUT_OPERATOR, sum(t.DURATION_FLOAT) from cdr_inter_itx_stat t where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') " +
                //    "and t.partition_day =  TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE in ('20','21','22')  and t.RESERVE3 = 'ITX7' group by t.TRUNKOUT_OPERATOR order by t.TRUNKOUT_OPERATOR";
                //ds1 = new DataSet();
                //da = new OracleDataAdapter(cmd);
                //da.Fill(ds1);

                //name = "From ITX7 (IGW to carriers)";

                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_without_Compare(ref xlWorkSheet2, ref ds1, rowNum, name, operator_type);


                //ds1.Dispose();
                //da.Dispose();

                LogViewer = "IDD Out: Successfully created second sheet.";
                /////////////////////////////////second sheet completed//////////////////////////////////////

                Excel.Worksheet xlWorkSheet3 = xlWorkBook.Sheets.Add(misValue, xlWorkSheet2, 1, misValue);
                xlWorkSheet3.Name = "Hour Wise for Local Operator";


                ///////////////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "select * from (select p.DURATION, p.HOURLY, p.TRUNKIN_OPERATOR from cdr_inter_icx_d_stat p " +
                    " where p.PARTITION_DAY = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and p.TRANSIT_TYPE = '12' and p.BILLINGCYCLE = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm')) " +
                    " pivot( " +
                    " sum(DURATION) " +
                    " for HOURLY in ('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23') " +
                    " ) piv";

                da = new OracleDataAdapter(cmd);
                ds1 = new DataSet();
                da.Fill(ds1);

                rowNum = 1;
                name = "From ICX (From local operators to overseas)";
                operator_type = "Calling Operators";

                rowNum = EW.ExcelPlot_Hourly_Sheet_IDD(ref xlWorkSheet3, ref ds1, rowNum, name, 60, operator_type);

                ds1.Dispose();
                da.Dispose();
                LogViewer = "IDD Out: Completed query of hourly ICX data for local operator.";


                //////////////////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "select * from (select p.DURATION_FLOAT, p.HOURLY, p.CALLING_OPERATOR from cdr_inter_itx_d_stat p " +
                    "where p.PARTITION_DAY = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and p.TRANSIT_TYPE in ('20','21','22') and p.BILLINGCYCLE = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm')) " +
                    "pivot(" +
                    " sum(DURATION_FLOAT)" +
                    " for HOURLY in ('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23') " +
                    " ) piv";

                da = new OracleDataAdapter(cmd);
                ds1 = new DataSet();
                da.Fill(ds1);

                name = "From IGW (From local operators to overseas)";
                rowNum++;
                rowNum++;
                rowNum++;
                rowNum = EW.ExcelPlot_Hourly_Sheet_IDD(ref xlWorkSheet3, ref ds1, rowNum, name, 1, operator_type);

                ds1.Dispose();
                da.Dispose();
                LogViewer = "IDD Out: Completed query of hourly IGW-Total data for local operator.";



                //////////////////////////////////////////////////////////////////////////////////////////////////
                //cmd.CommandText = "select * from (select p.DURATION_FLOAT, p.HOURLY, p.CALLING_OPERATOR from cdr_inter_itx_d_stat p " +
                //    "where p.PARTITION_DAY = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and p.TRANSIT_TYPE in ('20','21','22') and p.SWITCH_ID = 'BTCL_IGW1' and p.BILLINGCYCLE = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm')) " +
                //    "pivot(" +
                //    " sum(DURATION_FLOAT)" +
                //    " for HOURLY in ('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23') " +
                //    " ) piv";

                //da = new OracleDataAdapter(cmd);
                //ds1 = new DataSet();
                //da.Fill(ds1);

                //name = "From IGW1 (From local operators to overseas)";
                //rowNum++;
                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_Hourly_Sheet_IDD(ref xlWorkSheet3, ref ds1, rowNum, name, 1, operator_type);

                //ds1.Dispose();
                //da.Dispose();
                //LogViewer = "IDD Out: Completed query of hourly IGW1 data for local operator.";


                ///////////////////////////////////////////////////////////////////////////////////////////////////
                //cmd.CommandText = "select * from (select p.DURATION_FLOAT, p.HOURLY, p.CALLING_OPERATOR from cdr_inter_itx_d_stat p " +
                //    "where p.PARTITION_DAY = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and p.TRANSIT_TYPE in ('20','21','22') and p.SWITCH_ID = 'ITX7' and p.BILLINGCYCLE = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm')) " +
                //    "pivot(" +
                //    " sum(DURATION_FLOAT)" +
                //    " for HOURLY in ('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23') " +
                //    " ) piv";

                //da = new OracleDataAdapter(cmd);
                //ds1 = new DataSet();
                //da.Fill(ds1);

                //name = "From ITX7 (From local operators to overseas)";
                //rowNum++;
                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_Hourly_Sheet_IDD(ref xlWorkSheet3, ref ds1, rowNum, name, 1, operator_type);

                //ds1.Dispose();
                //da.Dispose();
                //LogViewer = "IDD Out: Completed query of hourly ITX7 data for local operator.";

                LogViewer = "IDD Out: Successfully created third sheet.";

                //Excel.Range _col = (Excel.Range)xlWorkSheet3.Columns[3];
                //_col.Activate();
                //_col.Application.ActiveWindow.SplitRow = 1;
                //_col.Application.ActiveWindow.FreezePanes = true;

                ((Excel._Worksheet)xlWorkSheet3).Activate();
                xlWorkSheet3.Application.ActiveWindow.SplitColumn = 2;
                xlWorkSheet3.Application.ActiveWindow.FreezePanes = true;

                /////////////////////////////////third sheet completed//////////////////////////////////////



                Excel.Worksheet xlWorkSheet4 = xlWorkBook.Sheets.Add(misValue, xlWorkSheet3, 1, misValue);
                xlWorkSheet4.Name = "Hour Wise for Carrier";


                ///////////////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "select * from (select p.DURATION_FLOAT, p.HOURLY, p.TRUNKOUT_OPERATOR from cdr_inter_itx_d_stat p" +
                    " where p.PARTITION_DAY = TO_CHAR((sysdate- " + SubtractiveDataDay + " ),'dd') and p.TRANSIT_TYPE in ('20','21','22') and p.BILLINGCYCLE = TO_CHAR((sysdate- " + SubtractiveDataDay + " ),'yyyymm')) " +
                    " pivot(" +
                    " sum(DURATION_FLOAT) " +
                    " for HOURLY in ('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23') " +
                    " ) piv";

                da = new OracleDataAdapter(cmd);
                ds1 = new DataSet();
                da.Fill(ds1);

                rowNum = 1;
                name = "From IGW1 (IGW to carriers)";
                operator_type = "Called Operatos";

                rowNum = EW.ExcelPlot_Hourly_Sheet_IDD(ref xlWorkSheet4, ref ds1, rowNum, name, 1, operator_type);

                ds1.Dispose();
                da.Dispose();
                LogViewer = "IDD Out: Completed query of hourly IGW-Total data for carrier.";



                //////////////////////////////////////////////////////////////////////////////////////////////////
                //cmd.CommandText = "select * from (select p.DURATION_FLOAT, p.HOURLY, p.TRUNKOUT_OPERATOR from cdr_inter_itx_d_stat p" +
                //    " where p.PARTITION_DAY = TO_CHAR((sysdate- " + SubtractiveDataDay + " ),'dd') and p.TRANSIT_TYPE in ('20','21','22') and p.BILLINGCYCLE = TO_CHAR((sysdate- " + SubtractiveDataDay + " ),'yyyymm') and p.SWITCH_ID = 'BTCL_IGW1')" +
                //    " pivot(" +
                //    " sum(DURATION_FLOAT) " +
                //    " for HOURLY in ('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23') " +
                //    " ) piv";

                //da = new OracleDataAdapter(cmd);
                //ds1 = new DataSet();
                //da.Fill(ds1);

                //name = "From IGW1 (IGW to carriers)";
                //rowNum++;
                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_Hourly_Sheet_IDD(ref xlWorkSheet4, ref ds1, rowNum, name, 1, operator_type);

                //ds1.Dispose();
                //da.Dispose();
                //LogViewer = "IDD Out: Completed query of hourly IGW1 data for carrier.";


                ///////////////////////////////////////////////////////////////////////////////////////////////////
                //cmd.CommandText = "select * from (select p.DURATION_FLOAT, p.HOURLY, p.TRUNKOUT_OPERATOR from cdr_inter_itx_d_stat p" +
                //    " where p.PARTITION_DAY = TO_CHAR((sysdate- " + SubtractiveDataDay + " ),'dd') and p.TRANSIT_TYPE in ('20','21','22') and p.BILLINGCYCLE = TO_CHAR((sysdate- " + SubtractiveDataDay + " ),'yyyymm') and p.SWITCH_ID = 'ITX7')" +
                //    " pivot(" +
                //    " sum(DURATION_FLOAT) " +
                //    " for HOURLY in ('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23') " +
                //    " ) piv";

                //da = new OracleDataAdapter(cmd);
                //ds1 = new DataSet();
                //da.Fill(ds1);

                //name = "From ITX7 (IGW to carriers)";
                //rowNum++;
                //rowNum++;
                //rowNum++;
                //rowNum = EW.ExcelPlot_Hourly_Sheet_IDD(ref xlWorkSheet4, ref ds1, rowNum, name, 1, operator_type);

                //ds1.Dispose();
                //da.Dispose();
                //LogViewer = "IDD Out: Completed query of hourly ITX7 data for carrier.";

                LogViewer = "IDD Out: Successfully created fourth sheet.";


                ((Excel._Worksheet)xlWorkSheet4).Activate();
                xlWorkSheet4.Application.ActiveWindow.SplitColumn = 2;
                xlWorkSheet4.Application.ActiveWindow.FreezePanes = true;

                /////////////////////////////////fourth sheet completed//////////////////////////////////////


                ((Excel._Worksheet)xlWorkSheet1).Activate();
                xlApp.DisplayAlerts = false;

                string yy = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("yy");
                string MMM = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("MMM");
                string dd = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("dd");

                _filename = _folderPath + "\\IDD_Out_" + dd + "-" + MMM + "-" + yy /*DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToShortDateString()*/ + ".xls";
                xlWorkBook.SaveAs(_filename, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                LogViewer = "Successfully created 2nd excel file for IDD outgoing traffic data.";


                cmd.Dispose();
                conn.Dispose();
                releaseObject(xlWorkSheet1);
                releaseObject(xlWorkSheet2);
                releaseObject(xlWorkSheet3);
                releaseObject(xlWorkSheet4);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                LogViewer = "Exception in IDD outgoing file creation: " + ex.Message;
                FileLogger = LogViewer;
                MessageBox.Show(LogViewer, "EReport", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            return _filename;
        }

        public double IDDIncomingDifferenceCheck(int SubtractiveDataDay)
        {
            double IDDincoming_Diff_actual = 0;
            OracleConnection conn = new OracleConnection(OracldbLive);  // C#
            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.CommandType = CommandType.Text;
            cmd.Connection = conn;

            double itx_val = 0, icx_val = 0;
            try
            {
                cmd.CommandText = "select sum(t.DURATION_FLOAT) MINUTES from cdr_inter_itx_stat t where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') and t.partition_day =  TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE in ('23','24','25')";
                
                OracleDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string s = "";
                    s = reader["MINUTES"].ToString();
                    if (s != "")
                        itx_val = Convert.ToDouble(s);
                }


                cmd.CommandText = "select sum(t.duration) MINUTES from cdr_inter_icx_stat t where t.billingcycle = TO_CHAR((sysdate- " + SubtractiveDataDay + "),'yyyymm') and t.partition_day =  TO_CHAR((sysdate- " + SubtractiveDataDay + "),'dd') and t.TRANSIT_TYPE = '11'";
                
                reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string s = "";
                    s = reader["MINUTES"].ToString();
                    if (s != "")
                        icx_val = Convert.ToDouble(s);
                }


                icx_val = icx_val / 60;
                reader.Dispose();

                cmd.Dispose();
                conn.Dispose();
            }
            catch (Exception ex)
            {
                LogViewer = "Exception in IDD incoming difference check: " + ex.Message;
                FileLogger = LogViewer;
                MessageBox.Show(LogViewer, "EReport", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            IDDincoming_Diff_actual = (icx_val - itx_val) * 100 / icx_val;
            if (IDDincoming_Diff_actual < 0)
            {
                IDDincoming_Diff_actual = IDDincoming_Diff_actual * (-1);
                LogViewer = "IDD incoming traffic difference: ICX is " + IDDincoming_Diff_actual.ToString("#0.00") + "% less than IGW for date: " + DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToShortDateString() + ".";
            }
            else
            {
                LogViewer = "IDD incoming traffic difference: ICX is " + IDDincoming_Diff_actual.ToString("#0.00") + "% greater than IGW for date: " + DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToShortDateString() + ".";
            }
            return IDDincoming_Diff_actual;
        }



        //public String QueryDatabaseforANSLocalIncoming(int SubtractiveDataDay, String _folderPath, double Local_acceptance_diff)
        //{
        //    String _filename = "";

        //    try
        //    {
        //        OracleDataAdapter da;
        //        OracleConnection conn = new OracleConnection(OracldbLive);  // C#
        //        conn.Open();
        //        OracleCommand cmd = new OracleCommand();
        //        cmd.CommandType = CommandType.Text;
        //        cmd.Connection = conn;


        //        //////////////////////////////////////////////////////////////////////////////////////
        //        cmd.CommandText = "select trunkin_operator, sum(t.cdr_amount), TO_CHAR(sum(t.duration) / 60) from cdr_inter_icx_stat t " +
        //            " where t.billingcycle = TO_CHAR((sysdate - " + SubtractiveDataDay + "), 'yyyymm') " + " and t.PARTITION_DAY = TO_CHAR((sysdate-  " + SubtractiveDataDay + "),'dd') " +
        //            " and t.trunkout_operator = 'BTCL' and t.transit_type = '10' group by t.trunkin_operator order by t.trunkin_operator";


        //        DataSet ds1;
        //        ds1 = new DataSet();
        //        da = new OracleDataAdapter(cmd);
        //        da.Fill(ds1);

        //        object misValue = System.Reflection.Missing.Value;
        //        Excel.Application xlApp = new Excel.Application();
        //        Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
        //        Excel.Worksheet xlWorkSheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
        //        xlWorkSheet1.Name = "ANS Local Incoming";

        //        int colNum = 1;
        //        int rowNum1 = 0, rowNum2 = 0;
        //        string name = "";
        //        name = "From ICX (Other operators to BTCL ANS)";
        //        string operator_type = "Calling Operator";

        //        rowNum1 = EW.ExcelPlot_with_Compare(ref xlWorkSheet1, ref ds1, colNum, name, operator_type);

        //        ds1.Dispose();
        //        da.Dispose();

        //        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //        ///
        //        cmd.CommandText = "select t.settle_operator_1, sum(t.cdr_amount), sum(t.duration_float)  from cdr_inter_ans_stat t " +
        //            "where t.billingcycle = TO_CHAR((sysdate - " + SubtractiveDataDay + "), 'yyyymm') " + "and t.PARTITION_DAY = TO_CHAR((sysdate-  " + SubtractiveDataDay + "),'dd') " +
        //            " and t.transit_type = '33' and t.trunkin_operator = 'BTCL ICX' group by t.settle_operator_1 order by t.settle_operator_1";


        //        ds1 = new DataSet();
        //        da = new OracleDataAdapter(cmd);
        //        da.Fill(ds1);

        //        name = "From ANS (Other operators to BTCL ANS)";
        //        operator_type = "Calling Operator";
        //        colNum += 5;

        //        rowNum2 = EW.ExcelPlot_with_Compare(ref xlWorkSheet1, ref ds1, colNum, name, operator_type);

        //        ds1.Dispose();
        //        da.Dispose();

        //        EW.Difference_Entry_in_Excel(ref xlWorkSheet1, rowNum1, rowNum2, Local_acceptance_diff, "LEFT");

        //        //LogViewer = "ANS Out: Successfully created first sheet.";
        //        ///////////////////////////////////////////////////////////////First sheet completed/////////////////////
        //        ///


        //        //////////////////////////////////////////////////////////////////////Second Sheet completed/////////////
        //        ///
        //        ((Excel._Worksheet)xlWorkSheet1).Activate();

        //        xlApp.DisplayAlerts = false;

        //        string yy = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("yy");
        //        string MMM = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("MMM");
        //        string dd = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("dd");

        //        _filename = _folderPath + "\\ANS_In_Local_" + dd + "-" + MMM + "-" + yy /*DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToShortDateString()*/ + ".xls";
        //        xlWorkBook.SaveAs(_filename, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        //        xlWorkBook.Close(true, misValue, misValue);
        //        xlApp.Quit();

        //        LogViewer = "Successfully created 3rd excel file for ANS Local incoming traffic data.";


        //        cmd.Dispose();
        //        conn.Dispose();
        //        releaseObject(xlWorkSheet1);
        //        releaseObject(xlWorkBook);
        //        releaseObject(xlApp);
        //    }
        //    catch (Exception ex)
        //    {
        //        LogViewer = "Exception in ANS local incoming file creation: " + ex.Message;
        //        FileLogger = LogViewer;
        //        MessageBox.Show(LogViewer, "EReport", MessageBoxButton.OK, MessageBoxImage.Error);
        //    }

        //    return _filename;
        //}

        public String QueryDatabaseforANSLocal(int SubtractiveDataDay, String _folderPath, double Local_acceptance_diff)
        {
            String _filename = "";

            try
            {
                OracleDataAdapter da;
                OracleConnection conn = new OracleConnection(OracldbLive);  // C#
                conn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.CommandType = CommandType.Text;
                cmd.Connection = conn;


                //////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "select trunkout_operator, sum(t.cdr_amount), TO_CHAR(sum(t.duration)/60) from cdr_inter_icx_stat t " +
                " where t.billingcycle = TO_CHAR((sysdate - " + SubtractiveDataDay + "), 'yyyymm') " + " and t.PARTITION_DAY = TO_CHAR((sysdate-  " + SubtractiveDataDay + "),'dd') " +
                " and t.trunkin_operator = 'BTCL' and t.transit_type = '10' group by t.trunkout_operator order by t.trunkout_operator";
                
                DataSet ds1;
                ds1 = new DataSet();
                da = new OracleDataAdapter(cmd);
                da.Fill(ds1);

                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
                Excel.Worksheet xlWorkSheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet1.Name = "ANS Local Outgoing";

                int colNum = 1;
                int rowNum1 = 0, rowNum2 = 0;
                string name = "";
                name = "From ICX (BTCL ANS to Other operators)";
                string operator_type = "Called Operator";

                rowNum1 = EW.ExcelPlot_with_Compare(ref xlWorkSheet1, ref ds1, colNum, name, operator_type);

                ds1.Dispose();
                da.Dispose();

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ///
                cmd.CommandText = "select t.settle_operator_1, sum(t.cdr_amount), sum(t.duration_float) from cdr_inter_ans_stat t " +
                "where t.billingcycle = TO_CHAR((sysdate - " + SubtractiveDataDay + "), 'yyyymm') " + " and t.PARTITION_DAY = TO_CHAR((sysdate-  " + SubtractiveDataDay + "),'dd') " +
                " and t.transit_type = '31' and trunkout_operator = 'BTCL ICX' group by t.settle_operator_1 order by t.settle_operator_1";


                ds1 = new DataSet();
                da = new OracleDataAdapter(cmd);
                da.Fill(ds1);

                name = "From ANS (BTCL ANS to Other operators)";
                operator_type = "Called Operator";
                colNum += 5;

                rowNum2 = EW.ExcelPlot_with_Compare(ref xlWorkSheet1, ref ds1, colNum, name, operator_type);

                ds1.Dispose();
                da.Dispose();

                EW.Difference_Entry_in_Excel(ref xlWorkSheet1, rowNum1, rowNum2, Local_acceptance_diff, "LEFT");
                LogViewer = "ANS Local Out: Successfully created 1st sheet.";


                //////////////////////////////////////////////////////////////////////First Sheet completed/////////////
                ///


                //////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "select trunkin_operator, sum(t.cdr_amount), TO_CHAR(sum(t.duration) / 60) from cdr_inter_icx_stat t " +
                    " where t.billingcycle = TO_CHAR((sysdate - " + SubtractiveDataDay + "), 'yyyymm') " + " and t.PARTITION_DAY = TO_CHAR((sysdate-  " + SubtractiveDataDay + "),'dd') " +
                    " and t.trunkout_operator = 'BTCL' and t.transit_type = '10' group by t.trunkin_operator order by t.trunkin_operator";


                ds1 = new DataSet();
                da = new OracleDataAdapter(cmd);
                da.Fill(ds1);

                
                Excel.Worksheet xlWorkSheet2 = xlWorkBook.Sheets.Add(misValue, xlWorkSheet1, 1, misValue);
                xlWorkSheet2.Name = "ANS Local Incoming";

                colNum = 1;
                rowNum1 = 0;
                rowNum2 = 0;
                name = "";
                name = "From ICX (Other operators to BTCL ANS)";
                operator_type = "Calling Operator";

                rowNum1 = EW.ExcelPlot_with_Compare(ref xlWorkSheet2, ref ds1, colNum, name, operator_type);

                ds1.Dispose();
                da.Dispose();

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ///
                cmd.CommandText = "select t.settle_operator_1, sum(t.cdr_amount), sum(t.duration_float)  from cdr_inter_ans_stat t " +
                    "where t.billingcycle = TO_CHAR((sysdate - " + SubtractiveDataDay + "), 'yyyymm') " + "and t.PARTITION_DAY = TO_CHAR((sysdate-  " + SubtractiveDataDay + "),'dd') " +
                    " and t.transit_type = '33' and t.trunkin_operator = 'BTCL ICX' group by t.settle_operator_1 order by t.settle_operator_1";


                ds1 = new DataSet();
                da = new OracleDataAdapter(cmd);
                da.Fill(ds1);

                name = "From ANS (Other operators to BTCL ANS)";
                operator_type = "Calling Operator";
                colNum += 5;

                rowNum2 = EW.ExcelPlot_with_Compare(ref xlWorkSheet2, ref ds1, colNum, name, operator_type);

                ds1.Dispose();
                da.Dispose();

                EW.Difference_Entry_in_Excel(ref xlWorkSheet2, rowNum1, rowNum2, Local_acceptance_diff, "LEFT");

                LogViewer = "ANS Local In: Successfully created 2nd sheet.";


                //////////////////////////////////////////////////////////////////////Second sheet completed/////////////////////////

                ((Excel._Worksheet)xlWorkSheet1).Activate();

                xlApp.DisplayAlerts = false;

                string yy = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("yy");
                string MMM = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("MMM");
                string dd = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("dd");

                _filename = _folderPath + "\\BTCL_Local_" + dd + "-" + MMM + "-" + yy /*DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToShortDateString()*/ + ".xls";
                xlWorkBook.SaveAs(_filename, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                LogViewer = "Successfully created 3rd excel file for ANS Local traffic data.";

                cmd.Dispose();
                conn.Dispose();
                releaseObject(xlWorkSheet1);
                releaseObject(xlWorkSheet2);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                LogViewer = "Exception in ANS local file creation: " + ex.Message;
                FileLogger = LogViewer;
                MessageBox.Show(LogViewer, "EReport", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            return _filename;
        }


        public String QueryDatabaseforICX(int SubtractiveDataDay, String _folderPath)
        {
            String _filename = "";

            try
            {
                OracleDataAdapter da;
                OracleConnection conn = new OracleConnection(OracldbLive);  // C#
                conn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.CommandType = CommandType.Text;
                cmd.Connection = conn;

                //////////////////////////////////////////////////////////////////////////////////////
                cmd.CommandText = "select distinct(t.TRUNKOUT_OPERATOR) as TRUNKOUT_OPERATOR from cdr_inter_icx_stat t " +
                " where t.BILLINGCYCLE = TO_CHAR((sysdate -  "+ SubtractiveDataDay + "), 'yyyymm') and t.PARTITION_DAY = TO_CHAR((sysdate - " + SubtractiveDataDay + "), 'dd') and t.TRANSIT_TYPE = '10' ";

                OracleDataReader reader = cmd.ExecuteReader();

                List<String> TRUNKOUT_OPERATOR = new List<string>();
                while (reader.Read())
                {
                    TRUNKOUT_OPERATOR.Add(reader["TRUNKOUT_OPERATOR"].ToString());
                }

                String _trunkout_operator = "";
                for(int i = 0; i < TRUNKOUT_OPERATOR.Count; i++)
                {
                    _trunkout_operator = _trunkout_operator + "'" + TRUNKOUT_OPERATOR[i] + "',";
                }
                _trunkout_operator = _trunkout_operator.Substring(0, _trunkout_operator.Length - 1);

                reader.Dispose();

                ///////////////////////////////////////////////////////////////////////////////////////////////////////////
                ///
                cmd.CommandText = "select * from " +
                    " (select p.DURATION, p.TRUNKOUT_OPERATOR, p.TRUNKIN_OPERATOR from cdr_inter_icx_stat p " +
                    " where p.PARTITION_DAY = TO_CHAR((sysdate - " + SubtractiveDataDay + "), 'dd') and p.TRANSIT_TYPE = '10' and p.BILLINGCYCLE = TO_CHAR((sysdate - " + SubtractiveDataDay + "), 'yyyymm')) " +
                    " pivot(" +
                    " SUM(DURATION) for TRUNKOUT_OPERATOR in (" + _trunkout_operator + ")" +
                    " ) piv";

                da = new OracleDataAdapter(cmd);
                DataSet ds1 = new DataSet();
                da.Fill(ds1);

                int rowNum = 1;

                object misValue = System.Reflection.Missing.Value;
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
                Excel.Worksheet xlWorkSheet1 = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlWorkSheet1.Name = "ICX Local Calls";

                rowNum = EW.ExcelPlot_ICX_Local(ref xlWorkSheet1, ref ds1, rowNum, 60, TRUNKOUT_OPERATOR);

                xlWorkSheet1.Application.ActiveWindow.SplitColumn = 2;
                xlWorkSheet1.Application.ActiveWindow.FreezePanes = true;

                xlWorkSheet1.Application.ActiveWindow.SplitRow = 2;
                xlWorkSheet1.Application.ActiveWindow.FreezePanes = true;

                xlApp.DisplayAlerts = false;

                string yy = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("yy");
                string MMM = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("MMM");
                string dd = DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToString("dd");

                _filename = _folderPath + "\\ICX_Local_" + dd + "-" + MMM + "-" + yy /*DateTime.Today.Subtract(TimeSpan.FromDays(SubtractiveDataDay)).ToShortDateString()*/ + ".xls";
                xlWorkBook.SaveAs(_filename, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();


                cmd.Dispose();
                conn.Dispose();
                releaseObject(xlWorkSheet1);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                ds1.Dispose();
                da.Dispose();
                LogViewer = "Successfully created 5th excel file for ICX Local traffic data.";
            }
            catch(Exception ex)
            {
                LogViewer = "Exception in ICX local file creation: " + ex.Message;
                FileLogger = LogViewer;
                MessageBox.Show(LogViewer, "EReport", MessageBoxButton.OK, MessageBoxImage.Error);
            }

                return _filename;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                LogViewer = "Exception Occured while releasing object " + ex.ToString();
                FileLogger = LogViewer;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString(), "EReport", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}

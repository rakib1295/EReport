using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace EReport
{
    class ExcelWork
    {
        public int ExcelPlot_with_Compare(ref Excel.Worksheet _xlWorkSheet, ref DataSet ds1, int colNum, string name, string operator_type)
        {
            int starting_num = colNum;
            string data;
            int i, j;

            //double d;

            //var f = new NumberFormatInfo { NumberGroupSeparator = "," };


            _xlWorkSheet.Cells[1, colNum] = name;
            _xlWorkSheet.Cells[1, colNum].EntireRow.Font.Bold = true;

            Excel.Range _range1 = (Excel.Range)_xlWorkSheet.Cells[1, colNum];
            Excel.Range _range2 = (Excel.Range)_xlWorkSheet.Cells[1, colNum + 3];
            Excel.Range workSheet_range = _xlWorkSheet.get_Range(_range1, _range2);
            workSheet_range.Merge();
            workSheet_range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            workSheet_range.Interior.Color = Excel.XlRgbColor.rgbLavender;


            _xlWorkSheet.Cells[2, colNum] = "#";
            _xlWorkSheet.Cells[2, colNum + 1] = operator_type;
            _xlWorkSheet.Cells[2, colNum + 2] = "Total Call";
            _xlWorkSheet.Cells[2, colNum + 3] = "Total Minutes";
            _xlWorkSheet.Cells[2, 1].EntireRow.Font.Bold = true;

            int rowNum = 2;
            for (i = 0; i <= ds1.Tables[0].Rows.Count - 1; i++)
            {
                rowNum++;
                _xlWorkSheet.Cells[rowNum, colNum] = i + 1;
                for (j = 0; j <= ds1.Tables[0].Columns.Count - 1; j++)
                {
                    data = ds1.Tables[0].Rows[i].ItemArray[j].ToString();
                    _xlWorkSheet.Cells[rowNum, j + colNum + 1] = data;
                    if (j == 1)
                    {
                        _xlWorkSheet.Cells[rowNum, j + colNum + 1].NumberFormat = "#,##0";
                    }
                    else if (j == 2)
                    {
                        _xlWorkSheet.Cells[rowNum, j + colNum + 1].NumberFormat = "#,##0.00";
                    }
                }
            }
            rowNum++;

            char C, D;
            C = Convert.ToChar(colNum + 1 + 65);
            D = Convert.ToChar(colNum + 2 + 65);

            _xlWorkSheet.Cells[rowNum, colNum + 1] = "Total";
            _xlWorkSheet.Cells[rowNum, colNum + 1].Font.Bold = true;

            String str = "=SUM(" + C + "3" + ":" + C + (rowNum - 1).ToString() + ")";
            _xlWorkSheet.Cells[rowNum, colNum + 2] = str;
            _xlWorkSheet.Cells[rowNum, colNum + 2].NumberFormat = "#,##0";
            _xlWorkSheet.Cells[rowNum, colNum + 2].Font.Bold = true;

            str = "=SUM(" + D + "3" + ":" + D + (rowNum - 1).ToString() + ")";
            _xlWorkSheet.Cells[rowNum, colNum + 3] = str;
            _xlWorkSheet.Cells[rowNum, colNum + 3].NumberFormat = "#,##0.00";
            _xlWorkSheet.Cells[rowNum, colNum + 3].Font.Bold = true;

            //_xlWorkSheet.Cells[rowNum, 1].EntireRow.Font.Bold = true; //it is problematic when mismatch row num



            Excel.Range _rangeA = (Excel.Range)_xlWorkSheet.Cells[1, starting_num];
            Excel.Range _rangeB = (Excel.Range)_xlWorkSheet.Cells[rowNum, colNum + 3];
            _xlWorkSheet.get_Range(_rangeA, _rangeB).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


            return rowNum;
        }

        public int ExcelPlot_without_Compare(ref Excel.Worksheet _xlWorkSheet, ref DataSet ds1, int rowNum, string name, string param1, string param2)
        {
            int starting_num = rowNum;
            string data;
            int i, j;

            //double d;

            //var f = new NumberFormatInfo { NumberGroupSeparator = "," };

            if(name != "")
            {
                _xlWorkSheet.Cells[rowNum, 1] = name;
                _xlWorkSheet.Cells[rowNum, 1].EntireRow.Font.Bold = true;
                Excel.Range _range1 = (Excel.Range)_xlWorkSheet.Cells[rowNum, 1];
                Excel.Range _range2;

                if (param2 == "") _range2 = (Excel.Range)_xlWorkSheet.Cells[rowNum, 4];
                else _range2 = (Excel.Range)_xlWorkSheet.Cells[rowNum, 5];


                Excel.Range workSheet_range = _xlWorkSheet.get_Range(_range1, _range2);
                workSheet_range.Merge();
                workSheet_range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                workSheet_range.Interior.Color = Excel.XlRgbColor.rgbLavender;
                rowNum++;
            }
            

            _xlWorkSheet.Cells[rowNum, 1] = "#";
            _xlWorkSheet.Cells[rowNum, 2] = param1;
            _xlWorkSheet.Cells[rowNum, 3] = "Total Call";
            _xlWorkSheet.Cells[rowNum, 4] = "Total Minutes";
            if(param2 != "") _xlWorkSheet.Cells[rowNum, 5] = param2;
            _xlWorkSheet.Cells[rowNum, 1].EntireRow.Font.Bold = true;
            //xlWorkSheet.get_Range("A1", "C11").Borders.Color = System.Drawing.Black.ToArgb();

            for (i = 0; i <= ds1.Tables[0].Rows.Count - 1; i++)
            {
                rowNum++;
                _xlWorkSheet.Cells[rowNum, 1] = i + 1;
                for (j = 0; j <= ds1.Tables[0].Columns.Count - 1; j++)
                {
                    data = ds1.Tables[0].Rows[i].ItemArray[j].ToString();
                    _xlWorkSheet.Cells[rowNum, j + 2] = data;
                    if (j == 1)
                    {
                        _xlWorkSheet.Cells[rowNum, j + 2].NumberFormat = "#,##0";
                    }
                    else if (j == 2)
                    {
                        _xlWorkSheet.Cells[rowNum, j + 2].NumberFormat = "#,##0.00";
                    }
                }
            }

            if (name != "")
            {
                rowNum++;
                _xlWorkSheet.Cells[rowNum, 2] = "Total";
                _xlWorkSheet.Cells[rowNum, 3] = "=SUM(C" + (starting_num + 2).ToString() + ":C" + (rowNum - 1).ToString() + ")";
                _xlWorkSheet.Cells[rowNum, 3].NumberFormat = "#,##0";
                _xlWorkSheet.Cells[rowNum, 4] = "=SUM(D" + (starting_num + 2).ToString() + ":D" + (rowNum - 1).ToString() + ")";
                _xlWorkSheet.Cells[rowNum, 4].NumberFormat = "#,##0.00";
                _xlWorkSheet.Cells[rowNum, 1].EntireRow.Font.Bold = true;
            }

            Excel.Range _rangeA = (Excel.Range)_xlWorkSheet.Cells[starting_num, 1];

            Excel.Range _rangeB;

            if (param2 == "") _rangeB = (Excel.Range)_xlWorkSheet.Cells[rowNum, 4];
            else _rangeB = (Excel.Range)_xlWorkSheet.Cells[rowNum, 5];

            _xlWorkSheet.get_Range(_rangeA, _rangeB).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            _xlWorkSheet.get_Range("A1", "Z100").Columns.AutoFit();
            _xlWorkSheet.get_Range("A1", "Z100").Rows.AutoFit();

            rowNum++;
            return rowNum;
        }
        public int ExcelPlot_Hourly_Sheet_IDD(ref Excel.Worksheet _xlWorkSheet, ref DataSet ds1, int rowNum, string name, int divisor, string operator_type)
        {
            int starting_num = rowNum;
            string data;
            int i, j;

            double d;

            var f = new NumberFormatInfo { NumberGroupSeparator = "," };

            _xlWorkSheet.Cells[rowNum, 1] = name;
            _xlWorkSheet.Cells[rowNum, 1].EntireRow.Font.Bold = true;

            Excel.Range _range1 = (Excel.Range)_xlWorkSheet.Cells[rowNum, 1];
            Excel.Range _range2 = (Excel.Range)_xlWorkSheet.Cells[rowNum, 27];
            Excel.Range workSheet_range = _xlWorkSheet.get_Range(_range1, _range2);
            workSheet_range.Merge();
            workSheet_range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            workSheet_range.Interior.Color = Excel.XlRgbColor.rgbLavender;

            rowNum++;

            _xlWorkSheet.Cells[rowNum, 1] = "#";
            _xlWorkSheet.Cells[rowNum, 2] = operator_type;
            _xlWorkSheet.Cells[rowNum, 2].EntireColumn.Font.Bold = true;
            _xlWorkSheet.Cells[rowNum, 1].EntireRow.Font.Bold = true;

            for (j = 0; j <= ds1.Tables[0].Columns.Count - 1; j++)
            {
                _xlWorkSheet.Cells[rowNum, j + 3] = j;
                _xlWorkSheet.Cells[rowNum, j + 3].NumberFormat = "00";
            }
            _xlWorkSheet.Cells[rowNum, 27] = "Total Minutes";
            _xlWorkSheet.Cells[rowNum, 27].EntireColumn.Font.Bold = true;

            for (i = 0; i <= ds1.Tables[0].Rows.Count - 1; i++)
            {
                rowNum++;
                _xlWorkSheet.Cells[rowNum, 1] = i + 1;
                for (j = 0; j <= ds1.Tables[0].Columns.Count - 1; j++)
                {
                    data = ds1.Tables[0].Rows[i].ItemArray[j].ToString();
                    if (j > 0 && data != "")
                    {
                        d = Convert.ToDouble(data);
                        d = d / divisor;
                        var s = d.ToString("n", f);
                        data = s.ToString();
                    }
                    _xlWorkSheet.Cells[rowNum, j + 2] = data;
                }
            }

            rowNum++;
            _xlWorkSheet.Cells[rowNum, 2] = "Total Min";
            _xlWorkSheet.Cells[rowNum, 1].EntireRow.Font.Bold = true;


            char C;

            for (j = 0; j <= ds1.Tables[0].Columns.Count - 2; j++)
            {
                C = Convert.ToChar(j + 2 + 65);
                _xlWorkSheet.Cells[rowNum, j + 3] = "=SUM(" + C + (starting_num + 2).ToString() + ":" + C + (rowNum - 1).ToString() + ")";
                _xlWorkSheet.Cells[rowNum, j + 3].NumberFormat = "#,##0.00";
            }

            for (i = 0; i <= ds1.Tables[0].Rows.Count - 1; i++)
            {
                _xlWorkSheet.Cells[starting_num + i + 2, 27] = "=SUM(C" + (starting_num + i + 2).ToString() + ":Z" + (starting_num + i + 2).ToString() + ")";
                _xlWorkSheet.Cells[starting_num + i + 2, 27].NumberFormat = "#,##0.00";
            }

            _xlWorkSheet.Cells[starting_num + i + 2, 27] = "=SUM(AA" + (starting_num + 2).ToString() + ":AA" + (starting_num + i + 1).ToString() + ")";
            _xlWorkSheet.Cells[starting_num + i + 2, 27].NumberFormat = "#,##0.00";

            Excel.Range _rangeA = (Excel.Range)_xlWorkSheet.Cells[starting_num, 1];
            Excel.Range _rangeB = (Excel.Range)_xlWorkSheet.Cells[starting_num + i + 2, 27];
            _xlWorkSheet.get_Range(_rangeA, _rangeB).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            _xlWorkSheet.get_Range("A1", "AZ100").Columns.AutoFit();
            _xlWorkSheet.get_Range("A1", "AZ100").Rows.AutoFit();
            rowNum++;
            return rowNum;
        }


        public int ExcelPlot_ICX_Local(ref Excel.Worksheet _xlWorkSheet, ref DataSet ds1, int rowNum, int divisor, IList<String> _trunkout_operator)
        {
            int starting_num = rowNum;
            string data;
            int i, j;

            double d;

            var f = new NumberFormatInfo { NumberGroupSeparator = "," };

            _xlWorkSheet.Cells[rowNum, 3] = "To (Called) Operators";
            _xlWorkSheet.Cells[rowNum, 3].EntireRow.Font.Bold = true;

            Excel.Range _range1 = (Excel.Range)_xlWorkSheet.Cells[rowNum, 3];
            Excel.Range _range2 = (Excel.Range)_xlWorkSheet.Cells[rowNum, _trunkout_operator.Count + 3];
            Excel.Range workSheet_range = _xlWorkSheet.get_Range(_range1, _range2);
            workSheet_range.Merge();
            workSheet_range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            workSheet_range.Interior.Color = Excel.XlRgbColor.rgbLavender;

            rowNum++;

            _xlWorkSheet.Cells[rowNum, 1] = "#";
            _xlWorkSheet.Cells[rowNum, 2] = "From (Calling) Operators";
            _xlWorkSheet.Cells[rowNum, 2].Interior.Color = Excel.XlRgbColor.rgbLavender;
            _xlWorkSheet.Cells[rowNum, 2].EntireColumn.Font.Bold = true;
            _xlWorkSheet.Cells[rowNum, 1].EntireRow.Font.Bold = true;

            for (j = 0; j < ds1.Tables[0].Columns.Count - 1; j++)
            {
                _xlWorkSheet.Cells[rowNum, j + 3] = _trunkout_operator[j];
            }
            _xlWorkSheet.Cells[rowNum, _trunkout_operator.Count + 3] = "Total Minutes";
            _xlWorkSheet.Cells[rowNum, _trunkout_operator.Count + 3].EntireColumn.Font.Bold = true;

            for (i = 0; i <= ds1.Tables[0].Rows.Count - 1; i++)
            {
                rowNum++;
                _xlWorkSheet.Cells[rowNum, 1] = i + 1;
                for (j = 0; j <= ds1.Tables[0].Columns.Count - 1; j++)
                {
                    data = ds1.Tables[0].Rows[i].ItemArray[j].ToString();
                    if (j > 0 && data != "")
                    {
                        d = Convert.ToDouble(data);
                        d = d / divisor;
                        var s = d.ToString("n", f);
                        data = s.ToString();
                    }
                    _xlWorkSheet.Cells[rowNum, j + 2] = data;
                }
            }

            rowNum++;
            _xlWorkSheet.Cells[rowNum, 2] = "Total Min";
            _xlWorkSheet.Cells[rowNum, 1].EntireRow.Font.Bold = true;


            
            string _Char = "";

            for (j = 0; j <= ds1.Tables[0].Columns.Count - 2; j++)
            {
                _Char = Find_Right_ColumnNumber(j + 2);
                _xlWorkSheet.Cells[rowNum, j + 3] = "=SUM(" + _Char + (starting_num + 2).ToString() + ":" + _Char + (rowNum - 1).ToString() + ")";
                _xlWorkSheet.Cells[rowNum, j + 3].NumberFormat = "#,##0.00";
            }


            _Char = Find_Right_ColumnNumber(_trunkout_operator.Count + 1);
            for (i = 0; i <= ds1.Tables[0].Rows.Count - 1; i++)
            {
                _xlWorkSheet.Cells[starting_num + i + 2, _trunkout_operator.Count + 3] = "=SUM(C" + (starting_num + i + 2).ToString() + ":" + _Char + (starting_num + i + 2).ToString() + ")";
                _xlWorkSheet.Cells[starting_num + i + 2, _trunkout_operator.Count + 3].NumberFormat = "#,##0.00";
            }

            _Char = Find_Right_ColumnNumber(_trunkout_operator.Count + 2);
            _xlWorkSheet.Cells[starting_num + i + 2, _trunkout_operator.Count + 3] = "=SUM(" + _Char + (starting_num + 2).ToString() + ":" + _Char + (starting_num + i + 1).ToString() + ")";
            _xlWorkSheet.Cells[starting_num + i + 2, _trunkout_operator.Count + 3].NumberFormat = "#,##0.00";

            Excel.Range _rangeA = (Excel.Range)_xlWorkSheet.Cells[starting_num, 1];
            Excel.Range _rangeB = (Excel.Range)_xlWorkSheet.Cells[starting_num + i + 2, _trunkout_operator.Count + 3];
            _xlWorkSheet.get_Range(_rangeA, _rangeB).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            _xlWorkSheet.get_Range("A1", "AZ100").Columns.AutoFit();
            _xlWorkSheet.get_Range("A1", "AZ100").Rows.AutoFit();
            rowNum++;
            return rowNum;
        }

        private string Find_Right_ColumnNumber(int i)
        {
            char C;
            string _Char = "";
            if (i + 65 > 90)
            {
                int num = i + 65 - 90;
                C = Convert.ToChar(num + 65 - 1);
                _Char = "A" + C.ToString();
            }
            else
            {
                C = Convert.ToChar(i + 65);
                _Char = C.ToString();
            }

            return _Char;
        }

        public int Difference_Entry_in_Excel(ref Excel.Worksheet _xlWorkSheet, int rowNum1, int rowNum2, double General_acceptance_diff, string greater)
        {
            int rn = 0;
            _xlWorkSheet.Cells[2, 11] = "Call Diff";
            _xlWorkSheet.Cells[2, 12] = "Minute Diff (in %)";
            if (rowNum1 == rowNum2)
            {
                for (int i = 3; i <= rowNum1; i++)
                {
                    if (greater == "LEFT")
                    {
                        _xlWorkSheet.Cells[i, 11] = "=C" + i.ToString() + "-H" + i.ToString();
                        _xlWorkSheet.Cells[i, 12] = "=(D" + i.ToString() + "-I" + i.ToString() + ")/D" + i.ToString() + "*100";
                    }
                    else if (greater == "RIGHT")
                    {
                        _xlWorkSheet.Cells[i, 11] = "=H" + i.ToString() + "-C" + i.ToString();
                        _xlWorkSheet.Cells[i, 12] = "=(I" + i.ToString() + "-D" + i.ToString() + ")/D" + i.ToString() + "*100";
                    }

                    _xlWorkSheet.Cells[i, 11].NumberFormat = "#,##0";
                    _xlWorkSheet.Cells[i, 12].NumberFormat = "#,##0.0000";

                    double d = Convert.ToDouble(_xlWorkSheet.Cells[i, 12].Value2.ToString());
                    if (d > General_acceptance_diff || d < General_acceptance_diff * (-1))
                    {
                        _xlWorkSheet.Cells[i, 12].EntireRow.Font.Color = Excel.XlRgbColor.rgbRed;
                    }
                }
                _xlWorkSheet.Cells[rowNum1, 11].Font.Bold = true;
                _xlWorkSheet.Cells[rowNum1, 12].Font.Bold = true;
            }
            else
            {
                if (greater == "LEFT")
                    _xlWorkSheet.Cells[3, 13] = "Due to mismatch of row numbers, auto difference is not possible. Please do it manually. At first match the cells then use this formula: (C-G)/C*100.";
                else if (greater == "RIGHT")
                    _xlWorkSheet.Cells[3, 13] = "Due to mismatch of row numbers, auto difference is not possible. Please do it manually. At first match the cells then use this formula: (G-C)/C*100.";

                _xlWorkSheet.Cells[3, 13].Font.Color = Excel.XlRgbColor.rgbRed;
            }

            Excel.Range _rangeA = (Excel.Range)_xlWorkSheet.Cells[2, 11];
            Excel.Range _rangeB;

            if (rowNum1 > rowNum2)
            {
                _rangeB = (Excel.Range)_xlWorkSheet.Cells[rowNum1, 12];
                rn = rowNum1;
            }
            else
            {
                _rangeB = (Excel.Range)_xlWorkSheet.Cells[rowNum2, 12];
                rn = rowNum2;
            }

            _xlWorkSheet.get_Range(_rangeA, _rangeB).Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            _xlWorkSheet.get_Range("A1", "Z100").Columns.AutoFit();
            _xlWorkSheet.get_Range("A1", "Z100").Rows.AutoFit();
            return rn;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;


using Excel = Microsoft.Office.Interop.Excel;






namespace ExcelReplace
{

    class TranslateItem
    {
        public TranslateItem()
        {
            Oringin = "";
            Translate = "";
        }

        public string Oringin;
        public string Translate;
    }


    class Program
    {
        static void Main(string[] args)
        {
            MergeExcel(@"F:\x2_vn\translate\SG145_script edit on 9.07.xls.xlsx", @"F:\x2_vn\translate\script.xls");
        }

        private static void MergeExcel(string source, string target)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook source_file = app.Workbooks.Open(source);
            Excel.Workbook targe_file = app.Workbooks.Open(target);

            Excel.Worksheet sheet1 = source_file.Sheets[1];
            Excel.Worksheet sheet2 = targe_file.Sheets[1];

            int last_row1 = sheet1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int last_row2 = sheet2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            Array values1 = (System.Array)sheet1.get_Range("A1", string.Format("B{0}", last_row1)).Cells.Value;
            Array values2 = (System.Array)sheet2.get_Range("A1", string.Format("B{0}", last_row2)).Cells.Value;

            List<TranslateItem> trans1 = new List<TranslateItem>();
            for (int i = 1; i <= last_row1; ++i)
            {
                TranslateItem tran = new TranslateItem();
                if (values1.GetValue(i, 1) != null)
                {
                    tran.Oringin = values1.GetValue(i, 1).ToString();
                }
                if (values1.GetValue(i, 2) != null)
                {
                    tran.Translate = values1.GetValue(i, 2).ToString();
                }
                trans1.Add(tran);
            }

            List<TranslateItem> trans2 = new List<TranslateItem>();
            for (int i = 1; i <= last_row2; ++i)
            {
                for (int j = 0; j < trans1.Count; ++j)
                {
                    if (values2.GetValue(i, 1) != null && (values2.GetValue(i, 1).ToString().Trim(new char[] { ' ', '\t' }) == trans1[j].Oringin.Trim(new char[] { ' ', '\t' })))
                    {
                        if (values2.GetValue(i, 2) == null)
                        {
                            sheet2.Cells[i, 2] = trans1[j].Translate;
                            continue;
                        }
                        else
                        {
                            string value = values2.GetValue(i, 2).ToString();
                            if (value != trans1[j].Translate && !String.IsNullOrEmpty(trans1[j].Translate))
                            {
                                sheet2.Cells[i, 2] = trans1[j].Translate;
                                continue;
                            }
                        }
                    }
                }
            }
        }

        private static void CompareFile(string file1, string file2)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = false;

            Excel.Workbook book1 = app.Workbooks.Open(file1);
            Excel.Workbook book2 = app.Workbooks.Open(file2);

            Excel.Worksheet sheet1 = book1.Sheets[1];
            Excel.Worksheet sheet2 = book2.Sheets[1];

            int last_row1 = sheet1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int last_row2 = sheet2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;


            Array values1 = (System.Array)sheet1.get_Range("A1", string.Format("A{0}", last_row1)).Cells.Value;

            Array values2 = (System.Array)sheet2.get_Range("A1", string.Format("A{0}", last_row2)).Cells.Value;

            List<string> strings1 = new List<string>();
            for (int i = 1; i <= last_row1; ++i)
            {
                if (values1.GetValue(i, 1) != null)
                {
                    strings1.Add(values1.GetValue(i, 1).ToString());
                }
            }

            List<string> strings2 = new List<string>();
            for (int i = 1; i <= last_row2; ++i)
            {
                if (values2.GetValue(i, 1) != null)
                {
                    strings2.Add(values2.GetValue(i, 1).ToString());
                }
            }


            List<string> same = new List<string>();
            List<string> doubles = new List<string>();
            for (int i = 0; i < strings1.Count; ++i)
            {
                if (!same.Contains(strings1[i]))
                {
                    same.Add(strings1[i]);
                }
                else
                {
                    doubles.Add(strings1[i]);
                }
            }



            List<string> difs = new List<string>();
            for (int i = 0; i < strings1.Count; ++i)
            {
                bool is_exists = false;
                for (int j = 0; j < strings2.Count; ++j)
                {
                    if (strings1[i] == strings2[j])
                    {
                        is_exists = true;
                        break;
                    }
                }
                if (!is_exists)
                {
                    difs.Add(strings1[i]);
                }
            }
        }

        struct ErrorInfo
        {
            public int line;
            public string origin;
            public string translate;
            public string new_translate;
        }

        private static void FindDifference(string file1, string file2)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = false;
            Excel.Workbook oldBook = app.Workbooks.Open(file1);
            Excel.Worksheet old_sheet = (Excel.Worksheet)oldBook.Sheets[1]; // Explicit cast is not required here

            int lastRow = old_sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            List<ErrorInfo> infos = new List<ErrorInfo>();

            Array values = (System.Array)old_sheet.get_Range("A1", string.Format("A{0}", lastRow)).Cells.Value;
            for (int i = 1; i <= lastRow; ++i)
            {
                ErrorInfo info = new ErrorInfo();
                info.line = i;

                if (values.GetValue(i, 1) != null)
                {
                    info.origin = values.GetValue(i, 1).ToString();
                }

                infos.Add(info);
            }


            values = (System.Array)old_sheet.get_Range("B1", string.Format("B{0}", lastRow)).Cells.Value;
            for (int i = 1; i <= lastRow; ++i)
            {
                if (values.GetValue(i, 1) != null)
                {
                    ErrorInfo info = infos[i];
                    info.translate = values.GetValue(i, 1).ToString();
                    infos[i] = info;
                }
            }


            Excel.Workbook newBook = app.Workbooks.Open(file2);
            Excel.Worksheet new_sheet = (Excel.Worksheet)newBook.Sheets[1]; // Explicit cast is not required here

            lastRow = new_sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;


            values = (System.Array)new_sheet.get_Range("B1", string.Format("B{0}", lastRow)).Cells.Value;
            for (int i = 1; i <= lastRow; ++i)
            {
                if (values.GetValue(i, 1) != null)
                {
                    ErrorInfo info = infos[i];
                    info.new_translate = values.GetValue(i, 1).ToString();
                    infos[i] = info;
                }
            }



            List<ErrorInfo> need_to_fix = new List<ErrorInfo>();

            List<int> lines = new List<int>();
            for (int i = 0; i < infos.Count; ++i)
            {
                if (infos[i].translate != infos[i].new_translate)
                {
                    lines.Add(infos[i].line);
                    need_to_fix.Add(infos[i]);
                }
            }

            Excel.Workbook book = app.Workbooks.Open(@"F:\x2_vn\translate\SG145_script edit on 9.05.xls.xlsx");
            Excel.Worksheet sheet = (Excel.Worksheet)book.Sheets[1]; // Explicit cast is not required here


            for (int i = 0; i < need_to_fix.Count; ++i)
            {
                sheet.Cells[need_to_fix[i].line - 1, 2] = need_to_fix[i].new_translate;
            }

            book.Save();


            oldBook.Close(false);
            newBook.Close(false);
            book.Close(true);

        }

        private static void FixError(string file)
        {
            var app = new Excel.Application();
            app.Visible = false;
            var book = app.Workbooks.Open(file);
            var sheet = (Excel.Worksheet)book.Sheets[1]; // Explicit cast is not required here

            int lastRow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            List<string> datas = new List<string>();
            List<string> after = new List<string>();
            Array values = (System.Array)sheet.get_Range("B1", string.Format("B{0}", lastRow)).Cells.Value;

            Regex regex = new Regex(@"(\|)(\d{1,5}), (\d{1,5})");
            for (int i = 1; i <= lastRow; ++i)
            {
                if (values.GetValue(i, 1) != null)
                {
                    string s = values.GetValue(i, 1).ToString();
                    if (regex.IsMatch(s))
                    {
                        datas.Add(s);
                        string new_string = regex.Replace(s, "$1$2,$3");
                        after.Add(new_string);
                        sheet.Cells[i, 2] = new_string;
                    }
                }
            }
            book.Save();

            values = (System.Array)sheet.get_Range("B1", string.Format("B{0}", lastRow)).Cells.Value;
            regex = new Regex(@"event: ");
            for (int i = 1; i <= lastRow; ++i)
            {
                if (values.GetValue(i, 1) != null)
                {
                    string s = values.GetValue(i, 1).ToString();
                    if (regex.IsMatch(s))
                    {
                        datas.Add(s);
                        string new_string = regex.Replace(s, "event:");
                        after.Add(new_string);
                        sheet.Cells[i, 2] = new_string;
                    }
                }
            }
            book.Save();
        }
    }
}

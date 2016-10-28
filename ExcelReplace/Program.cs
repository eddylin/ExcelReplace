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
    struct TranslateItem
    {
        public string Oringin;
        public string Translate;
    }

    struct ModifyTranslateItem
    {
        public string Oringin;
        public string Translate;
        public string NewTranslate;
    }

    struct ErrorInfo
    {
        public int line;
        public string origin;
        public string translate;
        public string new_translate;
    }

    public class TranslateItemWithLine
    {
        public TranslateItemWithLine()
        {
            Line = 0;
            Oringin = "";
            Translate = "";
        }

        public int Line;
        public string Oringin;
        public string Translate;
    }

    public class Duplicate
    {
        public string Origin;
        public List<TranslateItemWithLine> Lines;

        public Duplicate()
        {
            Origin = "";
            Lines = new List<TranslateItemWithLine>();
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            //RemoveDuplicates(@"H:\x2_vn\translate\script.xls", @"H:\x2_vn\translate\script_remove.xls");
            FindNotTranslateItems(@"H:\x2_vn\translate\server.xls", @"H:\x2_vn\translate\server_untranslate.xls");
        }

        private static void MergeExcel(string source, string target)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = false;
            Excel.Workbook source_file = app.Workbooks.Open(source);
            Excel.Workbook target_file = app.Workbooks.Open(target);

            Excel.Worksheet source_sheet = source_file.Sheets[1];
            Excel.Worksheet target_sheet = target_file.Sheets[1];

            int last_row1 = source_sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int last_row2 = target_sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            Array source_values = (System.Array)source_sheet.get_Range("A1", string.Format("B{0}", last_row1)).Cells.Value;
            Array target_values = (System.Array)target_sheet.get_Range("A1", string.Format("B{0}", last_row2)).Cells.Value;


            char[] strip_chars = new char[] { '\r', '\n', '\t', ' ' };

            List<TranslateItem> source_trans = new List<TranslateItem>();
            for (int i = 1; i <= last_row1; ++i)
            {
                TranslateItem tran = new TranslateItem();
                if (source_values.GetValue(i, 1) != null)
                {
                    tran.Oringin = source_values.GetValue(i, 1).ToString().Trim(strip_chars);
                }
                if (source_values.GetValue(i, 2) != null)
                {
                    tran.Translate = source_values.GetValue(i, 2).ToString().Trim(strip_chars);
                }
                source_trans.Add(tran);
            }

            List<TranslateItem> target_trans = new List<TranslateItem>();
            for (int i = 1; i <= last_row2; ++i)
            {
                Console.WriteLine("i: " + i);
                for (int j = 0; j < source_trans.Count; ++j)
                {
                    if (target_values.GetValue(i, 1) != null && (target_values.GetValue(i, 1).ToString().Trim(strip_chars) == source_trans[j].Oringin.Trim(strip_chars)))
                    {
                        try
                        {
                            target_sheet.Cells[i, 1] = target_values.GetValue(i, 1).ToString().Trim(strip_chars);
                            target_sheet.Cells[i, 2] = source_trans[j].Translate.ToString().Trim(strip_chars);
                        }
                        catch (Exception e)
                        {
                            var e2 = e;
                        }
                        break;
                    }
                }
            }

            target_file.Save();
            target_file.Close();
            source_file.Close();
        }

        private static void CompareFile(string file1, string file2, string output_file)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = false;

            Excel.Workbook book1 = app.Workbooks.Open(file1);
            Excel.Workbook book2 = app.Workbooks.Open(file2);

            Excel.Worksheet sheet1 = book1.Sheets[1];
            Excel.Worksheet sheet2 = book2.Sheets[1];

            int last_row1 = sheet1.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            int last_row2 = sheet2.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;


            Array values1 = (Array)sheet1.get_Range("A1", string.Format("B{0}", last_row1)).Cells.Value;
            Array values2 = (Array)sheet2.get_Range("A1", string.Format("B{0}", last_row2)).Cells.Value;


            TranslateItem[] items1 = new TranslateItem[last_row1];
            TranslateItem[] items2 = new TranslateItem[last_row2];

            for (int i = 1; i <= last_row1; ++i)
            {
                TranslateItem item = new TranslateItem();
                item.Oringin = (values1.GetValue(i, 1) ?? "").ToString();
                item.Translate = (values1.GetValue(i, 2) ?? "").ToString();
                items1[i - 1] = item;
            }

            for (int i = 1; i <= last_row2; ++i)
            {
                TranslateItem item = new TranslateItem();
                item.Oringin = (values2.GetValue(i, 1) ?? "").ToString();
                item.Translate = (values2.GetValue(i, 2) ?? "").ToString();
                items2[i - 1] = item;
            }


            List<TranslateItem> deleted_items = new List<TranslateItem>();
            List<TranslateItem> added_items = new List<TranslateItem>();
            List<ModifyTranslateItem> modified_items = new List<ModifyTranslateItem>();

            foreach (var item1 in items1)
            {
                bool is_deleted = true;
                foreach (var item2 in items2)
                {
                    if (item1.Oringin == item2.Oringin)
                    {
                        is_deleted = false;
                        if (item1.Translate != item2.Translate)
                        {
                            // modified
                            ModifyTranslateItem modified_item = new ModifyTranslateItem();
                            modified_item.Oringin = item1.Oringin;
                            modified_item.Translate = item1.Translate;
                            modified_item.NewTranslate = item2.Translate;
                            modified_items.Add(modified_item);
                        }
                    }
                }

                if (is_deleted)
                {
                    // deleted
                    deleted_items.Add(item1);
                }
            }

            foreach (var item2 in items2)
            {
                bool is_added = true;
                foreach (var item1 in items1)
                {
                    if (item2.Oringin == item1.Oringin)
                    {
                        is_added = false;
                        break;
                    }
                }
                if (is_added)
                {
                    // added
                    added_items.Add(item2);
                }
            }

            if (File.Exists(output_file)) File.Delete(output_file);

            var output_workbook = app.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            if (!Directory.Exists(Path.GetDirectoryName(output_file)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(output_file));
            }
            output_workbook.Sheets.Add();
            output_workbook.Sheets.Add();


            Excel.Worksheet sheet = output_workbook.Sheets[1];
            for (int i = 1; i <= modified_items.Count; ++i)
            {
                sheet.Cells[i, 1] = modified_items[i - 1].Oringin;
                sheet.Cells[i, 2] = modified_items[i - 1].Translate;
                sheet.Cells[i, 3] = modified_items[i - 1].NewTranslate;
            }

            sheet = output_workbook.Sheets[2];
            for (int i = 1; i <= added_items.Count; ++i)
            {
                sheet.Cells[i, 1] = added_items[i - 1].Oringin;
                sheet.Cells[i, 2] = added_items[i - 1].Translate;
            }

            sheet = output_workbook.Sheets[3];
            for (int i = 1; i <= deleted_items.Count; ++i)
            {
                sheet.Cells[i, 1] = deleted_items[i - 1].Oringin;
                sheet.Cells[i, 2] = deleted_items[i - 1].Translate;
            }

            output_workbook.SaveAs(output_file);

            output_workbook.Close();





        }

        private static void FindDifference(string file1, string file2, string output_file)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = false;
            Excel.Workbook oldBook = app.Workbooks.Open(file1);
            Excel.Worksheet old_sheet = (Excel.Worksheet)oldBook.Sheets[1]; // Explicit cast is not required here

            int lastRow = old_sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            List<ErrorInfo> infos = new List<ErrorInfo>();

            Array values = (Array)old_sheet.get_Range("A1", string.Format("A{0}", lastRow)).Cells.Value;
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


            values = (Array)old_sheet.get_Range("B1", string.Format("B{0}", lastRow)).Cells.Value;
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

        private static void FindDuplicate(string file)
        {
            var app = new Excel.Application();
            app.Visible = false;
            var book = app.Workbooks.Open(file);
            var sheet = (Excel.Worksheet)book.Sheets[1];

            int lastRow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            Array values = (System.Array)sheet.get_Range("A1", string.Format("B{0}", lastRow)).Cells.Value;

            List<string> strs = new List<string>();
            List<string> translates = new List<string>();
            for (int i = 1; i <= lastRow; ++i)
            {
                strs.Add((values.GetValue(i, 1) ?? "").ToString());
                translates.Add((values.GetValue(i, 2) ?? "").ToString());
            }

            Dictionary<string, Duplicate> duplicates = new Dictionary<string, Duplicate>();
            char[] trims = new char[] { ' ', '\t' };
            for (int i = 0; i < strs.Count - 1; ++i)
            {
                for (int j = i + 1; j < strs.Count; ++j)
                {
                    string s1 = strs[i].Trim(trims);
                    string s2 = strs[j].Trim(trims);
                    if (i != j && s1 == s2)
                    {
                        TranslateItemWithLine line_i = new TranslateItemWithLine();
                        line_i.Line = i + 1;
                        line_i.Oringin = strs[i];
                        line_i.Translate = translates[i];

                        TranslateItemWithLine line_j = new TranslateItemWithLine();
                        line_j.Line = j + 1;
                        line_j.Oringin = strs[j];
                        line_j.Translate = translates[j];

                        if (!duplicates.Keys.Contains(s1))
                        {
                            duplicates[s1] = new Duplicate();
                            duplicates[s1].Lines.Add(line_i);
                            duplicates[s1].Lines.Add(line_j);
                        }
                        else
                        {
                            duplicates[s1].Lines.Add(line_j);
                        }
                    }
                }
            }
        }

        private static void FindCaseIssues(string file, string output_file)
        {
            var app = new Excel.Application();
            app.Visible = false;
            var book = app.Workbooks.Open(file);
            var sheet = (Excel.Worksheet)book.Sheets[1];

            int lastRow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            Array values = (System.Array)sheet.get_Range("A1", string.Format("B{0}", lastRow)).Cells.Value;

            List<TranslateItem> items = new List<TranslateItem>();

            List<ErrorInfo> errors = new List<ErrorInfo>();
            Regex regex = new Regex(@"[a-zA-Z_0-9]+");
            for (int i = 1; i <= lastRow; ++i)
            {
                ErrorInfo item = new ErrorInfo();
                item.origin = (values.GetValue(i, 1) ?? "").ToString();
                item.translate = (values.GetValue(i, 2) ?? "").ToString();
                item.line = i;
                var matches = regex.Matches(item.origin);
                bool is_contain = true;
                foreach (var match in matches)
                {
                    string s = match.ToString();
                    if (!item.translate.Contains(s))
                    {
                        is_contain = false;
                        item.new_translate = (item.new_translate ?? "") + " " + s;
                    }
                }

                if (!is_contain)
                {
                    errors.Add(item);
                }
            }


            if (File.Exists(output_file)) File.Delete(output_file);

            var output_workbook = app.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            if (!Directory.Exists(Path.GetDirectoryName(output_file)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(output_file));
            }
            output_workbook.Sheets.Add();
            output_workbook.Sheets.Add();


            Excel.Worksheet output_sheet = output_workbook.Sheets[1];
            Excel.Range range = output_sheet.Range[output_sheet.Cells[1, 1], output_sheet.Cells[4, errors.Count]];

            for (int i = 1; i <= errors.Count; ++i)
            {
                range[i, 1] = errors[i - 1].line;
                range[i, 2] = errors[i - 1].origin;
                range[i, 3] = errors[i - 1].translate;
                range[i, 4] = errors[i - 1].new_translate;
            }

            output_workbook.SaveAs(output_file);
            output_workbook.Save();
            output_workbook.Close();
        }

        private static void FindNotTranslateItems(string file, string output_file)
        {
            var app = new Excel.Application();
            app.Visible = false;
            var book = app.Workbooks.Open(file);
            var sheet = (Excel.Worksheet)book.Sheets[1];

            int lastRow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            Array values = (System.Array)sheet.get_Range("A1", string.Format("B{0}", lastRow)).Cells.Value;

            List<ErrorInfo> errors = new List<ErrorInfo>();

            char[] strip_chars = new char[] { '\r', '\n', '\t', ' ' };
            for (int i = 1; i <= lastRow; ++i)
            {
                ErrorInfo item = new ErrorInfo();
                item.origin = (values.GetValue(i, 1) ?? "").ToString();
                item.translate = (values.GetValue(i, 2) ?? "").ToString();
                item.line = i;

                if (item.origin.Trim(strip_chars) == item.translate.Trim(strip_chars))
                {
                    errors.Add(item);
                }
            }


            if (File.Exists(output_file)) File.Delete(output_file);

            var output_workbook = app.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            if (!Directory.Exists(Path.GetDirectoryName(output_file)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(output_file));
            }
            output_workbook.Sheets.Add();
            output_workbook.Sheets.Add();


            Excel.Worksheet output_sheet = output_workbook.Sheets[1];
            Excel.Range range = output_sheet.Range[output_sheet.Cells[1, 1], output_sheet.Cells[3, errors.Count]];

            for (int i = 1; i <= errors.Count; ++i)
            {
                range[i, 1] = errors[i - 1].line;
                range[i, 2] = errors[i - 1].origin;
                range[i, 3] = errors[i - 1].translate;
                //range[i, 4] = errors[i - 1].new_translate;
            }

            output_workbook.SaveAs(output_file);
            output_workbook.Save();
            output_workbook.Close();

        }

        private static void CheckSpecialSymbol(string file, string output_file)
        {
            var app = new Excel.Application();
            app.Visible = false;
            var book = app.Workbooks.Open(file);
            var sheet = (Excel.Worksheet)book.Sheets[1];

            int lastRow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            Array values = (System.Array)sheet.get_Range("A1", string.Format("B{0}", lastRow)).Cells.Value;

            char[] symbols = new char[] { '[', ']', ',', '"', '\'', '<', '>', '-' };

            List<TranslateItem> items = new List<TranslateItem>();
            List<ErrorInfo> errors = new List<ErrorInfo>();
            for (int i = 1; i <= lastRow; ++i)
            {
                bool is_error = false;
                ErrorInfo item = new ErrorInfo();
                item.origin = (values.GetValue(i, 1) ?? "").ToString();
                item.translate = (values.GetValue(i, 2) ?? "").ToString();
                item.line = i;

                string issue = "";
                foreach (var c in symbols)
                {
                    int count1 = GetCharCount(item.origin, c);
                    int count2 = GetCharCount(item.translate, c);
                    if (count1 != count2)
                    {
                        issue = issue + " " + string.Format("c: {0} {1} {2}", c, count1, count2);
                        is_error = true;
                    }
                }

                if (is_error)
                {
                    item.new_translate = issue;
                    errors.Add(item);
                }
            }


            if (File.Exists(output_file)) File.Delete(output_file);

            var output_workbook = app.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            if (!Directory.Exists(Path.GetDirectoryName(output_file)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(output_file));
            }
            output_workbook.Sheets.Add();
            output_workbook.Sheets.Add();


            Excel.Worksheet output_sheet = output_workbook.Sheets[1];
            Excel.Range range = output_sheet.Range[output_sheet.Cells[1, 1], output_sheet.Cells[4, errors.Count]];

            for (int i = 1; i <= errors.Count; ++i)
            {
                range[i, 1] = errors[i - 1].line;
                range[i, 2] = errors[i - 1].origin;
                range[i, 3] = errors[i - 1].translate;
                range[i, 4] = errors[i - 1].new_translate;
            }

            output_workbook.SaveAs(output_file);
            output_workbook.Save();
            output_workbook.Close();
        }

        private static int GetCharCount(string s, char c)
        {
            int count = 0;
            foreach (var cha in s)
            {
                if (cha == c) ++count;
            }
            return count;
        }

        /// <summary>
        /// Remove duplicate in excel file. 
        /// </summary>
        /// <param name="source_excel">source excel</param>
        /// <param name="output_excel">output excel</param>
        private static void RemoveDuplicates(string source_excel, string output_excel)
        {
            #region argument check
            if (string.IsNullOrEmpty(source_excel))
            {
                Console.WriteLine("argument souce excel invalid, please check.");
                return;
            }
            if (!File.Exists(source_excel))
            {
                Console.WriteLine("source excel doese not exists, please check.");
                return;
            }
            #endregion

            var app = new Excel.Application();
            //app.Visible = false;
            var book = app.Workbooks.Open(source_excel);
            var sheet = (Excel.Worksheet)book.Sheets[1];

            int lastRow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            Array values = (Array)sheet.get_Range("A1", string.Format("B{0}", lastRow)).Cells.Value;
            
            Dictionary<string, string> translates = new Dictionary<string, string>();
            for (int i = 1; i <= lastRow; ++i)
            {
                string origin = (values.GetValue(i, 1) ?? "").ToString();
                string translate = (values.GetValue(i, 2) ?? "").ToString();
                translates[origin] = translate;
            }

            
            if (File.Exists(output_excel)) File.Delete(output_excel);

            var output_workbook = app.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            if (!Directory.Exists(Path.GetDirectoryName(output_excel)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(output_excel));
            }
            output_workbook.Sheets.Add();
            //output_workbook.Sheets.Add();


            var list = translates.ToList();

            list.Sort((a, b) => { return a.Key.Length.CompareTo(b.Key.Length); });

            Excel.Worksheet output_sheet = output_workbook.Sheets[1];
            for(int i = 1; i <= list.Count; ++i)
            {
                output_sheet.Cells[i, 1] = list[i - 1].Key;
                output_sheet.Cells[i, 2] = list[i - 1].Value;
            }
            output_workbook.SaveAs(output_excel);
            output_workbook.Close();
        }
    }
}

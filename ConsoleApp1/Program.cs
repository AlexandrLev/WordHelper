using System;
using System.Collections.Generic;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace ConsoleApp1
{
    class Program
    {
        public static Object templatePathexcel = Environment.CurrentDirectory + "\\" + "отчет.xlsm";
        static void Main(string[] args)
        {
            Word._Application wordApp;
            Word._Document wordDoc;
            // Создаём объект документа
            wordApp = new Word.Application();
            Object templatePathObj = Environment.CurrentDirectory+"\\"+ "ОЭБ.doc"; 
            //Делаем его видимым
            wordApp.Visible = true;

            //Object newTemplate = false;
            //Object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
            //Object visible = true;
            Object missingObj = System.Reflection.Missing.Value;
            Object trueObj = true;
            Object falseObj = false;
            //Создаем документ 1
            wordDoc = wordApp.Documents.Add(ref templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            var range = GetRangeForRaplace(wordDoc, "<Data>");
            range.Font.Size = 12;
            InputText(range, GetInfoFromExcel());
            string finalPath = Environment.CurrentDirectory + "\\" + "ОЭБ_" + DateTime.Now.ToString("dd.MM.yyyy") + ".docx";
            wordDoc.SaveAs(finalPath);

            //wordApp.ActiveDocument.Close();
            //wordApp.Quit();
        }

        static void InputText(Word.Range range, InfoObject info)
        {
            int index = 1; Object collapseDirection = Word.WdCollapseDirection.wdCollapseEnd;
            foreach (var item in info.infoList)
            {
                range.Font.Size = 12;
                range.Text = index++ +". ";
                range.Collapse(ref collapseDirection);
                foreach (var item2 in item)
                {
                    
                    range.Text = item2.Key + ": "+ item2.Value;
                    range.InsertParagraphAfter();
                    range.Font.Size = 12;
                    range.Collapse(ref collapseDirection);
                }
            }

            //range.Text = "Это текст заменит содержимое Range";
            //range.InsertParagraphAfter();
            //range.Collapse(ref collapseDirection);
            //range.Text = "Это текст заменит содержимое Range";
        }

        static InfoObject GetInfoFromExcel()
        {

            var infoObj = new InfoObject();
            var listInfo = new List<Dictionary<string, string>>();
            int rowIndex, columnIndex = 0;
            ExcelDocument excelDoc = new ExcelDocument((string)templatePathexcel);
            int usedRowsNum = excelDoc.GetUsedRowsNum();
            rowIndex = 7;
            columnIndex = 1;
            do
            {
                rowIndex++;
            } while (excelDoc.GetCellValue(rowIndex, 1)==null || excelDoc.GetCellValue(rowIndex, 1) == "№");
            for (; rowIndex <= usedRowsNum;)
            {
                if (excelDoc.GetCellValue(rowIndex, 1)!=null)
                {
                    var infoDir = new Dictionary<string, string>();
                    if (excelDoc.GetCellValue(rowIndex, 3) != null)
                    {


                        do
                        {
                            infoDir.Add(excelDoc.GetCellValue(rowIndex, 2), excelDoc.GetCellValue(rowIndex, 3));
                            rowIndex++;
                        } while (excelDoc.GetCellValue(rowIndex, 1) == null);
                        listInfo.Add(infoDir);
                    }else rowIndex++;
                }
                else rowIndex++;
            }


            excelDoc.Visible = true;
            excelDoc.Close();

            //listInfo.Add(new Dictionary<string, string>(){
            //    {"Франция", "1"},
            //    {"Германия", "1"},
            //    {"Великобритания", "1"}
            //});
            //listInfo.Add(new Dictionary<string, string>(){
            //    {"Германия", "1"},
            //    {"Мелкобритания", "1"}
            //});
            //listInfo.Add(new Dictionary<string, string>(){
            //    {"Франция", "1"},
            //    {"Германия", "1"},
            //    {"Великобритания", "1"}
            //});

            infoObj.infoList = listInfo;
            return infoObj;
        }
        
        static Word.Range GetRangeForRaplace(Word._Document _document, string stringToFind)
        {
            object stringToFindObj = stringToFind;
            Word.Range wordRange;
            bool rangeFound;


            //в цикле обходим все разделы документа, получаем Range, запускаем поиск
            // если поиск вернул true, он долже ужать Range до найденное строки, выходим и возвращаем Range
            // обходим все разделы документа
            for (int i = 1; i <= _document.Sections.Count; i++)
            {
                
                // берем всю секцию диапазоном
                wordRange = _document.Sections[i].Range;

                /*
                // Обходим редкий глюк в Find, ПРИЗНАННЫЙ MICROSOFT, метод Execute на некоторых машинах вылетает с ошибкой "Заглушке переданы неправильные данные / Stub received bad data"  Подробности: http://support.microsoft.com/default.aspx?scid=kb;en-us;313104
                // выполняем метод поиска и  замены обьекта диапазона ворд
                rangeFound = wordRange.Find.Execute(ref stringToFindObj, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing, ref wordMissing);
                */

                Word.Find wordFindObj = wordRange.Find;
                Object _missingObj = System.Reflection.Missing.Value;
                
                object[] wordFindParameters = new object[15] { stringToFindObj, _missingObj, _missingObj, _missingObj, _missingObj, _missingObj, _missingObj, _missingObj, _missingObj, _missingObj, _missingObj, _missingObj, _missingObj, _missingObj, _missingObj };

                rangeFound = (bool)wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

                if (rangeFound) { return wordRange; }
            }

            // если ничего не нашли, возвращаем null
            return null;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using OfficeOpenXml;

namespace ExcelExportTool
{
    class ExportExcel
    {
        //xlsx path list
        static List<string> pathLists = new List<string>();
        static List<string> pathTransLists = new List<string>();

        static ExportExcel _instance;

        static List<string> filterFolderList = new List<string>();//过滤不需要导出的文件夹

        static List<double> originIdList = new List<double>();
        static List<double> translateIdList = new List<double>();
        static List<double> translate2AddIdList = new List<double>();
        static List<double> translate2DelIdList = new List<double>();

        //static List<string> originIdList = new List<string>();
        //static List<string> translateIdList = new List<string>();
        //static List<string> translate2AddIdList = new List<string>();
        //static List<string> translate2DelIdList = new List<string>();

        public static ExportExcel Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new ExportExcel();
                }
                return _instance;
            }
        }

        public static bool Export(string rootPath, string[] transTypeList = null,string filterPath=null)
        {
            if (!Directory.Exists(rootPath)) {
                Form1.ShowTips("不存在" + rootPath + "此目录");
                return false;
            }
            transTypeList = transTypeList == null ? new string[] { "ch" } : transTypeList;
            rootPath = rootPath == "" ? @"E:\zzhx\trunk\data\ai" : rootPath;
            filterFolderList.Clear();
            if (filterPath != null && filterPath != "") {
                filterFolderList.Add(filterPath);
            }
            pathLists.Clear();
            TraverseFile(rootPath);

            string pattern = @"[\u4e00-\u9fa5]+";
            FileInfo excelFile;
            for (int i = 0; i < pathLists.Count; i++)
            {
                //excel origin table
                excelFile = new FileInfo(pathLists[i]);
                if (!excelFile.Exists)
                    continue;
                //record the col num which to be exported;be sure the first value is id
                List<int> collist = new List<int>();
                using (ExcelPackage package = new ExcelPackage(excelFile))//每一个原excel
                {
                    //the first sheet
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    int colStart = worksheet.Dimension.Start.Column;  //工作区开始列
                    int colEnd = worksheet.Dimension.End.Column;       //工作区结束列
                    int rowStart = worksheet.Dimension.Start.Row;       //工作区开始行号
                    int rowEnd = worksheet.Dimension.End.Row;       //工作区结束行号

                    int workRow = 2;//默认工作行
                    for (int row = rowStart; row <= 5; row++)
                    {//给5因为有些表的字段名下面还加了些描述等中文字符，故多遍历几次，以保证进入正文
                        for (int col = colStart; col <= colEnd; col++)
                        {
                            string text = worksheet.Cells[row, col].Text;
                            if (text == "id" || text == "ID")
                            {
                                if (!collist.Contains(col))
                                {
                                    collist.Add(col);
                                }
                                break;
                            }
                            //记录正文开始行
                            if (text.Contains("删"))
                            {
                                //删所在行的下一行为正文开始行
                                workRow = row + 1;
                                break;
                            }
                            MatchCollection mc = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
                            if (mc.Count > 0)
                            {
                                if (!collist.Contains(col))
                                {
                                    collist.Add(col);
                                }
                            }
                        }
                    }
                    //export
                    string pathTranslate = pathLists[i];
                    pathTranslate = pathTranslate.Substring(0, pathTranslate.Length - 5);//.xlsx
                    pathTranslate += "_translate.xlsx";
                    FileInfo newFile = new FileInfo(pathTranslate);
                    if (newFile.Exists)
                    {
                        //newFile.Delete();
                        using (ExcelPackage package2 = new ExcelPackage(newFile))//每一个translate excel
                        {
                            ExcelWorksheet workshee2 = package2.Workbook.Worksheets[1];
                            int transColStart;
                            int transColEnd;
                            int transRowStart;
                            int transRowEnd;
                            //比对改动后的原表与translate表的id的不同后添加删除
                            originIdList.Clear();
                            translateIdList.Clear();
                            translate2AddIdList.Clear();
                            translate2DelIdList.Clear();
                            //origin id
                            for (int row = workRow; row <= rowEnd; row++)
                            {
                                if (worksheet.Cells[row, collist[0]].Value != null) {
                                    originIdList.Add((double)worksheet.Cells[row, collist[0]].Value);
                                    //originIdList.Add(worksheet.Cells[row, collist[0]].Value.ToString());

                                }
                            }
                            //translation excel maybe 对不上 origin excel's col 编号  so origin and translation excel should
                            //use own row/col data
                            transRowEnd = workshee2.Dimension.End.Row;
                            transColStart = workshee2.Dimension.Start.Column;
                            transColEnd = workshee2.Dimension.End.Column;
                            for (int row = workRow; row <= transRowEnd; row++)
                            {
                                if (workshee2.Cells[row, transColStart].Value != null)
                                {
                                    translateIdList.Add((double)workshee2.Cells[row, transColStart].Value);
                                    //translateIdList.Add(workshee2.Cells[row, collist[0]].Value.ToString());
                                }
                            }
                            //比对origin 有 而translate无的
                            for (int index = 0; index < originIdList.Count; index++)
                            {
                                if (!translateIdList.Contains(originIdList[index]))
                                {
                                    translate2AddIdList.Add(originIdList[index]);
                                }
                            }

                            //比对origin 无 而translate有的
                            for (int index = 0; index < translateIdList.Count; index++)
                            {
                                if (!originIdList.Contains(translateIdList[index]))
                                {
                                    translate2DelIdList.Add(translateIdList[index]);
                                }
                            }

                            //对translate 表进行删改

                            //string transId = "";
                            double transId;
                            //删除
                            while (translate2DelIdList.Count > 0)
                            {
                                transColStart = workshee2.Dimension.Start.Column;
                                transColEnd = workshee2.Dimension.End.Column;
                                transRowStart = workRow;
                                transRowEnd = workshee2.Dimension.End.Row;
                                for (int row = transRowStart; row <= transRowEnd; row++)
                                {
                                    if (workshee2.Cells[row, transColStart].Value == null)
                                        continue;
                                    //transId = workshee2.Cells[row, transColStart].Value.ToString();
                                    transId =(double)workshee2.Cells[row, transColStart].Value;

                                    if (translate2DelIdList.Contains(transId))
                                    {
                                        //workshee2.Cells[row, transColStart, row, transColEnd].Value = null;
                                        //ExcelRange range = workshee2.Cells[row, transColStart, row, transColEnd];
                                        //range.Clear();
                                        workshee2.DeleteRow(row);
                                        translate2DelIdList.Remove(transId);
                                    }
                                }

                            }
                            //添加
                            int curTranslateCol = 1;
                            //string originId = "";
                            double originId;
                            while (translate2AddIdList.Count > 0)
                            {
                                transColStart = workshee2.Dimension.Start.Column;
                                transColEnd = workshee2.Dimension.End.Column;
                                transRowStart = workRow;
                                transRowEnd = workshee2.Dimension.End.Row;

                                for (int row = workRow; row <= rowEnd; row++)
                                {//原表中的新数据所在行
                                    if (worksheet.Cells[row, collist[0]].Value == null)
                                        continue;
                                    //originId = worksheet.Cells[row, collist[0]].Value.ToString();
                                    originId = (double)worksheet.Cells[row, collist[0]].Value;
                                    if (translate2AddIdList.Contains(originId))
                                    {//找到原表中的新增数据
                                        workshee2.InsertRow(row, 1);

                                        for (int colindex = 0; colindex < collist.Count; colindex++)
                                        {
                                            //将原表里的内容加入
                                            workshee2.Cells[row, curTranslateCol].Value = worksheet.Cells[row, collist[colindex]].Value;
                                            if (colindex == 0)
                                            {
                                                curTranslateCol++;
                                            }
                                            else
                                            {
                                                curTranslateCol += transTypeList.Length + 1;
                                            }
                                        }
                                        translate2AddIdList.Remove(originId);
                                    }
                                }
                            }

                            package2.Save();
                        }
                    }
                    else
                    {
                        newFile = new FileInfo(pathTranslate);
                        using (ExcelPackage package2 = new ExcelPackage(newFile))//每一个translate excel
                        {
                            ExcelWorksheet worksheet2 = package2.Workbook.Worksheets.Add("Sheet1");
                            int curTranslateCol = 1;
                            for (int colindex = 0; colindex < collist.Count; colindex++)
                            {
                                for (int rowindex = rowStart; rowindex <= rowEnd; rowindex++)
                                {
                                    //将原表里的内容加入
                                    worksheet2.Cells[rowindex, curTranslateCol].Value = worksheet.Cells[rowindex, collist[colindex]].Value;
                                    //根据翻译类型添加翻译字段
                                    //id不需要翻译
                                    if (colindex != 0)
                                    {
                                        for (int k = 0; k < transTypeList.Length; k++)
                                        {
                                            if (rowindex < workRow)
                                            {
                                                worksheet2.Cells[rowindex, curTranslateCol + 1 + k].Value = worksheet.Cells[rowindex, collist[colindex]].Value + "_" + transTypeList[k];
                                            }
                                            else
                                            {
                                                //worksheet2.Cells[rowindex, curTranslateCol + 1 + k].Value = worksheet.Cells[rowindex, collist[colindex]].Value;
                                            }
                                        }
                                    }
                                }
                                //多种语言时
                                if (colindex == 0)
                                {
                                    curTranslateCol++;
                                }
                                else
                                {
                                    curTranslateCol += transTypeList.Length + 1;
                                }
                            }
                            package2.Save();
                        };
                    }

                };
            }
            //Console.Read();
            return true;
        }

        static void TraverseFile(string rootPath)
        {
            DirectoryInfo folder = new DirectoryInfo(rootPath);

            DirectoryInfo[] directoryinfos = folder.GetDirectories();
            FileInfo[] fileinfos = folder.GetFiles();
            if (directoryinfos.Length > 0)
            {
                //目录遍历
                foreach (DirectoryInfo dir in directoryinfos)
                {
                    if (dir.Extension == ".meta")
                        continue;
                    //过滤
                    if (filterFolderList.Count > 0)
                    {
                        bool isFilter = false;
                        for (int i = 0; i < filterFolderList.Count; i++)
                        {
                            if (dir.FullName == filterFolderList[i])
                            {
                                isFilter = true;
                                break;
                            }
                        }
                        if (isFilter)
                        {
                            continue;
                        }
                    }

                    TraverseFile(rootPath + "/" + dir.Name);
                }
            }
            if (fileinfos.Length > 0)
            {
                //文件遍历
                foreach (FileInfo dir in fileinfos)
                {
                    //文件查找匹配
                    if (dir.Extension != ".xlsx")
                        continue;
                    if (dir.Name.Contains("translate"))
                    {
                        continue;
                    }

                    //pathLists.Add(dir.FullName + " " + dir.Name + " " + dir.DirectoryName);
                    pathLists.Add(dir.FullName);
                }
            }
        }

        static void TraverseTranslateFile(string rootPath) {
            DirectoryInfo folder = new DirectoryInfo(rootPath);

            DirectoryInfo[] directoryinfos = folder.GetDirectories();
            FileInfo[] fileinfos = folder.GetFiles();
            if (directoryinfos.Length > 0)
            {
                //目录遍历
                foreach (DirectoryInfo dir in directoryinfos)
                {
                    if (dir.Extension == ".meta")
                        continue;
                    TraverseTranslateFile(rootPath + "/" + dir.Name);
                }
            }
            if (fileinfos.Length > 0)
            {
                //文件遍历
                foreach (FileInfo dir in fileinfos)
                {
                    string pattern = @"[\s\S]+_translate[\s\S]";

                    //文件查找匹配
                    if (dir.Extension != ".xlsx")
                        continue;
                    //if (dir.Name.Contains("translate"))
                    //{
                    //    continue;
                    //}
                    MatchCollection mc = Regex.Matches(dir.Name, pattern, RegexOptions.IgnoreCase);
                    if (mc.Count > 0)
                    {
                        pathTransLists.Add(dir.FullName);
                        
                    }
                }
            }
        }


        /// <summary>
        /// 删除翻译表
        /// </summary>
        /// <param name="rootPath"></param>
        public static bool DelTranslate(string rootPath) {
            if (!Directory.Exists(rootPath)) {
                Form1.ShowTips("不存在"+rootPath+"此目录");
                return false;
            }
            pathTransLists.Clear();
            TraverseTranslateFile(rootPath);
            foreach (var i in pathTransLists)
            {
                //Form1.ShowTips(i);
                if (File.Exists(i)) {
                    File.Delete(i);
                }
            }
            return true;
        }
    }
}

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
        //origin xlsx path list
        static List<string> pathLists = new List<string>();
        static List<string> pathTransLists = new List<string>();


        static List<string> filterFolderList = new List<string>();//过滤不需要导出的文件夹

        static List<double> originIdList = new List<double>();
        static List<double> translateIdList = new List<double>();
        static List<double> translate2AddIdList = new List<double>();
        static List<double> translate2DelIdList = new List<double>();

        //del col data
        static List<string> originColNameList = new List<string>();
        static List<string> translateColNameList = new List<string>();
        static Dictionary<string, List<string>> translateColNamePostfixList = new Dictionary<string, List<string>>();
        static List<string> translate2AddColNameList = new List<string>();
        static List<string> translate2DelColNameList = new List<string>();

        static List<string> prevTransTypeList = new List<string>();//修改前的翻译类型，之前导出时使用的
        static List<string> nowTransTypeList = new List<string>();//修改后的翻译类型列表，目前最新的

        //add/del translateType data
        static List<string> translate2AddTransNameList = new List<string>();
        static List<string> translate2DelTransNameList = new List<string>();

        static bool canExportLua = false;
        //所有翻译原表的表名及其需要翻译的字段名列表（一张表可能有多个sheet，每个sheet在此视为一张单独的表）
        static Dictionary<string, List<string>> tbAndFieldName = new Dictionary<string, List<string>>();
        const string ORIGINIDENTITY = "_toTrans";
        //static List<string> originIdList = new List<string>();
        //static List<string> translateIdList = new List<string>();
        //static List<string> translate2AddIdList = new List<string>();
        //static List<string> translate2DelIdList = new List<string>();

        public static bool Export(string rootPath, string luaExportPath = null, string[] transTypeList = null, string filterPath = null)
        {
            try
            {
                if (!Directory.Exists(rootPath))
                {
                    Form1.ShowTips("不存在" + rootPath + "此目录");
                    return false;
                }
                transTypeList = transTypeList == null ? new string[] { "ch" } : transTypeList;
                nowTransTypeList = transTypeList.ToList<string>();

                rootPath = rootPath == "" ? @"E:\zzhx\trunk\data\ai" : rootPath;
                filterFolderList.Clear();
                if (filterPath != null && filterPath != "")
                {
                    filterFolderList.Add(filterPath);
                }
                if (luaExportPath != null && luaExportPath != "")
                {
                    //导出lua
                    canExportLua = true;
                    luaExportPath += @"\lua_tb_field.lua";
                }
                else
                {
                    canExportLua = false;
                }

                pathLists.Clear();
                tbAndFieldName.Clear();

                prevTransTypeList.Clear();

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
                        ExcelWorksheet worksheet = null;
                        for (int sheetIndex = 1; sheetIndex <= package.Workbook.Worksheets.Count; sheetIndex++)
                        {
                            worksheet = package.Workbook.Worksheets[sheetIndex];
                            collist.Clear();
                            //record the col name which to be exported;
                            originColNameList.Clear();
                            //Form1.ShowTips("name"+ worksheet.Name);
                            if (!worksheet.Name.Contains("tb_table"))
                            {
                                continue;
                            }
                            //空表不处理。
                            if (worksheet.Dimension == null)
                            {
                                continue;
                            }
                            string sheetName = worksheet.Name;
                            //Form1.ShowProgress("从"+excelFile.Name+"原表导出sheet:"+sheetName);
                            //从有内容的行列开始
                            int colStart = worksheet.Dimension.Start.Column;  //工作区开始列
                            int colEnd = worksheet.Dimension.End.Column;       //工作区结束列
                            int rowStart = worksheet.Dimension.Start.Row;       //工作区开始行号
                            int rowEnd = worksheet.Dimension.End.Row;       //工作区结束行号

                            int workRow = 2;//默认工作行
                            //bool isRightCol = false;//是否是正规的需要导出的翻译列
                            List<int> theRightColList = new List<int>();//正规的可能需要导出的列的编号列表
                            //记录每张表的翻译字段
                            for (int row = rowStart; row <= 5; row++)
                            {//给5因为有些表的字段名下面还加了些描述等中文字符，故多遍历几次，以保证进入正文(不给rowend，因为会做太多无谓的遍历，损坏性能）
                                for (int col = colStart; col <= colEnd; col++)
                                {
                                    string text = worksheet.Cells[row, col].Text;
                                    if (text == "id" || text == "ID" || text.ToLower() == "id")
                                    {
                                        if (!collist.Contains(col))
                                        {
                                            collist.Add(col);
                                        }
                                        break;
                                    }
                                    if (text.Trim().ToLower() == "str") {
                                        theRightColList.Add(col);
                                    }
                                    //记录正文开始行
                                    if (text.Contains("删"))
                                    {
                                        //删所在行的下一行为正文开始行
                                        workRow = row + 1;
                                        break;
                                    }
                                    MatchCollection mc = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
                                    //找到需要翻译的列
                                    if (theRightColList.Contains(col) &&mc.Count > 0)// && workRow > 2防止乱填无字段内容
                                    {
                                        if (!collist.Contains(col))
                                        {
                                            collist.Add(col);
                                        }
                                        //添加翻译表的翻译字段用以导出到lua表(表名+翻译字段名);
                                        //获取表名

                                        //---old:一张表一个sheet
                                        //int posSlash = pathLists[i].LastIndexOf(@"\");
                                        //string tbname = pathLists[i].Substring(posSlash + 1, pathLists[i].Length - posSlash - 1);
                                        //int posDot = tbname.LastIndexOf(".");
                                        //tbname = tbname.Substring(0, posDot);
                                        //---new:一张表多个sheet
                                        string tbname = sheetName;
                                        if (tbAndFieldName.ContainsKey(tbname))
                                        {
                                            //Form1.ShowTips(worksheet.Cells[rowStart, col].Text);
                                            if (!tbAndFieldName[tbname].Contains(worksheet.Cells[rowStart, col].Text))//worksheet.Cells[rowStart, collist[col]].Value.ToString()
                                            {
                                                tbAndFieldName[tbname].Add(worksheet.Cells[rowStart, col].Text);
                                            }
                                        }
                                        else
                                        {
                                            tbAndFieldName.Add(tbname, new List<string>());
                                            tbAndFieldName[tbname].Add(worksheet.Cells[rowStart, col].Text);
                                        }

                                    }
                                }
                            }
                            //只有id列，不用导出
                            if (collist.Count <= 1)
                            {
                                continue;
                            }

                            //记录下原表里所有需要翻译的字段的名字
                            foreach (int col in collist)
                            {
                                string text = worksheet.Cells[rowStart, col].Text;
                                if (!originColNameList.Contains(text))
                                {
                                    originColNameList.Add(text);
                                }
                            }

                            //export

                            //-------old
                            //string pathTranslate = pathLists[i];
                            //pathTranslate = pathTranslate.Substring(0, pathTranslate.Length - 5);//.xlsx
                            //pathTranslate += "_translate.xlsx";
                            //string dir = pathTranslate.Substring(0, pathTranslate.LastIndexOf("\\"));
                            //string transfilename = pathTranslate.Substring(pathTranslate.LastIndexOf("\\"));//.xlsx
                            //transfilename = transfilename.Substring(0, transfilename.LastIndexOf("."));
                            //if (!Directory.Exists(dir + "\\translate"))
                            //    Directory.CreateDirectory(dir + "\\translate");
                            //pathTranslate = dir + "\\translate" + transfilename + ".xlsx";

                            //-------new
                            string pathTranslate = pathLists[i];
                            pathTranslate = pathTranslate.Substring(0, pathTranslate.LastIndexOf("\\"));//.xlsx
                            pathTranslate += "\\" + sheetName + "_translate.xlsx";
                            string dir = pathTranslate.Substring(0, pathTranslate.LastIndexOf("\\"));
                            string transfilename = pathTranslate.Substring(pathTranslate.LastIndexOf("\\"));//.xlsx
                            transfilename = transfilename.Substring(0, transfilename.LastIndexOf("."));
                            if (!Directory.Exists(dir + "\\translate"))
                                Directory.CreateDirectory(dir + "\\translate");
                            pathTranslate = dir + "\\translate" + transfilename + ".xlsx";

                            FileInfo newFile = new FileInfo(pathTranslate);
                            if (newFile.Exists)
                            {
                                //newFile.Delete();
                                using (ExcelPackage package2 = new ExcelPackage(newFile))//每一个translate excel
                                {
                                    //**********:the order of  add/del of the origin excel：firstly add/del row(record),then modify transType,finally add/del col(field)

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
                                    Console.WriteLine(worksheet.Name);
                                    for (int row = workRow; row <= rowEnd; row++)
                                    {
                                        if (worksheet.Cells[row, collist[0]].Value != null)
                                        {
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

                                    //----------------------对translate 表row进行删改

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
                                            transId = (double)workshee2.Cells[row, transColStart].Value;

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
                                    //row添加
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

                                    //-----------------------------------------========================================华丽的 split line, the second step
                                    //翻译类型的增删,这一阶段要抹掉ORIGINIDENTITY，防止后面的field处理会全部重新生成field，且要将与现在的翻译类型不匹配的所有field的增改做完
                                    //后面的阶段直接用最新的翻译类型操作数据
                                    translate2AddTransNameList.Clear();
                                    translate2DelTransNameList.Clear();
                                    transRowStart = workshee2.Dimension.Start.Row;
                                    //init the prevTransTypeList and be sure to init it just one time
                                    if (prevTransTypeList.Count <= 0)
                                    {
                                        string name = "";
                                        bool canAdd = false;
                                        for (int col = transColStart; col <= transColEnd; col++)
                                        {
                                            if (workshee2.Cells[transRowStart, col].Value != null)
                                            {
                                                name = (string)workshee2.Cells[transRowStart, col].Value;

                                                if (name.EndsWith(ORIGINIDENTITY))
                                                { //the origin excel field
                                                    canAdd = !canAdd;
                                                    if (canAdd)
                                                    { //进入添加时机
                                                        continue;
                                                    }
                                                }
                                                if (canAdd)
                                                {
                                                    name = (string)workshee2.Cells[transRowStart, col].Value;
                                                    prevTransTypeList.Add(name.Substring(name.LastIndexOf("_") + 1));
                                                }
                                            }
                                        }
                                    }
                                    //previous transType have but now don't
                                    for (int index = 0; index < prevTransTypeList.Count; index++)
                                    {
                                        if (!nowTransTypeList.Contains(prevTransTypeList[index]))
                                        {
                                            translate2DelTransNameList.Add(prevTransTypeList[index]);
                                        }
                                    }

                                    //previous transType don't have but now have
                                    for (int index = 0; index < nowTransTypeList.Count; index++)
                                    {
                                        if (!prevTransTypeList.Contains(nowTransTypeList[index]))
                                        {
                                            translate2AddTransNameList.Add(nowTransTypeList[index]);
                                        }
                                    }

                                    string colPostfix = "";
                                    string colname = "";
                                    // delete the translatetype
                                    if (translate2DelTransNameList.Count > 0)
                                    {
                                        for (int col = transColStart; col <= transColEnd; )
                                        {
                                            transColStart = workshee2.Dimension.Start.Column;
                                            transColEnd = workshee2.Dimension.End.Column;
                                            transRowStart = workshee2.Dimension.Start.Row;
                                            transRowEnd = workshee2.Dimension.End.Row;
                                            if (workshee2.Cells[transRowStart, col].Value == null)
                                                continue;
                                            colPostfix = workshee2.Cells[transRowStart, col].Value as string;
                                            colPostfix = colPostfix.Substring(colPostfix.LastIndexOf("_") + 1);
                                            if (translate2DelTransNameList.Contains(colPostfix))
                                            {
                                                //del the cols that contains the col end with translate type
                                                workshee2.DeleteColumn(col);
                                                //the columns after the deleted col will be shifted forward,so col can't ++
                                            }
                                            else
                                            {
                                                col++;
                                            }
                                        }
                                    }

                                    bool insertOver = false;
                                    transColStart = workshee2.Dimension.Start.Column;

                                    while (translate2AddTransNameList.Count > 0 && !insertOver)
                                    {
                                        transColEnd = workshee2.Dimension.End.Column;
                                        transRowStart = workshee2.Dimension.Start.Row;
                                        transRowEnd = workshee2.Dimension.End.Row;
                                        // add the translatetype
                                        for (int col = transColStart; col <= transColEnd; col++)
                                        {
                                            if (workshee2.Cells[transRowStart, col].Value == null)
                                                continue;
                                            colname = workshee2.Cells[transRowStart, col].Value as string;
                                            int iindx = colname.LastIndexOf("_");
                                            if (iindx < 0)
                                            {
                                                continue;
                                            }
                                            colPostfix = colname.Substring(iindx);
                                            if (colPostfix == ORIGINIDENTITY)
                                            {
                                                colname = colname.Substring(0, colname.LastIndexOf("_"));
                                                //insert the new translation field after _toTrans field
                                                for (int s = 0; s < translate2AddTransNameList.Count; s++)
                                                {
                                                    workshee2.InsertColumn(col + 1, 1);

                                                    workshee2.Cells[transRowStart, col + 1].Value = colname + "_" + translate2AddTransNameList[s];
                                                    workshee2.Cells[transRowStart + 1, col + 1].Value = colname + "_" + translate2AddTransNameList[s];
                                                }

                                                //判断此列之后没有要插入新翻译列的列了
                                                if (transColEnd - col + translate2AddTransNameList.Count <= nowTransTypeList.Count)
                                                {
                                                    insertOver = true;
                                                }
                                                transColStart = col + 1;//插入数据后，表格结构已经发生了变换
                                                break;
                                            }
                                            if (col == transColEnd)
                                                insertOver = true;
                                        }
                                    }

                                    transColStart = workshee2.Dimension.Start.Column;
                                    transColEnd = workshee2.Dimension.End.Column;
                                    transRowStart = workshee2.Dimension.Start.Row;
                                    transRowEnd = workshee2.Dimension.End.Row;

                                    //record the col num which is the origin col to add back _toTrans
                                    //List<int> originTransColNum = new List<int>();
                                    //将_toTrans 去掉先
                                    for (int col = transColStart; col <= transColEnd; col++)
                                    {
                                        if (workshee2.Cells[transRowStart, col].Value == null)
                                            continue;
                                        colname = workshee2.Cells[transRowStart, col].Value as string;
                                        int iindex = colname.LastIndexOf("_");
                                        if (iindex < 0)
                                        {
                                            continue;
                                        }
                                        colPostfix = colname.Substring(iindex);
                                        if (colPostfix == ORIGINIDENTITY)
                                        {
                                            colname = colname.Substring(0, colname.LastIndexOf("_"));
                                            workshee2.Cells[transRowStart, col].Value = colname;

                                            //originTransColNum.Add(col);
                                        }
                                    }

                                    //========================================华丽的 split line, the third step
                                    //比对改动后的原表与translate表的字段名的不同后添加删除,by this time,the translateTypes maybe have been modify(add or delete)
                                    //so we shouldn't use the params transTypeList as judgment data

                                    translate2AddColNameList.Clear();
                                    translate2DelColNameList.Clear();
                                    translateColNameList.Clear();
                                    translateColNamePostfixList.Clear();

                                    transRowStart = workshee2.Dimension.Start.Row;


                                    //记录translate表里的不跟翻译后缀的原始字段名以及记录下对应的所有带后缀的字段名
                                    for (int col = transColStart; col <= transColEnd; col++)
                                    {
                                        if (workshee2.Cells[transRowStart, col].Value != null)
                                        {
                                            string name = (string)workshee2.Cells[transRowStart, col].Value;
                                            bool canAdd = true;
                                            string postfix = GetPostfixName(name);
                                            //将翻译类型的字段过滤
                                            //之前增删翻译类型阶段可能已经加了新的字段 故还要与nowTransTypelist比对
                                            if (postfix != null)
                                            {
                                                if (prevTransTypeList.Contains(postfix) || nowTransTypeList.Contains(postfix))
                                                    canAdd = false;
                                            }
                                            if (canAdd)
                                            {
                                                translateColNameList.Add(name);
                                            }
                                        }
                                    }
                                    //比对origin 有 而translate无的
                                    for (int index = 0; index < originColNameList.Count; index++)
                                    {
                                        if (!translateColNameList.Contains(originColNameList[index]))
                                        {
                                            translate2AddColNameList.Add(originColNameList[index]);
                                        }
                                    }

                                    //比对origin 无 而translate有的
                                    for (int index = 0; index < translateColNameList.Count; index++)
                                    {
                                        if (!originColNameList.Contains(translateColNameList[index]))
                                        {
                                            translate2DelColNameList.Add(translateColNameList[index]);
                                        }
                                    }

                                    //增删列之后  翻译表剩下的列名
                                    //foreach (string name in translate2DelColNameList) {
                                    //    if (translateColNameList.Contains(name)) {
                                    //        translateColNameList.Remove(name);
                                    //    }
                                    //}
                                    //translateColNameList.AddRange(translate2AddColNameList);
                                    //foreach (string name in translateColNameList) {
                                    //    Form1.ShowTips("translateColNameList:" + name);
                                    //}
                                    //----------------------对translate 表col进行删改
                                    //col 删
                                    //string colname = "";
                                    int toDelColNum = nowTransTypeList.Count;//the num of the col name which take translate postfix


                                    while (translate2DelColNameList.Count > 0)
                                    {
                                        transColStart = workshee2.Dimension.Start.Column;
                                        transColEnd = workshee2.Dimension.End.Column;
                                        transRowStart = workshee2.Dimension.Start.Row;
                                        transRowEnd = workshee2.Dimension.End.Row;
                                        for (int col = transColStart; col <= transColEnd; col++)
                                        {
                                            if (workshee2.Cells[transRowStart, col].Value == null)
                                                continue;
                                            colname = workshee2.Cells[transRowStart, col].Value as string;
                                            if (translate2DelColNameList.Contains(colname))
                                            {
                                                //del the cols that contains the col end with translate type
                                                for (int s = 0; s <= toDelColNum; s++)
                                                {//include all of the postfix col and the origin col
                                                    workshee2.DeleteColumn(col);
                                                }
                                                translate2DelColNameList.Remove(colname);
                                            }
                                        }
                                    }
                                    //used to judge the column be added or not
                                    while (translate2AddColNameList.Count > 0)
                                    {
                                        transColStart = workshee2.Dimension.Start.Column;
                                        transColEnd = workshee2.Dimension.End.Column;
                                        transRowStart = workshee2.Dimension.Start.Row;
                                        transRowEnd = workshee2.Dimension.End.Row;

                                        string newColName = "";
                                        int curTranslateRow = 1;
                                        int insertCol = 0;
                                        for (int col = colStart; col <= colEnd; col++)//colEnd
                                        {
                                            if (worksheet.Cells[rowStart, col].Value == null)
                                                continue;
                                            colname = worksheet.Cells[rowStart, col].Value as string;
                                            if (translate2AddColNameList.Contains(colname))
                                            {
                                                //find the new content from the origin excel
                                                //the col with the translate postfix just need to set the col name and the chinese description
                                                //每次都是往id列后面插入新列
                                                insertCol = 2;// transColEnd + 1;
                                                for (int s = transTypeList.Length - 1; s >= 0; s--)
                                                {
                                                    workshee2.InsertColumn(insertCol, 1);
                                                    newColName = colname + "_" + transTypeList[s];

                                                    workshee2.Cells[curTranslateRow, insertCol].Value = newColName;
                                                    workshee2.Cells[curTranslateRow + 1, insertCol].Value = worksheet.Cells[rowStart + 1, col].Value;
                                                }
                                                //the origin content
                                                workshee2.InsertColumn(insertCol, 1);
                                                workshee2.Cells[curTranslateRow, insertCol].Value = colname + ORIGINIDENTITY;
                                                workshee2.Cells[curTranslateRow + 1, insertCol].Value = worksheet.Cells[rowStart + 1, col].Value;
                                                curTranslateRow = 3;//standard format:the third row is the start row of work content
                                                for (int s = curTranslateRow; s <= transRowEnd; s++)
                                                {
                                                    workshee2.Cells[s, insertCol].Value = worksheet.Cells[rowStart + s - 1, col].Value;
                                                }
                                                translate2AddColNameList.Remove(colname);
                                                break;
                                            }
                                        }
                                    }
                                    //add back _toTrans
                                    transColStart = workshee2.Dimension.Start.Column;
                                    transColEnd = workshee2.Dimension.End.Column;
                                    transRowStart = workshee2.Dimension.Start.Row;
                                    transRowEnd = workshee2.Dimension.End.Row;
                                    for (int col = transColStart; col <= transColEnd; col++)
                                    {
                                        if (workshee2.Cells[transRowStart, col].Value == null)
                                            continue;
                                        colname = workshee2.Cells[rowStart, col].Value as string;
                                        string postfix = GetPostfixName(colname);
                                        if (colname.ToLower() != "id" && !nowTransTypeList.Contains(postfix))
                                        {
                                            string str = workshee2.Cells[rowStart, col].Value as string;

                                            if (str.EndsWith(ORIGINIDENTITY))
                                            {
                                                continue;
                                            }
                                            workshee2.Cells[rowStart, col].Value = colname + ORIGINIDENTITY;
                                        }
                                    }


                                    //存在不增删原表的行列，只修改行列里面的原数据的情况，故每次导出，直接用原表当前的内容覆盖翻译表的
                                    //对应翻译字段那列的内容（即标记了_toTrans)的字段
                                    //翻译表需要翻译的原字段名：列号映射
                                    Dictionary<string, int> name2Col = new Dictionary<string, int>();
                                    for (int col = transColStart; col <= transColEnd; col++)
                                    {
                                        if (workshee2.Cells[transRowStart, col].Value == null)
                                            continue;
                                        colname = workshee2.Cells[rowStart, col].Value as string;
                                        string postfix = GetPostfixName(colname);
                                        if (colname.ToLower() != "id" && !nowTransTypeList.Contains(postfix))
                                        {
                                            string str = workshee2.Cells[rowStart, col].Value as string;

                                            if (str.EndsWith(ORIGINIDENTITY))
                                            {
                                                name2Col.Add(GetPrefixName(str), col);
                                                continue;
                                            }
                                        }
                                    }

                                    //foreach(KeyValuePair<string,int> t in name2Col){
                                    //    Form1.ShowTips(t.Key.ToString() + " " + t.Value.ToString());
                                    //}

                                    foreach (int index in collist)
                                    {
                                        //从字段名的下一行开始，用原表内容覆盖翻译表内容
                                        string name = worksheet.Cells[1, index].Value as string;
                                        if (name2Col.ContainsKey(name))
                                        {
                                            int transCol = name2Col[name];
                                            for (int row = 2; row <= rowEnd; row++)
                                            {
                                                workshee2.Cells[row, transCol].Value = worksheet.Cells[row, index].Value;
                                            }
                                        }
                                    }
                                    package2.Save();
                                    Console.WriteLine(pathTranslate);
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
                                            object str = worksheet.Cells[rowindex, collist[colindex]].Value;
                                            //有些字段可能有附加描述列，不填则为空
                                            if (str != null && rowindex == rowStart && str.ToString().ToLower() != "id")
                                            { //field name need to add "_toTrans"
                                                str = str + ORIGINIDENTITY;
                                            }
                                            worksheet2.Cells[rowindex, curTranslateCol].Value = str;


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
                                    Console.WriteLine(pathTranslate);
                                };
                            }
                        }
                    };
                }

            }
            catch (Exception e)
            {
                Form1.ShowTips("导表异常:"+e.Message+" "+e.Source+" "+e.StackTrace);
            }
            //Console.Read();
            if (canExportLua)
            {
                try
                {
                    if (File.Exists(luaExportPath))
                        File.Delete(luaExportPath);

                    FileStream fs = new FileStream(luaExportPath, FileMode.Create);
                    StreamWriter sw = new StreamWriter(fs);
                    string preSpace = "\t";
                    sw.WriteLine("lua_tb_field={");
                    foreach (KeyValuePair<string, List<string>> keyValue in tbAndFieldName)
                    {
                        List<string> valu = keyValue.Value;
                        if (valu.Count <= 0)
                        {
                            continue;
                        }
                        sw.WriteLine(preSpace + keyValue.Key + "={");
                        int count = valu.Count;
                        for (int i = 0; i < valu.Count; i++)
                        {
                            if (valu[i].Trim() == "")
                            {
                                continue;
                            }
                            sw.WriteLine(preSpace + string.Format("[{0}]=\"{1}\",", (i + 1).ToString(), valu[i].Trim()));
                        }
                        sw.WriteLine(preSpace + "},");
                    }
                    sw.WriteLine("}");
                    sw.WriteLine("return lua_tb_field;");
                    sw.Flush();
                    sw.Close();
                    fs.Close();
                }
                catch (Exception e)
                {
                    Form1.ShowTips("导lua字段表异常:" + e.Message+" "+e.Source+" "+e.StackTrace);
                }
            }
            return true;
        }

        /// <summary>
        /// 获取所有原表的路径
        /// </summary>
        /// <param name="rootPath"></param>
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
                    if (dir.Extension != ".xlsx"&& dir.Extension != ".xlsm")
                        continue;
                    if (dir.Name.Contains("translate"))
                    {
                        continue;
                    }
                    pathLists.Add(dir.FullName);
                    //E: \UnityProject\test\excel\tb_table_copy.xlsx
                    //Console.WriteLine(dir.FullName);
                    //Console.WriteLine(dir.Name);
                    //tb_table_copy.xlsx
                    //一张表里可能的多个sheet也作为一张单独的表加入
                    //FileInfo excelFile = new FileInfo(dir.FullName);
                    //if (excelFile.Exists)
                    //{
                    //    using (ExcelPackage package = new ExcelPackage(excelFile))//每一个原excel
                    //    {
                    //        ExcelWorksheet worksheet = null;
                    //        //添加路径全名
                    //        for (int sheetIndex = 1; sheetIndex <= package.Workbook.Worksheets.Count; sheetIndex++)
                    //        {
                    //            worksheet = package.Workbook.Worksheets[sheetIndex];
                    //            if (!worksheet.Name.Contains("tb_table"))
                    //            {
                    //                continue;
                    //            }
                    //            pathLists.Add(dir.FullName.Substring(0, dir.FullName.LastIndexOf(@"\")) + @"\" + worksheet.Name);
                    //            string sheetName = worksheet.Name;
                    //            if (!tbAndFieldName.ContainsKey(sheetName))
                    //            {
                    //                tbAndFieldName.Add(sheetName, new List<string>());
                    //            }
                    //        }
                    //    }
                    //}
                    int pos = dir.Name.LastIndexOf(".");
                    string tbname = dir.Name.Substring(0, pos);
                    if (!tbAndFieldName.ContainsKey(tbname))
                    {
                        tbAndFieldName.Add(tbname, new List<string>());
                    }
                }
            }
        }

        static void TraverseTranslateFile(string rootPath)
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
                    if (dir.Extension != ".xlsx"&&dir.Extension!=".xlsm")
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

        public static string GetPostfixName(string str)
        {
            int index = str.LastIndexOf("_");
            if (index < 0)
            {
                return null;
            }
            else
            {
                return str.Substring(index + 1);
            }

        }

        public static string GetPrefixName(string str)
        {
            int index = str.LastIndexOf("_");
            if (index < 0)
            {
                return null;
            }
            else
            {
                return str.Substring(0,index);
            }

        }

        /// <summary>
        /// 删除翻译表
        /// </summary>
        /// <param name="rootPath"></param>
        public static bool DelTranslate(string rootPath)
        {
            if (!Directory.Exists(rootPath))
            {
                Form1.ShowTips("不存在" + rootPath + "此目录");
                return false;
            }
            pathTransLists.Clear();
            TraverseTranslateFile(rootPath);
            foreach (var i in pathTransLists)
            {
                //Form1.ShowTips(i);
                if (File.Exists(i))
                {
                    File.Delete(i);
                }
            }
            return true;
        }
    }
}

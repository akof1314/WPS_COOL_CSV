﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using AddInDesignerObjects;
using Excel;
using Office;
using wps_cool_csv.Properties;

namespace wps_cool_csv
{
    public class CoolCsv : IDTExtensibility2, IRibbonExtensibility
    {
        private Application app;
        private readonly Dictionary<string, Encoding> fileDict = new Dictionary<string, Encoding>();
        private bool flagSaveAs;

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            app = Application as Application;
            app.WorkbookOpen += AppOnWorkbookOpen;
            app.WorkbookBeforeSave += AppOnWorkbookBeforeSave;
            app.WorkbookBeforeClose += AppOnWorkbookBeforeClose;
            app.SheetSelectionChange += AppOnSheetSelectionChange;
            app.SheetActivate += AppOnSheetActivate;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            if (app != null)
            {
                app.WorkbookOpen -= AppOnWorkbookOpen;
                app.WorkbookBeforeSave -= AppOnWorkbookBeforeSave;
                app.WorkbookBeforeClose -= AppOnWorkbookBeforeClose;
                app.SheetSelectionChange -= AppOnSheetSelectionChange;
                app.SheetActivate -= AppOnSheetActivate;
            }
        }

        private void AppOnWorkbookOpen(Workbook wb)
        {
            ResetFormatConditionsHighlight(wb.ActiveSheet);
            ResetFreezeHeader(wb.ActiveSheet);
        }

        private void AppOnWorkbookBeforeSave(Workbook wb, bool saveasui, ref bool cancel)
        {
            // 手动调用的另存为就不再执行
            if (flagSaveAs)
            {
                flagSaveAs = false;
                return;
            }

            // 不需要保存处理的话，直接返回
            if (!SettingsCsv.Default.EnableSaveEncode)
            {
                return;
            }

            string fileName = wb.FullName;
            if (!fileName.ToLower().EndsWith(".csv"))
            {
                return;
            }

            if (!fileDict.ContainsKey(fileName))
            {
                Encoding encoding = GetFileEncoding(fileName);
                fileDict[fileName] = encoding;
            }

            Encoding fileEncoding = fileDict[fileName];
            if (Equals(fileEncoding, Encoding.UTF8) || Equals(fileEncoding, Encoding.Unicode) || Equals(fileEncoding, Encoding.BigEndianUnicode))
            {
                // 自己来操作
                cancel = true;

                // 如果是另存为，那么需要获取另存为的文件名，所以要自己显示另存为保存框
                if (saveasui)
                {
                    FileDialog fileDialog = app.FileDialog[MsoFileDialogType.msoFileDialogSaveAs];
                    fileDialog.InitialFileName = wb.Name;
                    fileDialog.AllowMultiSelect = false;
                    fileDialog.Title = "另存为";
                    FileDialogFilters fileDialogFilters = fileDialog.Filters;
                    bool flag = false;
                    for (int i = 1; i <= fileDialogFilters.Count; i++)
                    {
                        if ("*.csv".Equals(fileDialogFilters.Item(i).Extensions))
                        {
                            flag = true;
                            fileDialog.FilterIndex = i;
                            break;
                        }
                    }
                    if (!flag)
                    {
                        fileDialogFilters.Add("CSV (逗号分隔)", "*.csv");
                        fileDialog.FilterIndex = fileDialogFilters.Count;
                    }

                    // 取消了操作
                    if (fileDialog.Show() == 0)
                    {
                        return;
                    }

                    string fileNewName = fileDialog.SelectedItems.Item(1);
                    if (!fileNewName.ToLower().EndsWith(".csv"))
                    {
                        // 非csv则普通保存，需要标志一下，否则会再进来
                        flagSaveAs = true;
                        wb.SaveAs(fileNewName);
                        return;
                    }

                    fileDict[fileNewName] = fileEncoding;
                    fileName = fileNewName;
                }
                app.ScreenUpdating = false;

                // 取值
                StringBuilder sb = new StringBuilder();
                Worksheet sheet = wb.ActiveSheet;
                Range range = sheet.UsedRange;
                int row = range.Rows.Count;
                int col = range.Columns.Count;
                Console.WriteLine(range.Columns.NumberFormat);
                Console.WriteLine(range.Columns.NumberFormatLocal);

                object[,] tmp = sheet.UsedRange.Value;
                for (int i = 1; i <= row; i++)
                {
                    for (int j = 1; j <= col; j++)
                    {
                        if (j != 1)
                        {
                            sb.Append(',');
                        }

                        var obj = tmp[i, j];
                        if (obj == null)
                        {
                            continue;
                        }

                        var val = obj.ToString();
                        if (!string.IsNullOrEmpty(val))
                        {
                            sb.Append(ConvertToCsvCellString(val));
                        }
                    }

                    if (i != row)
                    {
                        sb.AppendLine(string.Empty);
                    }
                }

                range = app.ActiveCell;
                row = range.Row;
                col = range.Column;
                wb.Close(false);

                // 保存带编码的csv
                using (StreamWriter sw = new StreamWriter(fileName, false, fileEncoding))
                {
                    sw.Write(sb.ToString());
                    sw.Close();
                    sw.Dispose();
                }

                app.Workbooks.Open(fileName, null, null, XlFileFormat.xlCSV);
                sheet = app.ActiveSheet;
                sheet.UsedRange.Columns.NumberFormat = "@";
                sheet.Cells[row, col].Select();
                app.ActiveWorkbook.Saved = true;
                app.ScreenUpdating = true;
            }
        }

        private void AppOnWorkbookBeforeClose(Workbook wb, ref bool cancel)
        {
            fileDict.Remove(wb.FullName);
        }

        private void AppOnSheetActivate(object sh)
        {
            Worksheet worksheet = sh as Worksheet;
            ResetFormatConditionsHighlight(worksheet);
        }

        private void AppOnSheetSelectionChange(object sh, Range target)
        {
            if (!SettingsCsv.Default.EnableSelectHighlight)
            {
                return;
            }

            target.Calculate();
        }

        private void ResetFormatConditionsHighlight(Worksheet worksheet)
        {
            if (worksheet != null)
            {
                app.ScreenUpdating = false;
                var isSaved = app.ActiveWorkbook.Saved;

                int count = worksheet.Cells.FormatConditions.Count;
                for (int i = count; i >= 1; i--)
                {
                    FormatCondition formatCondition = worksheet.Cells.FormatConditions.Item(i);
                    if (formatCondition.Formula1 == "=CELL(\"row\")=ROW()")
                    {
                        formatCondition.Delete();
                    }
                    else if (formatCondition.Formula1 == "=AND(CELL(\"row\")=ROW(),CELL(\"col\")=COLUMN())")
                    {
                        formatCondition.Delete();
                    }
                }

                if (SettingsCsv.Default.EnableSelectHighlight)
                {
                    FormatCondition formatCondition2 = worksheet.Cells.FormatConditions.Add(XlFormatConditionType.xlExpression, null, "=AND(CELL(\"row\")=ROW(),CELL(\"col\")=COLUMN())");
                    formatCondition2.Interior.ColorIndex = 40;
                    FormatCondition formatCondition = worksheet.Cells.FormatConditions.Add(XlFormatConditionType.xlExpression, null, "=CELL(\"row\")=ROW()");
                    formatCondition.Interior.ColorIndex = 37;
                }
                app.ActiveWorkbook.Saved = isSaved;

                app.ScreenUpdating = true;
            }
        }

        private void ResetFreezeHeader(Worksheet worksheet)
        {
            if (worksheet != null && SettingsCsv.Default.EnableFreezeHeader)
            {
                string fileName = app.ActiveWorkbook.FullName;
                if (!fileName.ToLower().EndsWith(".csv"))
                {
                    return;
                }

                app.ScreenUpdating = false;

                // 根据个人使用情况，这里需要自己改代码
                int freezeRow = 0;
                int freezeCol = 0;
                Worksheet sheet = worksheet;
                if (sheet.Cells[1, 1].value == null)
                {
                    if (sheet.Cells[2, 1].value == "备注")
                    {
                        freezeRow = 2;
                        freezeCol = 2;
                    }
                    else if (sheet.Cells[4, 1].value == "备注")
                    {
                        freezeRow = 4;
                        freezeCol = 2;
                    }
                }

                if (freezeRow == 0)
                {
                    // 在这里实行自定义表头，否则认为是第一行
                    freezeRow = 1;
                }

                freezeRow++;
                freezeCol++;
                sheet.Cells[freezeRow, freezeCol].Select();
                sheet.Application.ActiveWindow.FreezePanes = true;

                app.ScreenUpdating = true;
            }
        }

        public void OnAddInsUpdate(ref Array custom)
        {
        }

        public void OnStartupComplete(ref Array custom)
        {
        }

        public void OnBeginShutdown(ref Array custom)
        {
        }

        public string GetCustomUI(string RibbonID)
        {
            return ResourceCsv.RibbonCsv;
        }

        public void About(IRibbonControl ctrl)
        {
            System.Diagnostics.Process.Start("https://blog.csdn.net/akof1314");
        }

        public bool OnGetPressedCheckBoxSave(IRibbonControl control)
        {
            return SettingsCsv.Default.EnableSaveEncode;
        }

        public bool OnGetPressedCheckBoxFreeze(IRibbonControl control)
        {
            return SettingsCsv.Default.EnableFreezeHeader;
        }

        public bool OnGetPressedCheckBoxSelect(IRibbonControl control)
        {
            return SettingsCsv.Default.EnableSelectHighlight;
        }

        public void OnCheckBoxSave(IRibbonControl ctrl, bool pressed)
        {
            SettingsCsv.Default.EnableSaveEncode = pressed;
            SettingsCsv.Default.Save();
        }

        public void OnCheckBoxFreeze(IRibbonControl ctrl, bool pressed)
        {
            SettingsCsv.Default.EnableFreezeHeader = pressed;
            SettingsCsv.Default.Save();
        }

        public void OnCheckBoxSelect(IRibbonControl ctrl, bool pressed)
        {
            SettingsCsv.Default.EnableSelectHighlight = pressed;
            SettingsCsv.Default.Save();
            ResetFormatConditionsHighlight(app.ActiveSheet);
        }

        private static string ConvertToCsvCellString(string value)
        {
            if (value.Contains(',') || value.Contains('\n'))
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }
            return value;
        }

        private Encoding GetFileEncoding(string fileName)
        {
            try
            {
                using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    byte[] bits = new byte[3];
                    fs.Read(bits, 0, 3);
                    fs.Close();

                    if (bits[0] == 0xEF && bits[1] == 0xBB && bits[2] == 0xBF)
                    {
                        return Encoding.UTF8;
                    }
                    if (bits[0] == 0xFF && bits[1] == 0xFE)
                    {
                        return Encoding.Unicode;
                    }
                    if (bits[0] == 0xFE && bits[1] == 0xFF)
                    {
                        return Encoding.BigEndianUnicode;
                    }
                }
            }
            catch (Exception)
            {
                // ignored
            }

            return Encoding.Default;
        }
    }
}

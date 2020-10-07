using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;

namespace WinMatrixEXCEL
{
    public class WinMatrixInfo
    {
        public string Account { get; set; }
        public string ComputerNM { get; set; }
        public string Email { get; set; }
        public string SDate { get; set; }
        public string EDate { get; set; }
        public string Reason { get; set; }
        public string Remark { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            string outPutPath = @"E:\日盛文件\Winmatrix最高權限批次匯入申請表單_網頁平台組\AutoCreate\";
            string sourceDataPath = @"E:\日盛文件\Winmatrix最高權限批次匯入申請表單_網頁平台組\AutoCreate\data.xls";
            var dataList = GetSourceData(sourceDataPath);
            DateTime NxtMon = GetNxtWeekMonday();
            int rowNum = 1;

            for (int i = 1; i <= 5; i++)
            {
                HSSFWorkbook wb = new HSSFWorkbook();
                ISheet sheet = wb.CreateSheet("最高權限");
                var fileNm = NxtMon.ToString("yyyy/MM/dd");

                //設定樣式
                ICellStyle headerStyle = wb.CreateCellStyle();
                IFont headerFont = wb.CreateFont();
                headerStyle.Alignment = HorizontalAlignment.Center;
                headerStyle.VerticalAlignment = VerticalAlignment.Center;
                headerStyle.FillForegroundColor = HSSFColor.Black.Index;
                headerStyle.FillPattern = FillPattern.SolidForeground;
                headerFont.FontName = "微軟正黑體";
                headerFont.Color = IndexedColors.White.Index;
                headerFont.FontHeightInPoints = 12;
                headerStyle.SetFont(headerFont);

                sheet.CreateRow(0);
                HSSFRow titleRow = (HSSFRow)sheet.GetRow(0);
                titleRow.CreateCell(0, CellType.String).SetCellValue("申請帳號");
                titleRow.CreateCell(1, CellType.String).SetCellValue("電腦名稱");
                titleRow.CreateCell(2, CellType.String).SetCellValue("使用人(電子郵件)");
                titleRow.CreateCell(3, CellType.String).SetCellValue("申請期間(起始時間)");
                titleRow.CreateCell(4, CellType.String).SetCellValue("申請期間(結束時間)");
                titleRow.CreateCell(5, CellType.String).SetCellValue("申請原因");
                titleRow.CreateCell(6, CellType.String).SetCellValue("備註");
                titleRow.GetCell(0).CellStyle = headerStyle;
                titleRow.GetCell(1).CellStyle = headerStyle;
                titleRow.GetCell(2).CellStyle = headerStyle;
                titleRow.GetCell(3).CellStyle = headerStyle;
                titleRow.GetCell(4).CellStyle = headerStyle;
                titleRow.GetCell(5).CellStyle = headerStyle;
                titleRow.GetCell(6).CellStyle = headerStyle;

                foreach (var rowData in dataList.Skip(1))
                {
                    sheet.CreateRow(rowNum);
                    HSSFRow row = (HSSFRow)sheet.GetRow(rowNum++);
                    row.CreateCell(0, CellType.String).SetCellValue(rowData.Account);
                    row.CreateCell(1, CellType.String).SetCellValue(rowData.ComputerNM);
                    row.CreateCell(2, CellType.String).SetCellValue(rowData.Email);
                    row.CreateCell(3, CellType.String).SetCellValue(fileNm + " 08:00:00");
                    row.CreateCell(4, CellType.String).SetCellValue(fileNm + " 23:59:00");
                    row.CreateCell(5, CellType.String).SetCellValue(rowData.Reason);
                    row.CreateCell(6, CellType.String).SetCellValue(rowData.Remark);
                }

                FileStream file = new FileStream(outPutPath + NxtMon.ToString("yyyyMMdd") + ".xls", FileMode.Create, FileAccess.Write);
                wb.Write(file);
                file.Close();
                SetColumnWidth(sheet);

                rowNum = 1;
                NxtMon = NxtMon.AddDays(1);
            }
        }

        private static void SetColumnWidth(ISheet st)
        {
            int cellNum = st.GetRow(0).LastCellNum;
            //for (int i = 0; i < cellNum; i++)
            //{
            //    st.AutoSizeColumn(i);
            //}

            for (int i = 0; i <= cellNum; i++)
            {
                st.AutoSizeColumn(i);
            }
            //獲取當前列的寬度，然後對比本列的長度，取最大值
            for (int columnNum = 0; columnNum <= cellNum; columnNum++)
            {
                int columnWidth = st.GetColumnWidth(columnNum) / 256;
                for (int rowNum = 1; rowNum <= st.LastRowNum; rowNum++)
                {
                    IRow currentRow;
                    //當前行未被使用過
                    if (st.GetRow(rowNum) == null)
                    {
                        currentRow = st.CreateRow(rowNum);
                    }
                    else
                    {
                        currentRow = st.GetRow(rowNum);
                    }

                    if (currentRow.GetCell(columnNum) != null)
                    {
                        ICell currentCell = currentRow.GetCell(columnNum);
                        int length = Encoding.Default.GetBytes(currentCell.ToString()).Length;
                        if (columnWidth < length)
                        {
                            columnWidth = length;
                        }
                    }
                }
                st.SetColumnWidth(columnNum, columnWidth * 256);
            }
        }

        private static DateTime GetNxtWeekMonday()
        {
            var nowDt = DateTime.Now;
            int week = Convert.ToInt32(nowDt.DayOfWeek);
            week = week == 0 ? 7 : week;
            return nowDt.AddDays(1 - week + 7);
        }

        private static List<WinMatrixInfo> GetSourceData(string path)
        {
            List<string> fieldNm = typeof(WinMatrixInfo).GetProperties().ToList().Select(n => n.Name).ToList();
            FileStream fs = File.OpenRead(path);
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            string dataJsonStr = string.Empty;

            fs.Close();
            ISheet sheet = wb.GetSheetAt(0);

            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                var row = sheet.GetRow(i);
                string rowComma = i == sheet.LastRowNum ? string.Empty : ",";

                if (row != null)
                {
                    dataJsonStr += "{" + string.Join(Environment.NewLine, row.Cells.Select((value, index) =>
                    {
                        string comma = index == row.Cells.Count - 1 ? string.Empty : ",";
                        return "\"" + fieldNm[index] + "\":\"" + value.ToString().Replace("\\","\\\\") + "\"" + comma;
                    }).ToList()) + "}" + rowComma;
                }
            }

            return new JavaScriptSerializer().Deserialize<List<WinMatrixInfo>>("[" + dataJsonStr + "]");
        }
    }
}
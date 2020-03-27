using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace Convert
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"H:\AutoPV\业务\标准规范\1.建造\分析设计\高温持久强度.xlsx";
            IWorkbook workbook = WorkbookFactory.Create(filePath);
            // 总行数
            int rows = 46;
            // 总列数
            int columns = 13;
            // 第一张表
            ISheet sheet = workbook.GetSheetAt(0);
            // 第一行
            IRow firstRow = sheet.GetRow(0);

            string newFilePath = @"H:\AutoPV\业务\标准规范\1.建造\分析设计\高温持久强度_new.xlsx";
            IWorkbook wb = new XSSFWorkbook();
            ISheet s = wb.CreateSheet("Table-1");

            int newRow = 0;
            // 从第二行开始遍历
            for (int i = 1; i< rows; i++)
            {
                IRow currentRow = sheet.GetRow(i);

                string category = currentRow.GetCell(0).StringCellValue;
                string type = currentRow.GetCell(1).StringCellValue;
                string std = currentRow.GetCell(2).StringCellValue;
                string name = currentRow.GetCell(3).StringCellValue;

                // 遍历每一列数据
                for (int j = 4; j < columns; j++)
                {
                    // 第一行对应的单元格
                    ICell firstCell = firstRow.Cells[j];

                    // 当前行单元格
                    ICell currentCell = currentRow.Cells[j];
                    double stress = currentCell.NumericCellValue;
                    if (stress > 0)
                    {
                        IRow targetRow = s.CreateRow(newRow);

                        targetRow.CreateCell(0).SetCellValue(category);
                        targetRow.CreateCell(1).SetCellValue(type);
                        targetRow.CreateCell(2).SetCellValue(std);
                        targetRow.CreateCell(3).SetCellValue(name);
                        targetRow.CreateCell(4).SetCellValue(firstCell.NumericCellValue);
                        targetRow.CreateCell(5).SetCellValue(stress);

                        newRow++;
                    }
                }
            }

            FileStream fs = File.OpenWrite(newFilePath);
            wb.Write(fs);
            fs.Close();
        }
    }
}
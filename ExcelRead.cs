using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using LoggingWithLog4Net;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelWithNPOI
{
    //Note excel should  be closed for this class to work.
    public class ExcelRead
    {
        //for work book
        IWorkbook Workbook = null;

        //for work sheet
        ISheet Sheet = null;

        //File Reader. Required for opneing excel file.
        FileStream FileStr = null;
        //construtor for initailising Excel
        public ExcelRead(String StrFullPath, String StrSheetName)
        {
            try
            {
                if (System.IO.File.Exists(StrFullPath))
                {
                    using (FileStr = new FileStream(StrFullPath, FileMode.Open, FileAccess.Read))
                    {
                        // hssfwb = new HSSFWorkbook(file);
                        Workbook = new XSSFWorkbook(FileStr);
                    }
                    Sheet = Workbook.GetSheet(StrSheetName);
                }
            }
            catch (Exception e) { }

        }

        //When calling this function update logger or change to console. In dll console is not present.
        public void PrintAll()
        {
            if (Sheet == null)
                return;
            for (int row = 0; row <= Sheet.LastRowNum; row++)
            {

                if (Sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                    //for (int col = 0; col < Sheet.GetRow(row).LastCellNum; col++)
                    //    StaticLogger.LogInfoMessage(string.Format("Row {0} = {1}", row, GetCellValue(row, col)));
                }
            }
        }

        //Close work book and file stream.
        public void CloseAllConnection()
        {
            if (FileStr != null)
                FileStr.Close();
            if (Workbook != null)
                Workbook.Close();
        }

        //get row number of a particular string from any column
        public int GetRowNumber(String StrValue, int ColumnNumber)
        {
            if (Sheet == null || ColumnNumber<0)
                return -1;
            int row = -1;
            for (int rowIndex = 0; row <= Sheet.LastRowNum; row++)
            {
                if (Sheet.GetRow(rowIndex) != null) //null is when the row only contains empty cells 
                {
                    if (GetCellValue(row, ColumnNumber).Equals(StrValue))
                        return rowIndex;
                }
            }
            return row;
        }

        //Get total columns inside a row
        private int GetTotalColumns(int RowNum)
        {
            if (Sheet == null)
                return -1;
            return Sheet.GetRow(RowNum).LastCellNum;
        }

        //Get a list of string with all row data
        public List<String> GetRowData(int RowNumber)
        {
            List<String> Row = new List<string>();
            if (Sheet != null)
            {
                if (RowNumber > Sheet.LastRowNum || Sheet.GetRow(RowNumber) == null)
                {
                    return Row;
                }
                else
                {
                    for (int col = 0; col < Sheet.GetRow(RowNumber).LastCellNum; col++)
                        Row.Add(GetCellValue(RowNumber, col));
                }
            }
            return Row;
        }

        //When calling this function update logger or change to console. In dll console is not present.
        public void PrintCellType()
        {
            if (Sheet == null)
                return;
            for (int row = 0; row <= Sheet.LastRowNum; row++)
            {
                if (Sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                {
                   // for (int col = 0; col < Sheet.GetRow(row).LastCellNum; col++)
                     //   StaticLogger.LogInfoMessage(string.Format("Row {0} = {1}", row, GetCellType(row, col)));
                }
            }
        }
        
        //Get cell type of any cell
        private CellType GetCellType(int Row, int Col)
        {
            try
            {
                return Sheet.GetRow(Row).GetCell(Col).CellType;
            }
            catch (Exception)
            {
                return CellType.Error;
            }
        }

        //Get cell value
        public string GetCellValue(int Row, int Col)
        {
            try
            {
                switch (GetCellType(Row, Col))
                {

                    case CellType.String:
                        return Sheet.GetRow(Row).GetCell(Col).StringCellValue.ToString();
                    case CellType.Numeric://date is displayed as numeric
                        return Sheet.GetRow(Row).GetCell(Col).NumericCellValue.ToString();
                    case CellType.Boolean:
                        return Sheet.GetRow(Row).GetCell(Col).BooleanCellValue.ToString();
                    //CellFormula.ToString(); will return actual formula eg today()
                    case CellType.Formula:
                        if (Sheet.GetRow(Row).GetCell(Col).CachedFormulaResultType == CellType.Numeric)
                            return Sheet.GetRow(Row).GetCell(Col).NumericCellValue.ToString();
                        else if (Sheet.GetRow(Row).GetCell(Col).CachedFormulaResultType == CellType.String)
                            return Sheet.GetRow(Row).GetCell(Col).StringCellValue.ToString();
                        else if (Sheet.GetRow(Row).GetCell(Col).CachedFormulaResultType == CellType.Boolean)
                            return Sheet.GetRow(Row).GetCell(Col).BooleanCellValue.ToString();
                        else
                            return Workbook.GetCreationHelper().CreateFormulaEvaluator().Evaluate(Sheet.GetRow(Row).GetCell(Col)).FormatAsString();

                    default:
                        return "";
                }
            }
            catch (Exception)
            {
                return "";
            }
        }

        //Get all columns data
        public List<string> GetColumnData(int ColNum,int StartRow = 1)
        {
            List<String> Row = new List<string>();
            if (Sheet != null && ColNum >= 0)
            {
                for (int row = StartRow; row <= Sheet.LastRowNum && ColNum<=GetTotalColumns(row); row++)
                {
                    Row.Add(GetCellValue(row, ColNum));
                }
            }
            return Row;
        }

        //Get all row data by index in a dictionary
        public Dictionary<string,string> GetRowDataInDict(int Row)
        {
            Dictionary<string, string> AllRowData = new Dictionary<string, string>();
            try
            {
                List<string> RowHeader = GetRowData(0);
                List<string> RowData = GetRowData(Row);
                for (int i = 0; i < RowHeader.Count; i++)
                {
                    AllRowData.Add(RowHeader[i], RowData[i]);
                }
            }
            catch (Exception e) { }
           
            return AllRowData;
        }

        //Get all column data with respect any column as index
        public Dictionary<string, string> GetColumenDataInDict(int ColNumber, int KeyCol=0)
        {
            Dictionary<string, string> AllRowData = new Dictionary<string, string>();
            try
            {
                List<string> ColumneKey = GetColumnData(KeyCol);
                List<string> ColumnValue = GetColumnData(ColNumber);
                for (int i = 0; i < ColumneKey.Count; i++)
                {
                    AllRowData.Add(ColumneKey[i], ColumnValue[i]);
                }
            }
            catch (Exception e) { }

            return AllRowData;
        }    
    }
}



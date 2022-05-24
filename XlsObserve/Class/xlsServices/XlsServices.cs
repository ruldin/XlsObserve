using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.POIFS.FileSystem;
using System.Data;
using XlsObserve.Class.logs;

namespace XlsObserve.Class.xlsServices
{
    public class XlsServices
    {

        /// <summary>
        /// This method Add data from DataTable as source
        /// 
        /// </summary>
        /// <param name="pathDoXls"></param>
        /// <param name="dtXls"></param>
        /// <param name="indexSheet"></param>
        /// <returns></returns>
        public int AppendData(string pathDoXls, DataTable dtXls, int indexSheet)
        {
            //to use both types of files (xls / xlsx)
            //dependency injection is used
            IWorkbook book;
            int rowsProcessed = 0;

            using (FileStream file1 = new FileStream(pathDoXls, FileMode.Open, FileAccess.Read))
            {
                if (Path.GetExtension(pathDoXls).ToLower().Equals(".xlsx"))
                {
                    //xlsx format
                    book = new XSSFWorkbook(file1);
                }
                else //xls format
                {
                    book = new HSSFWorkbook(file1);
                }


                //ISheet sheet = hssfwb.GetSheet("Hoja1"); //change by name on config
                ISheet sheet = book.GetSheetAt(indexSheet); //change by name on config
                int lastRowNum = sheet.LastRowNum;
                foreach (DataRow dr in dtXls.Rows)
                {
                    rowsProcessed++;
                    IRow worksheetRow = sheet.CreateRow(lastRowNum++);
                    for (int cols = 0; cols < dtXls.Columns.Count; cols++)
                    {
                        ICell cell = worksheetRow.CreateCell(cols);
                        cell.SetCellValue(dr[cols].ToString());
                    }
                }
                FileStream file = new FileStream(pathDoXls, FileMode.Create);
                book.Write(file);
                file.Close();
            }
            return rowsProcessed;
        }



        /// <summary>
        /// this method obtains the data from the sheet and 
        /// transforms it into a DataTable
        /// </summary>
        /// <param name="pFilePath"></param>
        /// <param name="pSheetIndex"></param>
        /// <returns></returns>
        public DataTable Excel_To_DataTable(string pFilePath, int pSheetIndex)
        {
            DataTable Tabla = null;
            try
            {
                if (System.IO.File.Exists(pFilePath))
                {

                    IWorkbook workbook = null;         
                    ISheet worksheet = null;
                    string first_sheet_name = "";

                    using (FileStream FS = new FileStream(pFilePath, FileMode.Open, FileAccess.Read))
                    {
                        workbook = WorkbookFactory.Create(FS);
                        worksheet = workbook.GetSheetAt(pSheetIndex);
                        first_sheet_name = worksheet.SheetName;         

                        Tabla = new DataTable(first_sheet_name);
                        Tabla.Rows.Clear();
                        Tabla.Columns.Clear();

                        // read row by row
                        for (int rowIndex = 0; rowIndex <= worksheet.LastRowNum; rowIndex++)
                        {
                            DataRow NewReg = null;
                            IRow row = worksheet.GetRow(rowIndex);
                            IRow row2 = null;
                            IRow row3 = null;

                            if (rowIndex == 0)
                            {
                                row2 = worksheet.GetRow(rowIndex + 1); 
                                row3 = worksheet.GetRow(rowIndex + 2); 
                            }

                            if (row != null) 
                            {
                                if (rowIndex > 0) NewReg = Tabla.NewRow();

                                int colIndex = 0;
                                
                                foreach (ICell cell in row.Cells)
                                {
                                    object valorCell = null;
                                    string cellType = "";
                                    string[] cellType2 = new string[2];

                                    if (rowIndex == 0) //take title rows at 0
                                    {
                                        for (int i = 0; i < 2; i++)
                                        {
                                            ICell cell2 = null;
                                            if (i == 0) { cell2 = row2.GetCell(cell.ColumnIndex); }
                                            else { cell2 = row3.GetCell(cell.ColumnIndex); }

                                            if (cell2 != null)
                                            {
                                                switch (cell2.CellType)
                                                {
                                                    case CellType.Blank: break;
                                                    case CellType.Boolean: cellType2[i] = "System.Boolean"; break;
                                                    case CellType.String: cellType2[i] = "System.String"; break;
                                                    case CellType.Numeric:
                                                        if (HSSFDateUtil.IsCellDateFormatted(cell2)) { cellType2[i] = "System.DateTime"; }
                                                        else
                                                        {
                                                            cellType2[i] = "System.Double";  //valorCell = cell2.NumericCellValue;
                                                        }
                                                        break;

                                                    case CellType.Formula:
                                                        bool continuar = true;
                                                        switch (cell2.CachedFormulaResultType)
                                                        {
                                                            case CellType.Boolean: cellType2[i] = "System.Boolean"; break;
                                                            case CellType.String: cellType2[i] = "System.String"; break;
                                                            case CellType.Numeric:
                                                                if (HSSFDateUtil.IsCellDateFormatted(cell2)) { cellType2[i] = "System.DateTime"; }
                                                                else
                                                                {
                                                                    try
                                                                    {
                                                                        //check if is boolean
                                                                        if (cell2.CellFormula == "TRUE()") { cellType2[i] = "System.Boolean"; continuar = false; }
                                                                        if (continuar && cell2.CellFormula == "FALSE()") { cellType2[i] = "System.Boolean"; continuar = false; }
                                                                        if (continuar) { cellType2[i] = "System.Double"; continuar = false; }
                                                                    }
                                                                    catch { }
                                                                }
                                                                break;
                                                        }
                                                        break;
                                                    default:
                                                        cellType2[i] = "System.String"; break;
                                                }
                                            }
                                        }

                                        //Resolve data type
                                        if (cellType2[0] == cellType2[1]) { cellType = cellType2[0]; }
                                        else
                                        {
                                            if (cellType2[0] == null) cellType = cellType2[1];
                                            if (cellType2[1] == null) cellType = cellType2[0];
                                            if (cellType == "") cellType = "System.String";
                                        }

                                        //get column name
                                        string colName = "Column_{0}";
                                        try { colName = cell.StringCellValue; }
                                        catch { colName = string.Format(colName, colIndex); }

                                        //check column name is unique
                                        foreach (DataColumn col in Tabla.Columns)
                                        {
                                            if (col.ColumnName == colName) colName = string.Format("{0}_{1}", colName, colIndex);
                                        }

                                        //add fields to table
                                        DataColumn codigo = new DataColumn(colName, System.Type.GetType(cellType));
                                        Tabla.Columns.Add(codigo); colIndex++;
                                    }
                                    else
                                    {
                                        
                                        switch (cell.CellType)
                                        {
                                            case CellType.Blank: valorCell = DBNull.Value; break;
                                            case CellType.Boolean: valorCell = cell.BooleanCellValue; break;
                                            case CellType.String: valorCell = cell.StringCellValue; break;
                                            case CellType.Numeric:
                                                if (HSSFDateUtil.IsCellDateFormatted(cell)) { valorCell = cell.DateCellValue; }
                                                else { valorCell = cell.NumericCellValue; }
                                                break;
                                            case CellType.Formula:
                                                switch (cell.CachedFormulaResultType)
                                                {
                                                    case CellType.Blank: valorCell = DBNull.Value; break;
                                                    case CellType.String: valorCell = cell.StringCellValue; break;
                                                    case CellType.Boolean: valorCell = cell.BooleanCellValue; break;
                                                    case CellType.Numeric:
                                                        if (HSSFDateUtil.IsCellDateFormatted(cell)) { valorCell = cell.DateCellValue; }
                                                        else { valorCell = cell.NumericCellValue; }
                                                        break;
                                                }
                                                break;
                                            default: valorCell = cell.StringCellValue; break;
                                        }
                                        
                                        if (cell.ColumnIndex <= Tabla.Columns.Count - 1) NewReg[cell.ColumnIndex] = valorCell;
                                    }
                                }
                            }
                            if (rowIndex > 0) Tabla.Rows.Add(NewReg);
                        }
                        Tabla.AcceptChanges();
                    }
                }
                else
                {
                    ClsLogs.ErrorLog($"Xls to DataTable, File not found: {pFilePath}");

                }
            }
            catch (Exception exep)
            {
                ClsLogs.ErrorLog($"Xls to DT Error: ");
            }
            return Tabla;
        }




    }
}

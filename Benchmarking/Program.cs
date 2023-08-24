using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using ClosedXML.Excel;
using CsvHelper;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System.Data;
using System.Globalization;
using System.Reflection;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
var summary = BenchmarkRunner.Run<BenchmarkDriver>();

public class BenchmarkDriver
{

    private IEnumerable<DataInstance> data;

    public BenchmarkDriver()
    {
        using (var reader = new StreamReader("data.csv"))
        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
        {
            data = csv.GetRecords<DataInstance>();
        }
    }

    [Benchmark]
    public void NpoiTest()
    {
        var dataTable = CreateDataTableFromIEnumerable(data);
        using (IWorkbook workbook = new XSSFWorkbook())
        {
            CreateSheetFromDataTable(workbook, 0, dataTable);

            //FileStream sw = File.Create("File.xlsx");
            //workbook.Write(sw, false);
            //sw.Close();
        }
    }

    [Benchmark]
    public void EPPlusTest()
    {
        using (var p = new ExcelPackage())
        {
            var rowIndex = 2;
            var ws = p.Workbook.Worksheets.Add("export");
            ws.Cells[1, 1].Value = "hlpi_name";
            ws.Cells[1, 2].Value = "series_ref";
            ws.Cells[1, 3].Value = "quarter";
            ws.Cells[1, 4].Value = "hlpi";
            ws.Cells[1, 5].Value = "nzhec";
            ws.Cells[1, 6].Value = "nzhec_name";
            ws.Cells[1, 7].Value = "nzhec_short";
            ws.Cells[1, 8].Value = "level";
            ws.Cells[1, 9].Value = "index";
            ws.Cells[1, 10].Value = "changeQ";
            ws.Cells[1, 11].Value = "changeA";
            foreach (var instance in data)
            {
                ws.Cells[rowIndex, 1].Value = instance.hlpi_name;
                ws.Cells[rowIndex, 2].Value = instance.series_ref;
                ws.Cells[rowIndex, 3].Value = instance.quarter;
                ws.Cells[rowIndex, 4].Value = instance.hlpi;
                ws.Cells[rowIndex, 5].Value = instance.nzhec;
                ws.Cells[rowIndex, 6].Value = instance.nzhec_name;
                ws.Cells[rowIndex, 7].Value = instance.nzhec_short;
                ws.Cells[rowIndex, 8].Value = instance.level;
                ws.Cells[rowIndex, 9].Value = instance.index;
                ws.Cells[rowIndex, 10].Value = instance.changeQ;
                ws.Cells[rowIndex, 11].Value = instance.changeA;
                rowIndex++;
            }
        }
    }

    [Benchmark]
    public void ClosedXmlTest()
    {
        using (var workbook = new XLWorkbook())
        {
            var rowIndex = 2;
            var ws = workbook.Worksheets.Add("export");
            ws.Cell(1, 1).Value = "hlpi_name";
            ws.Cell(1, 2).Value = "series_ref";
            ws.Cell(1, 3).Value = "quarter";
            ws.Cell(1, 4).Value = "hlpi";
            ws.Cell(1, 5).Value = "nzhec";
            ws.Cell(1, 6).Value = "nzhec_name";
            ws.Cell(1, 7).Value = "nzhec_short";
            ws.Cell(1, 8).Value = "level";
            ws.Cell(1, 9).Value = "index";
            ws.Cell(1, 10).Value = "changeQ";
            ws.Cell(1, 11).Value = "changeA";
            foreach (var instance in data)
            {
                ws.Cell(rowIndex, 1).Value = instance.hlpi_name;
                ws.Cell(rowIndex, 2).Value = instance.series_ref;
                ws.Cell(rowIndex, 3).Value = instance.quarter;
                ws.Cell(rowIndex, 4).Value = instance.hlpi;
                ws.Cell(rowIndex, 5).Value = instance.nzhec;
                ws.Cell(rowIndex, 6).Value = instance.nzhec_name;
                ws.Cell(rowIndex, 7).Value = instance.nzhec_short;
                ws.Cell(rowIndex, 8).Value = instance.level;
                ws.Cell(rowIndex, 9).Value = instance.index;
                ws.Cell(rowIndex, 10).Value = instance.changeQ;
                ws.Cell(rowIndex, 11).Value = instance.changeA;
                rowIndex++;
            }
        }
    }

    private static DataTable CreateDataTableFromIEnumerable<T>(IEnumerable<T> data)
    {
        var tb = new DataTable(typeof(T).Name);
        PropertyInfo[] props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);

        foreach (PropertyInfo prop in props)
        {
            Type t = GetCoreType(prop.PropertyType);
            tb.Columns.Add(prop.Name, t);
        }

        foreach (T item in data)
        {
            var values = new object[props.Length];

            for (int i = 0; i < props.Length; i++)
            {
                values[i] = props[i].GetValue(item, null);
            }

            tb.Rows.Add(values);
        }

        return tb;
    }

    private static bool IsNullable(Type t)
    {
        return !t.IsValueType || (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>));
    }

    private static Type GetCoreType(Type t)
    {
        if (t != null && IsNullable(t))
        {
            if (!t.IsValueType)
            {
                return t;
            }
            else
            {
                return Nullable.GetUnderlyingType(t);
            }
        }
        else
        {
            return t;
        }
    }

    private static void CreateSheetFromDataTable(IWorkbook workbook, int dataTableIndex, DataTable dataTable)
    {
        var tableName = string.IsNullOrEmpty(dataTable.TableName) ? $"Sheet {dataTableIndex}" : dataTable.TableName;
        var sheet = (XSSFSheet)workbook.CreateSheet(tableName);
        var columnCount = dataTable.Columns.Count;
        var rowCount = dataTable.Rows.Count;

        // add column headers
        var row = sheet.CreateRow(0);
        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
        {
            var col = dataTable.Columns[columnIndex];
            row.CreateCell(columnIndex).SetCellValue(col.ColumnName);
        }

        // add data rows
        for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            var dataRow = dataTable.Rows[rowIndex];
            var sheetRow = sheet.CreateRow(rowIndex + 1);
            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                sheetRow.CreateCell(columnIndex).SetCellValue(dataRow[columnIndex].ToString());
            }
        }

        // format the cell range as a table
        // note: if id, name, displayName are not set, Excel will not support the table
        // note: if id=0, Excel will not support the table
        XSSFTable xssfTable = sheet.CreateTable();
        CT_Table ctTable = xssfTable.GetCTTable();
        AreaReference myDataRange = new AreaReference(new CellReference(0, 0), new CellReference(rowCount, columnCount - 1));
        var tableId = uint.Parse((dataTableIndex + 1).ToString());
        ctTable.@ref = myDataRange.FormatAsString();
        ctTable.id = tableId;
        ctTable.name = $"Table{tableId}";
        ctTable.displayName = $"Table{tableId}";
        ctTable.tableStyleInfo = new CT_TableStyleInfo();
        ctTable.tableStyleInfo.name = "TableStyleMedium2"; // TableStyleMedium2 is one of XSSFBuiltinTableStyle
        ctTable.tableStyleInfo.showRowStripes = true;
        ctTable.tableColumns = new CT_TableColumns();
        ctTable.tableColumns.tableColumn = new List<CT_TableColumn>();
        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
        {
            var col = dataTable.Columns[columnIndex];
            var colId = uint.Parse((columnIndex + 1).ToString());
            // note: if id=0, Excel will not support the table
            ctTable.tableColumns.tableColumn.Add(new CT_TableColumn() { id = colId, name = col.ColumnName });
        }

        // turn on filtering
        ctTable.autoFilter = new CT_AutoFilter();
        ctTable.autoFilter.@ref = myDataRange.FormatAsString();

        // auto size columns
        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
        {
            sheet.AutoSizeColumn(columnIndex);
        }
        for (var columnIndex = 0; columnIndex < columnCount; columnIndex++)
        {
            // make room for the filter button and add a bit more
            var colWidth = sheet.GetColumnWidth(columnIndex);
            sheet.SetColumnWidth(columnIndex, colWidth + 1500);
        }
    }
}

public class DataInstance
{
    public string hlpi_name { get; set; } = string.Empty;
    public string series_ref { get; set; } = string.Empty;
    public string quarter { get; set; } = string.Empty;
    public string hlpi { get; set; } = string.Empty;
    public float nzhec { get; set; }
    public string nzhec_name { get; set; } = string.Empty;
    public string nzhec_short { get; set; } = string.Empty;
    public string level { get; set; } = string.Empty;
    public int index { get; set; }
    public string changeQ { get; set; } = string.Empty;
    public string changeA { get; set; } = string.Empty;
}

using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using ConvertEXCELtoINSERTSQL.Model;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ConvertEXCELtoINSERTSQL.Controllers
{
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(ILogger<ExcelController> logger)
        {
            _logger = logger;
        }

        [HttpPost("UploadExcel")]
        public async Task<IActionResult> UploadExcel(DataModel model)
        {
            if (string.IsNullOrEmpty(model.File) || string.IsNullOrEmpty(model.FileExtension))
            {
                return BadRequest("Invalid File");
            }

            try
            {
                var dt = ImportExcel(model);
                var sqlScripts = GenerateSqlScripts(dt);
               
                var concatenatedScripts = string.Join(" ", sqlScripts.Select(script => RemoveExtraWhitespace(script)));
                                
                return Ok(concatenatedScripts);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing the Excel file");
                return StatusCode(500, "Error processing the file");
            }
        }
        
        private string RemoveExtraWhitespace(string input)
        {
            return input.Replace("\r", "").Replace("\n", "").Trim();
        }
               
        private DataTable ImportExcel(DataModel model)
        {
            DataTable dt = new DataTable();
            var bytes = Convert.FromBase64String(model.File);
            using (XLWorkbook workBook = new XLWorkbook(new MemoryStream(bytes)))
            {
                var workSheet = workBook.Worksheet(1);

                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        dt.Rows.Add();
                        int i = 0;
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }
                }
            }

            return dt;
        }

        private List<string> GenerateSqlScripts(DataTable dt)
        {
            var scripts = new List<string>();
            var tableName = "ExcelTable";

            var createTableSql = GenerateCreateTableScript(dt, tableName);
            scripts.Add(createTableSql);

            var insertSql = GenerateInsertQueries(dt, tableName);
            scripts.AddRange(insertSql);

            return scripts;
        }

        
        private string GenerateCreateTableScript(DataTable dt, string tableName)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"CREATE TABLE {tableName} (");
            
            foreach (DataColumn column in dt.Columns)
            {
                string sqlType = GivenSqlTypeForColumn(dt, column);
                sb.AppendLine($"    {column.ColumnName} {sqlType},");
            }

            sb.Length -= 3;
            sb.Append(");");

            return sb.ToString();
        }

        private string GivenSqlTypeForColumn(DataTable dt, DataColumn column)
        {
            var columnValues = dt.AsEnumerable().Select(row => row[column].ToString()).ToList();

            bool hasOnlyNumbers = columnValues.All(value => double.TryParse(value, out _));
            bool hasOnlyIntegers = columnValues.All(value => int.TryParse(value, out _));
                        
            bool hasDecimals = columnValues.Any(value => value.Contains('.'));

            if (hasOnlyIntegers)
                return "INT";
            if (hasOnlyNumbers && hasDecimals)
                return "DECIMAL(18,2)";
            if (hasOnlyNumbers)
                return "VARCHAR(10)";

            return "VARCHAR(512)";
        }

        //This method creates INSERT SQL statements for each row in the DataTable ---- By Harshit
        private List<string> GenerateInsertQueries(DataTable dt, string tableName)
        {
            var queries = new List<string>();

            foreach (DataRow row in dt.Rows)
            {                
                var columns = dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToList();
                var values = row.ItemArray.Select(value => value == DBNull.Value ? "NULL" : $"'{value.ToString().Replace("'", "''")}'").ToList();

                var columnList = string.Join(", ", columns);
                var valueList = string.Join(", ", values);

                var sql = $"INSERT INTO {tableName} ({columnList}) VALUES ({valueList});";
                queries.Add(sql);
            }

            return queries;
        }
    }
}

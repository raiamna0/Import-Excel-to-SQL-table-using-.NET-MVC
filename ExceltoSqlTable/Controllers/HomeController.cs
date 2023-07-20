using ExceltoSqlTable.Models;
using Newtonsoft.Json;
using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Web;
using System.Web.Mvc;

public class HomeController : Controller
{
	string cs = ConfigurationManager.ConnectionStrings["dbcs"].ConnectionString;
	public ActionResult Upload()
	{
		return View();
	}

	[HttpPost]
	public ActionResult Upload(HttpPostedFileBase fileUpload)
	{
		if (fileUpload != null && fileUpload.ContentLength > 0)
		{
			string filePath = Path.Combine(Server.MapPath("~/UploadedFiles"), Path.GetFileName(fileUpload.FileName));
			fileUpload.SaveAs(filePath);

			DataTable excelData = ExtractExcelData(filePath);
			if (excelData != null)
			{
				string createTableQuery = GetCreateTableQuery(excelData);
				CreateSQLTable(createTableQuery);
				SaveDataToSQLTable(excelData);
				var model = new ExcelDataModel
				{
					FilePath = filePath,
					ExcelData = excelData
				};

				// Save the extracted column names and data in the model to the database table.
				// (Note: Implement the database connection and table creation code here.)

				var columnNames = GetColumnNames(excelData);
				ViewBag.ColumnNames = JsonConvert.SerializeObject(columnNames);

				return View("Success", model);
			}
		}

		return View();
	}

	
	private string GetCreateTableQuery(DataTable dataTable)
	{
		// Generate the CREATE TABLE query based on the DataTable columns and inferred data types
		string createTableQuery = "CREATE TABLE YourTableName (";

		foreach (DataColumn column in dataTable.Columns)
		{
			string dataType = GetSQLDataType(column);
			createTableQuery += $"[{column.ColumnName}] {dataType}, ";
		}

		createTableQuery = createTableQuery.TrimEnd(',', ' ') + ")";

		return createTableQuery;
	}
	private string GetSQLDataType(DataColumn column)
	{
		// Infer the SQL data type based on the Excel data type in the DataTable column
		Type dataType = column.DataType;

		if (dataType == typeof(int) || dataType == typeof(long))
			return "INT";
		else if (dataType == typeof(decimal) || dataType == typeof(double))
			return "DECIMAL(18, 2)";
		else if (dataType == typeof(DateTime))
			return "DATETIME";
		else
			return "NVARCHAR(255)"; // Default to NVARCHAR for other data types
	}
	private void CreateSQLTable(string createTableQuery)
	{
		using (SqlConnection connection = new SqlConnection(cs))
		{
			connection.Open();

			using (SqlCommand command = new SqlCommand(createTableQuery, connection))
			{
				command.ExecuteNonQuery();
			}

			connection.Close();
		}
	}
	private DataTable ExtractExcelData(string filePath)
	{
		string connString = string.Empty;
		string extension = Path.GetExtension(filePath);

		if (extension.ToLower() == ".xls")
		{
			connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=YES;'";
		}
		else if (extension.ToLower() == ".xlsx")
		{
			connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=YES;'";
		}

		using (OleDbConnection connection = new OleDbConnection(connString))
		{
			connection.Open();
			DataTable dtSchema = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
			string sheetName = dtSchema.Rows[0]["TABLE_NAME"].ToString();
			OleDbDataAdapter dataAdapter = new OleDbDataAdapter("SELECT * FROM [" + sheetName + "]", connection);
			DataTable dt = new DataTable();
			dataAdapter.Fill(dt);

			return dt;
		}
	}
	private string[] GetColumnNames(DataTable table)
	{
		string[] columnNames = new string[table.Columns.Count];
		for (int i = 0; i < table.Columns.Count; i++)
		{
			columnNames[i] = table.Columns[i].ColumnName;
		}
		return columnNames;
	}
	private void SaveDataToSQLTable(DataTable dataTable)
	{
		using (SqlConnection connection = new SqlConnection(cs))
		{
			connection.Open();

			using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection))
			{
				bulkCopy.DestinationTableName = "YourTableName"; // Replace with your actual table name
				bulkCopy.WriteToServer(dataTable);
			}

			connection.Close();
		}
	}
}

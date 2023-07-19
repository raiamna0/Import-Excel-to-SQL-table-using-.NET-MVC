using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace ExceltoSqlTable.Models
{
	public class ExcelDataModel
	{
		public string FilePath { get; set; }
		public DataTable ExcelData { get; set; }
	}

}
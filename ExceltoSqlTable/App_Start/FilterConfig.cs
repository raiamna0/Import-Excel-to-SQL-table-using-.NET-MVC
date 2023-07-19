using System.Web;
using System.Web.Mvc;

namespace ExceltoSqlTable
{
	public class FilterConfig
	{
		public static void RegisterGlobalFilters(GlobalFilterCollection filters)
		{
			filters.Add(new HandleErrorAttribute());
		}
	}
}

using System;
using System.Runtime.CompilerServices;

namespace WorkReportConverter.Model
{
	public class ColumnAttribute: Attribute
	{
		public ColumnAttribute(int order, [CallerMemberName]string columnName = "unnamed")
		{
			ColumnIndex = order;
			ColumnName = columnName;
		}

		public int ColumnIndex { get; }
		public string ColumnName { get; }
	}
}

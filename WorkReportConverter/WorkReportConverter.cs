using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using WorkReportConverter.Model;

namespace WorkReportConverter
{
	public class WorkReportConverter
	{
		private readonly Configuration configuration;

		public WorkReportConverter(Configuration configuration)
		{
			this.configuration = configuration;
		}

		public void ConvertWorkReport()
		{
			using (ExcelEngine excelEngine = new ExcelEngine())
			{
				var workReportRows = new List<WorkReportRow>();
				IApplication app = excelEngine.Excel;

				using (var stream = new FileStream(configuration.Source, FileMode.Open))
				{
					var workbook = app.Workbooks.Open(stream, ExcelOpenType.Automatic);
					var worksheet = workbook.Worksheets[0];
					worksheet.Name = "Timesheet";
					var rows = worksheet.Rows;
					foreach (var row in rows.Skip(1))
					{
						var workReportRow = MakeWorkReportRow(row);
						if (workReportRow.Day.Year == configuration.Year && workReportRow.Month == configuration.Month)
							workReportRows.Add(workReportRow);
					}
					workbook.Close();
				}

				var monthFilePostfix = configuration.Year + ("0" + configuration.Month)[^2..];
				var destinationFileName = configuration.Destination + "_" + monthFilePostfix + ".xlsx";
				using (var stream = new FileStream(destinationFileName, FileMode.OpenOrCreate))
				{
					IWorkbook workbook = app.Workbooks.Create();
					IWorksheet worksheet = workbook.Worksheets[0];
					var properties = typeof(WorkReportRow).GetProperties();

					foreach (var property in properties)
					{
						var columnAttribute = property.GetCustomAttributes(typeof(ColumnAttribute), false).FirstOrDefault() as ColumnAttribute;
						if (columnAttribute != null)
							worksheet.SetValueRowCol(columnAttribute.ColumnName, 1, columnAttribute.ColumnIndex + 1);
					}

					for (var i = 0; i < workReportRows.Count; i++)
					{
						foreach (var property in properties)
						{
							var columnAttribute = property.GetCustomAttributes(typeof(ColumnAttribute), false).FirstOrDefault() as ColumnAttribute;
							if (columnAttribute != null)
							{
								var value = property.GetValue(workReportRows[i]);
								worksheet.SetValueRowCol(value, i + 2, columnAttribute.ColumnIndex + 1);
							}
						}
					}

					workbook.SaveAs(stream);
					workbook.Close();
				}
			}
		}

		private WorkReportRow MakeWorkReportRow(IRange row)
		{
			var parsedDescription = ParseDescription(row.Cells[1].Text);
			return new WorkReportRow()
			{
				Category = configuration.Category,
				Day = row.Cells[0].DateTime,
				Effort = row.Cells[3].Number,
				Month = row.Cells[0].DateTime.Month,
				PBIID = parsedDescription[1],
				TaskID = parsedDescription[2],
				Sprint = parsedDescription[0],
				Task = parsedDescription[3],
				Project = configuration.Project,
				ResourceName = configuration.ResourceName,
				Role = configuration.Role,
				Team = configuration.Team,
				Week = GetWeekNumberOfMonth(row.Cells[0].DateTime)
			};
		}

		private string regexPBI = "\\(PBI:(?'pbi'[^)]+)\\)";
		private string regexTASK = "\\(TASK:(?'task'[^)]+)\\)";
		private string regexSPRINT = "\\(SPRINT:(?'sprint'[^)]+)\\)";

		private string[] ParseDescription(string description)
		{
			var pbis = Regex.Match(description, regexPBI)?.Groups["pbi"]?.Value;
			var tasks = Regex.Match(description, regexTASK)?.Groups["task"]?.Value;
			var sprint = Regex.Match(description, regexSPRINT)?.Groups["sprint"]?.Value;
			var rest = description;
			rest = Regex.Replace(rest, regexPBI, "");
			rest = Regex.Replace(rest, regexTASK, "");
			rest = Regex.Replace(rest, regexSPRINT, "");
			return new string[] { sprint, pbis, tasks, rest }.Select(x => x.Trim()).ToArray();
		}

		private int GetWeekNumberOfMonth(DateTime date)
		{
			date = date.Date;
			DateTime firstMonthDay = new DateTime(date.Year, date.Month, 1);
			DateTime firstMonthMonday = firstMonthDay.AddDays((DayOfWeek.Monday + 7 - firstMonthDay.DayOfWeek) % 7);
			if (firstMonthMonday > date)
			{
				firstMonthDay = firstMonthDay.AddMonths(-1);
				firstMonthMonday = firstMonthDay.AddDays((DayOfWeek.Monday + 7 - firstMonthDay.DayOfWeek) % 7);
			}
			return (date - firstMonthMonday).Days / 7 + 1;
		}
	}
}

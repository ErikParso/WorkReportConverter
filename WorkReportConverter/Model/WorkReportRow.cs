using System;

namespace WorkReportConverter.Model
{
	public class WorkReportRow
	{
		[Column(0)]
		public string Name { get; set; }

		[Column(1)]
		public DateTime Date { get; set; }

		[Column(2)]
		public string Sprint { get; set; }

		[Column(3)]
		public string Role { get; set; }

		[Column(4)]
		public string Project { get; set; }

		[Column(5)]
		public string Category { get; set; }

		[Column(6, "PBI/BUG")]
		public string PBIID { get; set; }

		[Column(7, "Task ID")]
		public string TaskID { get; set; }


		[Column(8, "Task description")]
		public string TaskDescription { get; set; }

		[Column(9)]
		public double Effort { get; set; }
	}
}

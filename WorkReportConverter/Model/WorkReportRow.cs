using System;

namespace WorkReportConverter.Model
{
	public class WorkReportRow
	{
		[Column(0, "Resource name")]
		public string ResourceName { get; set; }

		[Column(1)]
		public DateTime Day { get; set; }

		[Column(2)]
		public string Project { get; set; }

		[Column(3)]
		public string Task { get; set; }

		[Column(4)]
		public double Effort { get; set; }

		[Column(5)]
		public string Role { get; set; }

		[Column(6)]
		public int Week { get; set; }

		[Column(7)]
		public int Month { get; set; }

		[Column(8)]
		public string Category { get; set; }

		[Column(9)]
		public string Team { get; set; }

		[Column(10)]
		public string Sprint { get; set; }

		[Column(11, "PBI/BUG")]
		public string PBIID { get; set; }

		[Column(12, "Task ID")]
		public string TaskID { get; set; }
	}
}

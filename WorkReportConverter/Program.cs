using System;
using System.IO;
using System.Text.Json;
using WorkReportConverter.Model;

namespace WorkReportConverter
{
	class Program
	{
		static void Main(string[] args)
		{
			try
			{
				var configuration = JsonSerializer.Deserialize<Configuration>(File.ReadAllText("config.json"));
				var workReportConverter = new WorkReportConverter(configuration);
				workReportConverter.ConvertWorkReport();
			}
			catch (Exception e)
			{
				Console.WriteLine(e.Message);
				Console.ReadLine();
			}
		}
	}
}

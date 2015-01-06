using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using CommandLine;
using CommandLine.Text;
using CsvHelper;
using CsvHelper.Configuration;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace GCBirthdays
{
	public class Birhday
	{
		public string Name { get; set; }
		private DateTime _date;

		public DateTime Date
		{
			get { return _date.Date; }
			set { _date = value.Date; }
		}
	}

	class InputArgs
	{
		[Option('f', "file", Required = true, HelpText = "File with birthday data in CSV format.")]
		public string DataFile { get; set; }

		[Option('c', "credentials", Required = true)]
		public string CredFile { get; set; }

		[Option]
		public string Calendar { get; set; }

		[ParserState]
		public IParserState LastParserState { get; set; }

		[HelpOption]
		public string GetUsage()
		{
			return HelpText.AutoBuild(this,
			  (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
		}
	}

	class Program
	{
		static void WriteError(string message, params object[] args)
		{
			Console.WriteLine(message, args);
			Environment.Exit(-1);
		}

		static CalendarService GetService(string credFile)
		{
			UserCredential credential;
			using (FileStream stream = new FileStream(credFile, FileMode.Open, FileAccess.Read))
			{
				credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
					GoogleClientSecrets.Load(stream).Secrets,
					new[] { CalendarService.Scope.Calendar },
					"user", CancellationToken.None,
					new FileDataStore("Calendar.Auth.Store")).Result;
			}

			return new CalendarService(new BaseClientService.Initializer()
			{
				HttpClientInitializer = credential,
				ApplicationName = "GCCalendar"
			});
		}

		static void Main(string[] args)
		{
			var options = new InputArgs();
			if (!CommandLine.Parser.Default.ParseArguments(args, options))
				return;
			
			if (!File.Exists(options.DataFile))
				WriteError("File does not exist.");

			using (TextReader reader = new StreamReader(options.DataFile))
			{
				CsvReader csv = new CsvReader(reader, new CsvConfiguration() { Encoding = Encoding.UTF8 });
				IEnumerable<Birhday> birthdays = csv.GetRecords<Birhday>();

				CalendarService service = GetService(options.CredFile);

				IList<CalendarListEntry> calendars = service.CalendarList.List().Execute().Items;
				CalendarListEntry calendar = String.IsNullOrEmpty(options.Calendar)
					? calendars.FirstOrDefault(c => c.Primary == true)
					: calendars.FirstOrDefault(c => String.Equals(c.Summary, options.Calendar));
				
				if (calendar == null)
					WriteError("Calendar not found.");


				IList<Event> events = service.Events.List(calendar.Id).Execute().Items;
				//Dictionary<string, string> eventDict = events.ToDictionary(e => e.Summary, e => e.Id);
				HashSet<string> eventSet = new HashSet<string>(events.Select(e => e.Summary));

				List<Task<Event>> tasks = new List<Task<Event>>();
				foreach (Birhday b in birthdays)
				{
					string summary = "Birthday: " + b.Name;
					if (eventSet.Contains(summary))
						continue;
					EventDateTime dateTime = new EventDateTime()
					{
						Date = b.Date.ToString("yyyy-MM-dd"),
						TimeZone = calendar.TimeZone
					};
			
					Event bevent = new Event()
					{
						Summary = summary,
						Transparency = "transparent",
						Recurrence = new string[] { "RRULE:FREQ=YEARLY" },
						Start = dateTime,
						End = dateTime,
						Reminders = 
					};
					Task<Event> task = service.Events.Insert(bevent, calendar.Id).ExecuteAsync();
					tasks.Add(task);
					Console.WriteLine("{0} | {1:dd.MM.yyyy}", b.Name, b.Date);
				}
				Task.WhenAll(tasks).Wait();
				Console.WriteLine("Done!");
			}
		}
	}
}

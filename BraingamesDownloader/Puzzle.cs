using System;

namespace BraingamesDownloader
{
	public class Puzzle
	{
		public string Name { get; internal set; }
		public string Category { get; internal set; }
		public int Weight { get; internal set; }
		public DateTime Date { get; internal set; }
		public int Rating { get; internal set; }
		public string SiteLink { get; internal set; }
		public string Description { get; internal set; }
		public string Answer { get; internal set; }
		public string FullAnswer { get; internal set; }
		public DateTime ResolvedDate { get; internal set; }
		public string DescriptionLink { get; internal set; }
	}
}

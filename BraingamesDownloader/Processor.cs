using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using GemBox.Spreadsheet;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace BraingamesDownloader
{
	[TestFixture]
	public class Processor
	{
		private IWebDriver driver;
		private WebDriverWait wait;

		[SetUp]
		public void Start()
		{
			var options = new ChromeOptions();
			options.AddArgument("headless");
			driver = new ChromeDriver(options);

			wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
			//Debug.WriteLine(options);
		}

		[TearDown]
		public void Stop()
		{
			driver.Quit();
			driver = null;
		}

		[Test]
		public void DownloadAllSolvedPuzzles()
		{
			Login();
			Thread.Sleep(500);

			//раздел решенных задач в личном кабинете
			driver.Url = "http://www.braingames.ru/?path=privatedata&action=myanswers&answers=checked";
			var pageCount = driver.FindElements(By.CssSelector("a.rating")).Count + 1;
			var allPuzzles = new List<Puzzle>();

			//pageCount + 1
			//
			//
			//
			//
			for (var i = 1; i < pageCount + 1; i++)
				//
				//
				//
				//
			{
				driver.Url = $"http://www.braingames.ru/?path=privatedata&action=myanswers&answers=checked&page={i}";

				var puzzlesOnPageCount = driver.FindElements(By.CssSelector("a.answerlink")).Count;
				for (int k = 0; k < puzzlesOnPageCount; k++)
				{
					var p = new Puzzle();
					var pButton = driver.FindElements(By.CssSelector("a.answerlink"))[k];
					p.SiteLink = pButton.GetAttribute("href");
					p.ResolvedDate = DateTime.Parse(driver.FindElement(By.CssSelector($"body tbody tr:nth-child({k + 2}) > td:nth-child(5)")).Text.Trim(' ').Trim(Environment.NewLine.ToCharArray()).Replace("n/a", "01.01.1990"));

					pButton.Click();
					//Thread.Sleep(500);

					p.Name = driver.FindElement(By.CssSelector("div:nth-child(1) > .tasktitle")).Text;
					p.Category = driver.FindElement(By.CssSelector(".taskdescription > .taskcategory")).Text;
					p.Weight = Int32.Parse(driver.FindElement(By.CssSelector("span.taskdescription > b")).Text);

					p.Rating = Int32.Parse(driver.FindElement(By.CssSelector("span.taskdescription > span > b")).Text.TrimEnd('%'));

					p.Date = DateTime.Parse(driver.FindElement(By.CssSelector("div:nth-child(2) > span.taskdescription")).Text.Split(',').Last().Trim(' ').Trim(Environment.NewLine.ToCharArray()));

					p.Description = driver.FindElement(By.CssSelector("#puzzleInfo td > div:nth-child(4)")).Text;

					wait = new WebDriverWait(driver, TimeSpan.FromSeconds(1));
					if (driver.FindElements(By.CssSelector("#puzzleInfo > table > tbody > tr > td > div:nth-child(4) > a")).Count != 0)
					{ p.DescriptionLink = driver.FindElement(By.CssSelector("#puzzleInfo > table > tbody > tr > td > div:nth-child(4) > a")).GetAttribute("href"); }
					wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));

					p.Answer = driver.FindElement(By.CssSelector("div.blockText.comments.spoilersupport")).Text;

					allPuzzles.Add(p);

					driver.Url = $"http://www.braingames.ru/?path=privatedata&action=myanswers&answers=checked&page={i}";
					Debug.WriteLine($"{DateTime.UtcNow.ToString("HH:mm:ss.fff")} - страница {i} - задача {k + 1} - {p.Name}");
				}
			}

			SpreadsheetInfo.SetLicense("EIKU-U5LX-6MSF-Z84S");
			ExcelFile ef = new ExcelFile();
			ExcelWorksheet ws = ef.Worksheets.Add("решенные задачи");

			DataTable dt = new DataTable();

			// add columns
			var c1 = dt.Columns.Add("Название", typeof(string));
			var c2 = dt.Columns.Add("Категория", typeof(string));
			var c3 = dt.Columns.Add("Вес", typeof(int));
			var c4 = dt.Columns.Add("Дата", typeof(DateTime));
			var c5 = dt.Columns.Add("Зачтена", typeof(DateTime));
			var c6 = dt.Columns.Add("Симпатии", typeof(int));
			var c7 = dt.Columns.Add("Условие", typeof(string));
			var c8 = dt.Columns.Add("Решение", typeof(string));
			var c9 = dt.Columns.Add("Ссылка из условия", typeof(string));
			var c10 = dt.Columns.Add("Ссылка на сайт", typeof(string));
			var c11 = dt.Columns.Add("Полное решение", typeof(string));

			foreach (var p in allPuzzles)
			{
				dt.Rows.Add(
						p.Name,
						p.Category,
						p.Weight,
						p.Date,
						p.ResolvedDate,
						p.Rating,
						p.Description,
						p.Answer,
						p.DescriptionLink,
						p.SiteLink,
						p.FullAnswer);
			}

			// add cell
			// ws.Cells[0, 0].Value = "DataTable insert example:";

			// Insert DataTable into an Excel worksheet.
			ws.InsertDataTable(dt,
				new InsertDataTableOptions()
				{
					ColumnHeaders = true,
					StartRow = 0
				});

			//AutoFitAllColumns(ws);

			var date = DateTime.Now.ToString("yyyy.MM.dd_HH-mm-ss");
			var fileName = String.Concat(@"D:/braingames", date, ".xlsx");
			ef.Save(fileName);
		}

		private static void AutoFitAllColumns(ExcelWorksheet ws)
		{
			var columnCount = ws.CalculateMaxUsedColumns();
			for (int i = 0; i < columnCount; i++)
				ws.Columns[i].AutoFit();
		}

		private void Login()
		{
			var login = ConfigurationManager.AppSettings["login"];
			var pass = ConfigurationManager.AppSettings["password"];

			if (login == "" || pass == "") { throw new System.ArgumentException("Please provide correct login data"); }

			driver.Url = "http://www.braingames.ru/";
			driver.FindElement(By.CssSelector("#login")).SendKeys(login);
			driver.FindElement(By.CssSelector("#password")).SendKeys(pass);
			driver.FindElement(By.CssSelector("[name=btnLogin]")).Click();
		}

	}
}
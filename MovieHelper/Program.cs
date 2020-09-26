using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using FilmWebAPI;
using OfficeOpenXml;

namespace MovieHelper
{
    class Program
    {
        private static FilmWeb _filmWeb;

        async static Task Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            _filmWeb = new FilmWeb();
            var exampleFileLocation = @"m:\Documents\ListaFilmówDoObejrzenia.xlsx";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(exampleFileLocation));
            var firstSheet = package.Workbook.Worksheets[0];
            var table = firstSheet.Tables[0];
            var cells = table.WorkSheet.Cells;
            var addr = table.Address;
            var firstRow = addr.Start.Row;
            var firstCol = addr.Start.Column;
            var lastRow = addr.End.Row;
            for (int currentRow = firstRow + 1; currentRow <= lastRow; currentRow++)
            {
                await processMovie(cells, currentRow, firstCol);
            }

            package.Save();
        }

        private static async Task processMovie(ExcelRange cells, int currentRow, int firstCol)
        {
            var movieName = cells[currentRow, firstCol].Value.ToString();
            var movieId = await _filmWeb.GetMovieId(movieName);

            if (!movieId.HasValue) return;
            var movieAverageVote =
                Convert.ToDouble(
                    (await _filmWeb.GetFilmAvgVote(movieId.Value)).ToString(CultureInfo.CurrentCulture));

            Console.WriteLine($"Movie: {movieName} : {movieAverageVote.ToString(CultureInfo.CurrentCulture)}");
            cells[currentRow, firstCol + 4].Value = movieAverageVote;
            cells[currentRow, firstCol + 4].Style.Numberformat.Format = "0.00";
        }
    }
}
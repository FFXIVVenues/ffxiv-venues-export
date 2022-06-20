using FFXIVVenues.VenueModels.V2022;
using OfficeOpenXml;
using System.Diagnostics;
using System.Text.Json;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var request = await new HttpClient().GetAsync("https://raw.githubusercontent.com/FFXIVVenues/ffxiv-venues-web/master/src/venues.json");
var venues = await JsonSerializer.DeserializeAsync<FFXIVVenues.VenueModels.V2021.Venue[]>(await request.Content.ReadAsStreamAsync());

if (venues == null)
{
    Console.ForegroundColor = ConsoleColor.Red;
    Console.Error.WriteLine("Error: Could not deserialize venues.");
    return; 
}

using (var output = File.Create("venues.xlsx"))
using (var package = new ExcelPackage(output))
{
    var worksheet = package.Workbook.Worksheets.Add("Venues");

    worksheet.Row(1).Style.Font.Bold = true;
    worksheet.Column(2).Width = 15;
    worksheet.Cells[1, 1].Value = "Id";
    worksheet.Column(2).Width = 40;
    worksheet.Column(2).Style.Font.Bold = true;
    worksheet.Cells[1, 2].Value = "Name";
    worksheet.Column(3).Width = 40;
    worksheet.Cells[1, 3].Value = "Contacts";
    worksheet.Cells[1, 4].Value = "Data Center";
    worksheet.Column(5).Width = 15;
    worksheet.Cells[1, 5].Value = "World";
    worksheet.Column(6).Width = 15;
    worksheet.Cells[1, 6].Value = "District";
    worksheet.Cells[1, 7].Value = "Ward";
    worksheet.Cells[1, 8].Value = "Plot";
    worksheet.Cells[1, 9].Value = "Apartment";
    worksheet.Cells[1, 10].Value = "Is Subdivision";
    worksheet.Column(11).Width = 40;
    worksheet.Cells[1, 11].Value = "Website";
    worksheet.Column(12).Width = 40;
    worksheet.Cells[1, 12].Value = "Discord";
    worksheet.Column(13).Width = 80;
    worksheet.Cells[1, 13].Value = "Tags";

    var col = 14;
    for (var d = 0; d < 7; d++)
    {
        var dayName = d switch
        {
            0 => "Mon",
            1 => "Tue",
            2 => "Wed",
            3 => "Thu",
            4 => "Fri",
            5 => "Sat",
            6 => "Sun"
        };
        for (var h = 0; h < 24; h++)
        {
            worksheet.Column(col).Width = 10;
            worksheet.Cells[1, col].Value = $"{dayName} {h.ToString("D2")}:00";
            col++;
        }
    }

    for (var i = 0; i < venues.Length; i++)
    {
        var venue = new FFXIVVenues.VenueModels.V2022.Venue(venues[i]);
        var row = i + 2;
        Console.WriteLine(venue.Name);
        worksheet.Cells[row, 1].Value = venue.Id;
        worksheet.Cells[row, 2].Value = venue.Name;
        worksheet.Cells[row, 3].Value = venue.Contacts != null ? string.Join(", ", venue.Contacts) : "";
        worksheet.Cells[row, 4].Value = venue.Location.DataCenter;
        worksheet.Cells[row, 5].Value = venue.Location.World;
        worksheet.Cells[row, 6].Value = venue.Location.District;
        worksheet.Cells[row, 7].Value = venue.Location.Ward;
        worksheet.Cells[row, 8].Value = venue.Location.Plot == 0 ? "" : venue.Location.Plot;
        worksheet.Cells[row, 9].Value = venue.Location.Apartment == 0 ? "" : venue.Location.Apartment;
        worksheet.Cells[row, 10].Value = venue.Location.Subdivision;

        worksheet.Cells[row, 11].Style.Font.UnderLine = true;
        worksheet.Cells[row, 11].Hyperlink = venue.Website;
        worksheet.Cells[row, 11].Value = venue.Website;

        worksheet.Cells[row, 12].Style.Font.UnderLine = true;
        worksheet.Cells[row, 12].Hyperlink = venue.Discord;
        worksheet.Cells[row, 12].Value = venue.Discord;

        worksheet.Cells[row, 13].Value = venue.Tags != null ? string.Join(", ", venue.Tags) : "";

        var hourGrid = new float[][] { new float[24], new float[24], new float[24], new float[24], new float[24], new float[24], new float[24] };
        foreach (var opening in venue.Openings)
        {
            var day = opening.Day;
            if (opening.Start.NextDay) day++;
            if ((int)day == 7) day = 0;

            var end = opening.End != null ?
                opening.End :
                new Time
                {
                    Hour = (ushort)((opening.Start.Hour + 2) % 24),
                    Minute = opening.Start.Minute,
                    NextDay = opening.Start.NextDay || opening.Start.Hour + 2 > 23,
                    TimeZone = opening.Start.TimeZone
                };

            var curHour = opening.Start.Hour;
            hourGrid[(int)day][curHour++] = (float)1 - (opening.Start.Minute / 60);
            if (!opening.Start.NextDay && end.NextDay)
            {
                for (; curHour < 24; curHour++)
                    hourGrid[(int)day][curHour] = 1;
                curHour = 0;
                day++;
                if ((int)day == 7) day = 0;
            }

            for (; curHour < end.Hour; curHour++)
                hourGrid[(int)day][curHour] = 1;

            hourGrid[(int)day][curHour++] = (float)end.Minute / 60;
        }

        for (var d = 0; d < 7; d++)
            for (var h = 0; h < 24; h++)
            {
                var cell = worksheet.Cells[row, 14 + ((d * 24) + h)];
                cell.Style.Numberformat.Format = "0.0";
                cell.Value = hourGrid[d][h];
            }
    }

    package.Save();
}

using Process fileopener = new Process();
fileopener.StartInfo.FileName = "explorer";
fileopener.StartInfo.Arguments = $"\"{new FileInfo("venues.xlsx").FullName}\"";
fileopener.Start();

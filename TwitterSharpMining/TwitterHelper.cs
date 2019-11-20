using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Tweetinvi;
using Tweetinvi.Models.DTO.QueryDTO;

namespace TwitterSharpMining
{
    public static class TwitterHelper
    {
        /*
         * Uncomment and set with the values you get for your Tweeter application
        */
        static readonly string _accessToken = "";  Set with values generated for Twitter Api
        static readonly string _accessTokenSecret = ""; Set with values generated for Twitter Api
        static readonly string _consumerKey = ""; Set with values generated for Twitter Api
        static readonly string _consumerSecret = ""; Set with values generated for Twitter Api

        static TwitterHelper()
        {
            Auth.SetUserCredentials(_consumerKey, _consumerSecret, _accessToken, _accessTokenSecret);
            RateLimit.RateLimitTrackerMode = RateLimitTrackerMode.TrackOnly;
        }

        public static void DumpFollowers(string screenName, DateTime createdSince, long nextCursor = 0, long totalCounted = 0, long newUsers = 0, int cyclesBeforeDump = 4)
        {
            /*
             * createdSince filters the users that were created before the specified date
             * You can use nextCursors totalCounted newUsers if the app closed and you have dumped intermediary files and you have the last cursor
             * cyclesBEforeDump says after how many cycles to dump an intermediary excelfile where you will fint also the last cursor Default is 4
             * and every cycle lasts 15 minutes for Twitter 15 quieries limitation So every cycle = 15*200= 3000 followers parsed. So with a cycleBeforeDump=4 you will parse 12.000 followers in around hour
             */

            var query = string.Format("https://api.twitter.com/1.1/followers/list.json?screen_name={0}&count=200", screenName);

            XSSFWorkbook workBook;
            ISheet followersSheet;
            IRow sheetRow;
            FileStream excelFile;
            ICellStyle dateStyle;

            CreateExcelFile(screenName,out excelFile, out workBook, out followersSheet, out dateStyle);
            int rowNumber = 1; //we already created the first row with the headers
            int fileDumpCount = 0;
            try
            {
                do
                {
                    var queryRateLimits = RateLimit.GetQueryRateLimit(query);

                    if (queryRateLimits != null && queryRateLimits.Remaining == 0)
                    {
                        //  RateLimit.AwaitForQueryRateLimit(queryRateLimits);
                        //Sometimes is does not work so you will need to wait a little more I am adding 20 seconds
                        var waitSeconds = TimeSpan.FromSeconds(queryRateLimits.ResetDateTimeInSeconds).Add(TimeSpan.FromSeconds(20));
                        Console.WriteLine($"[{DateTime.Now.ToString("HH:mm")}] Waiting until {DateTime.Now.AddSeconds(waitSeconds.TotalSeconds).ToString("HH:mm:ss")} for next request cycle ...");
                        Thread.Sleep(waitSeconds);
                    }

                    var followers = GetFollowers(query, nextCursor, out nextCursor);
                    fileDumpCount++;
                    foreach (var results in followers)
                    {
                        foreach (var user in results.Users)
                        {
                            totalCounted++;
                            if (user.CreatedAt > createdSince && !user.Verified)
                            {
                                sheetRow = followersSheet.CreateRow(rowNumber++);
                                sheetRow.CreateCell(0).SetCellValue(++newUsers);
                                sheetRow.CreateCell(1).SetCellValue(user.IdStr);
                                sheetRow.CreateCell(2).SetCellValue(user.Name);
                                sheetRow.CreateCell(3).SetCellValue(user.ScreenName);
                                sheetRow.CreateCell(4).SetCellValue(user.FollowersCount);
                                sheetRow.CreateCell(5).SetCellValue(user.FriendsCount);
                                sheetRow.CreateCell(6).SetCellValue(user.ListedCount.HasValue ? user.ListedCount.Value : 0);

                                var cell6 = sheetRow.CreateCell(7);
                                cell6.SetCellValue(user.CreatedAt);
                                cell6.CellStyle = dateStyle;

                                sheetRow.CreateCell(8).SetCellValue(user.StatusesCount);
                            }
                        }

                    }
                    Console.WriteLine($"[{DateTime.Now.ToString("HH:mm")}] Parsed {totalCounted.ToString("D9")} followers. Created since {createdSince.ToString("dd.MM.yyyy")}: {newUsers.ToString("D9")} ...");
                    if (fileDumpCount == cyclesBeforeDump)
                    {
                        CloseExcelFile(excelFile, workBook, nextCursor, totalCounted, newUsers);
                        CreateExcelFile(screenName,out excelFile, out workBook, out followersSheet, out dateStyle);
                        rowNumber = 1;
                        fileDumpCount = 0;
                    }

                }
                while (nextCursor != -1 && nextCursor != 0);
            }
            catch { }
            CloseExcelFile(excelFile, workBook, nextCursor, totalCounted, newUsers);
        }

        static void CreateExcelFile(string screenName,out FileStream fs, out XSSFWorkbook workBook, out ISheet followersSheet, out ICellStyle _dateStyle)
        {
            fs = new FileStream($"{screenName}_{DateTime.Now.ToString("dd_HH_mm")}.xlsx", FileMode.Create, FileAccess.Write);
            Console.WriteLine($"[{DateTime.Now.ToString("HH:mm")}] New file {fs.Name} ...");

            workBook = new XSSFWorkbook();

            var creationHelper = workBook.GetCreationHelper();
            _dateStyle = workBook.CreateCellStyle();
            _dateStyle.DataFormat = creationHelper.CreateDataFormat().GetFormat("dd mmm yy hh:mm");
            _dateStyle.Alignment = HorizontalAlignment.Left;

            followersSheet = workBook.CreateSheet("Followers");
            var row = followersSheet.CreateRow(0);

            row.CreateCell(0).SetCellValue("Count");
            row.CreateCell(1).SetCellValue("Id");

            row.CreateCell(2).SetCellValue("Name");
            row.CreateCell(3).SetCellValue("Screen Name");
            row.CreateCell(4).SetCellValue("Followers Count");
            row.CreateCell(5).SetCellValue("Friends Count");
            row.CreateCell(6).SetCellValue("Listed");
            row.CreateCell(7).SetCellValue("Created At");
            row.CreateCell(8).SetCellValue("Status Count");

        }

        static void CloseExcelFile(FileStream fs, XSSFWorkbook workBook, long nextCursor, long totalCounted, long newUsers)
        {
            ISheet summarySheet = workBook.CreateSheet("Summary");
            IRow row = summarySheet.CreateRow(0);

            row.CreateCell(0).SetCellValue("TotalCounted");
            row.CreateCell(1).SetCellValue("NewAccounts");
            row.CreateCell(2).SetCellValue("Cursor");

            row = summarySheet.CreateRow(1);
            row.CreateCell(0).SetCellValue(totalCounted);
            row.CreateCell(1).SetCellValue(newUsers);
            row.CreateCell(2).SetCellValue(nextCursor);

            workBook.Write(fs);
            fs.Close();
            Console.WriteLine($"[{DateTime.Now.ToString("HH:mm")}] Flushed  {fs.Name} to disk ...");
            fs.Dispose();
          
        }

        private static IEnumerable<IUserCursorQueryResultDTO> GetFollowers(string query, long cursor, out long nextCursor)
        {
            var results = TwitterAccessor.ExecuteCursorGETCursorQueryResult<IUserCursorQueryResultDTO>(query, cursor: cursor).ToArray();

            if (!results.Any())
            {
                // Something went wrong. The RateLimits operation tokens got used before we performed our query
                RateLimit.ClearRateLimitCache();
                RateLimit.AwaitForQueryRateLimit(query);
                results = TwitterAccessor.ExecuteCursorGETCursorQueryResult<IUserCursorQueryResultDTO>(query, cursor: cursor).ToArray();
            }

            if (results.Any())
                nextCursor = results.Last().NextCursor;
            else
                nextCursor = -1;
            return results;
        }

    }
}

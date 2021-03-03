using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;

namespace BhpAssetComplianceWpfOneDesktop.Utility
{
    public static class ExportImageToCsv
    {
        public static string ConvertImageToString(string imageFilePath)
        {
            var image = Image.FromFile(imageFilePath);
            var imageConverter = new ImageConverter();
            var imageByte = (byte[])imageConverter.ConvertTo(image, typeof(byte[]));
            return Convert.ToBase64String(imageByte ?? Array.Empty<byte>());
        }

        public static void AppendImageToCSV(string imageFilePath, string place, DateTime date, string targetFilePath)
        {
            var imageString = ConvertImageToString(imageFilePath);
            var content = $"{date},{place},{imageString}";
            File.AppendAllLines(targetFilePath, new List<string> { content });
        }

        public static void AppendImageDepressurizationToCSV(string imageFilePath, DateTime date, string targetFilePath)
        {
            var imageString = ConvertImageToString(imageFilePath);
            var content = $"{date},{imageString}";
            File.AppendAllLines(targetFilePath, new List<string> { content });
        }

        public static void AppendImageMineSequenceToCSV(string imageFilePath1, string imageFilePath2, DateTime date, string targetFilePath)
        {
            var imageString1 = ConvertImageToString(imageFilePath1);
            var imageString2 = ConvertImageToString(imageFilePath2);
            var content = $"{date},{imageString1},{imageString2}";
            File.AppendAllLines(targetFilePath, new List<string> { content });
        }

        public static int SearchByDate(string place, DateTime date, string targetFilePath)
        {
            var strLines = File.ReadLines(targetFilePath);
            var count = 0;
            foreach (var line in strLines)
            {

                var splits = line.Split(',');
                if (splits[0] == "Date")
                {
                    count++;
                }
                else
                {
                    var lineDate = DateTime.Parse(splits[0]).Date;
                    var linePlace = splits[1];
                    if (lineDate == date.Date && linePlace == place)
                    {
                        return count;
                    }
                    count++;

                }

            }
            return -1;
        }

        public static int SearchByDateMineSequence(DateTime date, string targetFilePath)
        {
            var strLines = File.ReadLines(targetFilePath);
            var count = 0;
            foreach (var line in strLines)
            {

                var splits = line.Split(',');
                if (splits[0] == "Date")
                    count++;
                else
                {
                    var lineDate = DateTime.Parse(splits[0]).Date;
                    if (lineDate == date.Date)
                        return count;
                    count++;
                }

            }
            return -1;
        }

        public static void RemoveItem(string targetFilePath, int index)
        {
            var lines = File.ReadLines(targetFilePath).ToList();
            lines.RemoveAt(index);
            File.WriteAllLines(targetFilePath, lines);
        }

        public static string ToNullSafeString(this object obj)
        {
            return (obj ?? string.Empty).ToString();
        }
    }
}

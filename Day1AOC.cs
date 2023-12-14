namespace AdventOfCode2023
{
    using System;
    using System.Data;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Reflection.PortableExecutable;
    using System.Text;
    using System.Text.RegularExpressions;
    using ExcelDataReader;

    public class Day1AOC
    {
        public static int GetSumFirstLast(string input)
        {
            return (GetFirstInt(input) * 10) + GetLastInt(input);
        }

        private static void Main()
        {
             var excelFilePath = @"C:\Users\alwib\source\repos\AdventOfCode2023\AdventOfCode2023\Day1Input.xlsx";
             int totalSum = 0;

             using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
             {
                var dataTable = ReadDataFromExcel(stream);

                foreach (DataRow dataRow in dataTable.Rows)
                {
                    var inputString = dataRow[0]?.ToString();
                    inputString = TurnWordsIntoNumbers(inputString);

                    if (!string.IsNullOrEmpty(inputString))
                    {
                        var sum = GetSumFirstLast(inputString);
                        Console.WriteLine(sum);
                        totalSum += sum;
                    }
                }
             }

             Console.WriteLine($"Total Sum: {totalSum}");
        }

        private static DataTable ReadDataFromExcel(Stream stream)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
            ExcelDataSetConfiguration conf = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = false,
                },
            };

            var dataSet = reader.AsDataSet(conf);
            return dataSet.Tables[0];
        }

        private static int GetFirstInt(string input)
        {
            var result = 0;
            for (var i = 0; i < input.Length; i++)
            {
                if (Char.IsDigit(input[i]))
                {
                    return input[i] - '0';
                }
            }

            return result;
        }

        private static int GetLastInt(string input)
        {
            var result = 0;
            string reversedInput = Reverse(input);
            for (var i = 0; i < reversedInput.Length; i++)
            {
                if (Char.IsDigit(reversedInput[i]))
                {
                    return reversedInput[i] - '0';
                }
            }

            return result;
        }

        private static string Reverse(string s)
        {
            char[] charArray = s.ToCharArray();
            Array.Reverse(charArray);
            return new string(charArray);
        }

        private static string TurnWordsIntoNumbers(string input)
        {
            string pattern = @"(one|two|three|four|five|six|seven|eight|nine)";

            MatchCollection matches = Regex.Matches(input, pattern);

            Dictionary<string, string> wordToNumber = new Dictionary<string, string>();

            while (matches.Count > 0)
            {
                foreach (Match match in matches)
                {
                    string word = match.Value;
                    string number = WordToNumberString(word);
                    wordToNumber[word] = number;
                }

                foreach (var entry in wordToNumber)
                {
                    input = input.Replace(entry.Key, entry.Value.ToString());
                }

                matches = Regex.Matches(input, pattern);
            }

            return input;

        }

        private static string WordToNumberString(string word)
        {
            switch (word)
            {
                case "one": return "o1e";
                case "two": return "t2o";
                case "three": return "t3e";
                case "four": return "f4r";
                case "five": return "f5e";
                case "six": return "s6x";
                case "seven": return "s7n";
                case "eight": return "e8t";
                case "nine": return "n9e";
                default: return "eror";
            }
        }
    }
}
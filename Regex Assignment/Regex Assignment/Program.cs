// First Challenge

using System;
using System.Text.RegularExpressions;

public class Program
{
    public static void Main()
    {
        while (true)
        {
            Console.WriteLine("Enter a date (mm/dd/yyyy) or type 'exit' to quit:");
            string input = Console.ReadLine();

            if (input.Equals("exit", StringComparison.CurrentCultureIgnoreCase))
            {
                break;
            }

            try
            {
                string result = ReverseDateFormat(input);
                Console.WriteLine($"Converted date: {result}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public static string ReverseDateFormat(string dateInput)
    {
        if (string.IsNullOrWhiteSpace(dateInput))
        {
            throw new ArgumentException("Input date cannot be null or empty.");
        }

        string pattern = @"^(?<mon>\d{1,2})/(?<day>\d{1,2})/(?<year>\d{2,4})$";
        TimeSpan timeout = TimeSpan.FromMilliseconds(500);

        try
        {
            Match match = Regex.Match(dateInput, pattern, RegexOptions.None, timeout);

            if (match.Success)
            {
                string month = match.Groups["mon"].Value;
                string day = match.Groups["day"].Value;
                string year = match.Groups["year"].Value;

                if (year.Length == 2)
                {
                    year = "20" + year;
                }

                string normalizedDate = $"{month}/{day}/{year}";
                if (DateTime.TryParse(normalizedDate, out DateTime parsedDate))
                {
                    return parsedDate.ToString("yyyy-MM-dd");
                }
                else
                {
                    throw new ArgumentException("Invalid date format.");
                }
            }
            else
            {
                throw new ArgumentException("Invalid date format.");
            }
        }
        catch (RegexMatchTimeoutException)
        {
            throw new TimeoutException("The regex operation timed out.");
        }
    }
}




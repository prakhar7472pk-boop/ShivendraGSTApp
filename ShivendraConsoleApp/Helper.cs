using System;
using System.Linq;

namespace ShivendraConsoleApp;

internal class Helper
{
    private const string del = " - ";

    internal static string? GetFieldValue(string value, params string[] options)
    {
        string? title = value.Split('(', '-', ')').FirstOrDefault(s => !s.Equals(string.Empty))?.Trim();
        if (title is null) return null;

        foreach (var option in options)
        {
            if (string.Equals(title, option, StringComparison.OrdinalIgnoreCase))
            {
                string sub = value.Substring(option.Length);
                int i = 0;
                while (i < sub.Length && (char.IsWhiteSpace(sub[i]) || sub[i].Equals('-'))) i++;

                sub = sub.Substring(i);

                return option + del + (sub.StartsWith(option, StringComparison.OrdinalIgnoreCase) ? sub.Substring(option.Length) : sub).Trim();
            }
        }

        return null;
    }
}
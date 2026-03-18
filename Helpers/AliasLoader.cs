using System.IO;
using MSSLTeamsMessenger.Models;

namespace MSSLTeamsMessenger.Helpers;

public static class AliasLoader
{
    public static List<AliasEntry> LoadFromFile(string filePath)
    {
        if (!File.Exists(filePath))
            return new List<AliasEntry>();

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var results = new List<AliasEntry>();

        foreach (var rawLine in File.ReadLines(filePath))
        {
            var line = rawLine.Trim();
            if (string.IsNullOrWhiteSpace(line))
                continue;

            // Strip @microsoft.com suffix if present
            var alias = line;
            var atIndex = alias.IndexOf('@');
            if (atIndex > 0)
                alias = alias.Substring(0, atIndex);

            alias = alias.Trim();
            if (string.IsNullOrWhiteSpace(alias))
                continue;

            if (!seen.Add(alias.ToLowerInvariant()))
                continue; // duplicate

            results.Add(new AliasEntry { Alias = alias });
        }

        return results;
    }
}

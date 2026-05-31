using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace readbom;

internal static class SwDocumentManagerLicense
{
    private const string EnvironmentVariableName = "READBOM_SWDM_LICENSE";

    public static string Load()
    {
        return LoadCandidates().FirstOrDefault()
               ?? throw new InvalidOperationException(
                   "未找到 SolidWorks Document Manager 许可证。请设置环境变量 READBOM_SWDM_LICENSE，" +
                   "或在程序目录放置 sw-document-manager-license.txt。");
    }

    public static IReadOnlyList<string> LoadCandidates()
    {
        var candidates = new List<string>();
        var fromEnvironment = Environment.GetEnvironmentVariable(EnvironmentVariableName);
        AddCandidate(candidates, fromEnvironment);

        foreach (var path in GetPlainTextLicenseCandidates())
        {
            candidates.AddRange(TryReadPlainTextLicenses(path).Where(candidate => !ContainsCandidate(candidates, candidate)));
        }

        foreach (var path in GetSwVersionCandidates())
        {
            var value = TryReadSwVersionLicense(path);
            AddCandidate(candidates, value);
        }

        return candidates;
    }

    private static IEnumerable<string> GetPlainTextLicenseCandidates()
    {
        yield return Path.Combine(AppContext.BaseDirectory, "sw-document-manager-license.txt");
        yield return Path.Combine(AppContext.BaseDirectory, "SwDocumentManagerLicenseKey.txt");
        yield return Path.Combine(GetSolution1TreehousePath(), "SwDocumentManagerLicenseKey.txt");
        yield return @"D:\2022\Treehouse\SwDocumentManagerLicenseKey.txt";
    }

    private static IEnumerable<string> GetSwVersionCandidates()
    {
        yield return Path.Combine(GetSolution1TreehousePath(), "SwVersion.cs");
        yield return @"D:\2022\Treehouse\SwVersion.cs";
    }

    private static string GetSolution1TreehousePath()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "RiderProjects",
            "Solution1",
            "Treehouse");
    }

    private static List<string> TryReadPlainTextLicenses(string path)
    {
        var result = new List<string>();
        try
        {
            if (!File.Exists(path))
            {
                return result;
            }

            var text = File.ReadAllText(path).Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                return result;
            }

            AddCandidate(result, text);
            var entries = text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            foreach (var entry in entries)
            {
                var parts = entry.Split(new[] { '-' }, 2, StringSplitOptions.None);
                var value = parts.Length == 2 ? parts[1] : parts[0];
                AddCandidate(result, value);
            }

            return result;
        }
        catch
        {
            return result;
        }
    }

    private static string? TryReadSwVersionLicense(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return null;
            }

            var content = File.ReadAllText(path);
            var match = Regex.Match(content, @"new\s+byte\[\d+\]\s*\{([^}]+)\}", RegexOptions.Singleline);
            if (!match.Success)
            {
                return null;
            }

            var numbers = Regex.Matches(match.Groups[1].Value, @"\d+");
            if (numbers.Count == 0)
            {
                return null;
            }

            var bytes = new byte[numbers.Count];
            for (var i = 0; i < numbers.Count; i++)
            {
                bytes[i] = Convert.ToByte(numbers[i].Value, CultureInfo.InvariantCulture);
            }

            var chars = new List<char>();
            for (var i = 0; i < bytes.Length; i++)
            {
                bytes[i] = (byte)(bytes[i] ^ (i + 1));
                if ((i & 1) == 1)
                {
                    chars.Add((char)bytes[i - 1]);
                }
            }

            return new string(chars.ToArray()).Trim();
        }
        catch
        {
            return null;
        }
    }

    private static void AddCandidate(ICollection<string> candidates, string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        var candidate = value.Trim();
        if (!ContainsCandidate(candidates, candidate))
        {
            candidates.Add(candidate);
        }
    }

    private static bool ContainsCandidate(IEnumerable<string> candidates, string value)
    {
        return candidates.Any(candidate => string.Equals(candidate, value, StringComparison.Ordinal));
    }
}

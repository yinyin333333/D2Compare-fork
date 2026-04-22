namespace D2Compare.Core.Services;

public readonly record struct RowInstanceKey(string Name, int Occurrence);

public static class RowInstanceKeyHelper
{
    public static List<RowInstanceKey> BuildKeys(IReadOnlyList<string> rowNames)
    {
        var occurrenceByName = new Dictionary<string, int>(StringComparer.Ordinal);
        var keys = new List<RowInstanceKey>(rowNames.Count);

        foreach (var rowName in rowNames)
        {
            occurrenceByName.TryGetValue(rowName, out var occurrence);
            occurrence++;
            occurrenceByName[rowName] = occurrence;

            keys.Add(new RowInstanceKey(rowName, occurrence));
        }

        return keys;
    }
}

using System.Text.RegularExpressions;

using D2Compare.Core.Models;

namespace D2Compare.Core.Services;

public static class DiffEngine
{
    public static Dictionary<T, int> GetRemovedRows<T>(
        IReadOnlyList<T> file1Rows,
        IReadOnlyList<T> file2Rows)
        where T : notnull
    {
        var removedRows = new Dictionary<T, int>();

        foreach (var row in file1Rows)
        {
            if (removedRows.ContainsKey(row))
                removedRows[row]++;
            else
                removedRows[row] = 1;
        }

        foreach (var row in file2Rows)
        {
            if (removedRows.ContainsKey(row))
            {
                removedRows[row]--;
                if (removedRows[row] == 0)
                    removedRows.Remove(row);
            }
        }

        return removedRows;
    }

    public static List<T> ExpandCounts<T>(Dictionary<T, int> dictionary)
        where T : notnull
    {
        var list = new List<T>();
        foreach (var kvp in dictionary)
        {
            for (int i = 0; i < kvp.Value; i++)
                list.Add(kvp.Key);
        }
        return list;
    }

    public static List<DiffGroup> GetGroupedDifferences(
        Dictionary<string, List<string>> file1Data,
        Dictionary<string, List<string>> file2Data,
        HashSet<string> allHeaders,
        string rowHeaderColumn,
        bool includeNewRows)
    {
        var groupedDifferences = new Dictionary<string, List<(string Diff, string ColIndex)>>();
        var newRowKeys = new HashSet<string>();

        // Build stable header ordering and index lookup once
        var headerList = allHeaders.ToList();
        var headerIndexMap = new Dictionary<string, int>(headerList.Count);
        for (int i = 0; i < headerList.Count; i++)
            headerIndexMap[headerList[i]] = i;

        // Build row index lookups to avoid O(n) IndexOf per column
        var file1Rows = file1Data[rowHeaderColumn];
        var file1RowKeys = RowInstanceKeyHelper.BuildKeys(file1Rows);
        var file1RowIndices = new Dictionary<RowInstanceKey, int>();
        for (int i = 0; i < file1RowKeys.Count; i++)
            file1RowIndices[file1RowKeys[i]] = i;

        var file2Rows = file2Data[rowHeaderColumn];
        var file2RowKeys = RowInstanceKeyHelper.BuildKeys(file2Rows);
        var file2RowIndices = new Dictionary<RowInstanceKey, int>();
        for (int i = 0; i < file2RowKeys.Count; i++)
            file2RowIndices[file2RowKeys[i]] = i;

        var allRowHeaders = new HashSet<RowInstanceKey>(file1RowKeys);
        allRowHeaders.UnionWith(file2RowKeys);

        Parallel.ForEach(allRowHeaders, rowKey =>
        {
            bool inFile1 = file1RowIndices.ContainsKey(rowKey);
            bool inFile2 = file2RowIndices.ContainsKey(rowKey);

            if (inFile1 && inFile2)
            {
                int index1 = file1RowIndices[rowKey];
                int index2 = file2RowIndices[rowKey];

                foreach (var header in headerList)
                {
                    if (!file1Data.ContainsKey(header) || !file2Data.ContainsKey(header))
                        continue;

                    var value1 = file1Data[header][index1];
                    var value2 = file2Data[header][index2];

                    if (value1 != value2)
                    {
                        string valueDifference = $"{header}: '{value1}' -> '{value2}'";
                        string column0Value = $"(Row {Math.Min(index1, index2) + 1}) {rowKey.Name}";
                        int columnIndex = headerIndexMap.TryGetValue(header, out var idx) ? idx : 0;

                        lock (groupedDifferences)
                        {
                            if (!groupedDifferences.ContainsKey(column0Value))
                                groupedDifferences[column0Value] = new List<(string, string)>();

                            groupedDifferences[column0Value].Add((valueDifference, columnIndex.ToString()));
                        }
                    }
                }
            }
            else if (includeNewRows && !inFile1)
            {
                if (!file2RowIndices.TryGetValue(rowKey, out int index2))
                    return;

                foreach (var header in headerList)
                {
                    if (!file2Data.ContainsKey(header))
                        continue;

                    var value2 = file2Data[header][index2];

                    string valueDifference = $"{header}: '{value2}'";
                    string column0Value = $"(Row {index2 + 1}) {rowKey.Name}";
                    int columnIndex = headerIndexMap.TryGetValue(header, out var idx) ? idx : 0;

                    lock (groupedDifferences)
                    {
                        if (!groupedDifferences.ContainsKey(column0Value))
                            groupedDifferences[column0Value] = new List<(string, string)>();

                        groupedDifferences[column0Value].Add((valueDifference, columnIndex.ToString()));
                        newRowKeys.Add(column0Value);
                    }
                }
            }
        });

        return groupedDifferences
            .OrderBy(pair => int.Parse(Regex.Match(pair.Key, @"\(Row (\d+)\)").Groups[1].Value))
            .Select(pair =>
            {
                var sorted = pair.Value.OrderBy(t => int.Parse(t.ColIndex)).ToList();
                return new DiffGroup(pair.Key, sorted.Select(t => t.Diff).ToList(), newRowKeys.Contains(pair.Key));
            })
            .ToList();
    }
}

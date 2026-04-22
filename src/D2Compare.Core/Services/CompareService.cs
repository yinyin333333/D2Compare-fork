using D2Compare.Core.Models;

namespace D2Compare.Core.Services;

public static class CompareService
{
    public static CompareResult CompareFile(string sourcePath, string targetPath, bool includeNewRows)
    {
        var sourceData = CsvParser.Parse(sourcePath);
        var targetData = CsvParser.Parse(targetPath);

        var allHeaders = new HashSet<string>(sourceData.Keys);
        allHeaders.UnionWith(targetData.Keys);

        var rowHeaderColumn = allHeaders.FirstOrDefault(h => sourceData.ContainsKey(h) && targetData.ContainsKey(h));
        if (rowHeaderColumn is null)
        {
            return new CompareResult(
                Path.GetFileName(sourcePath), [], [], [], [], [], [],
                []);
        }

        // Column diffs
        var sourceKeyList = sourceData.Keys.ToList();
        var targetKeyList = targetData.Keys.ToList();
        var addedColumns = targetData.Keys.Except(sourceData.Keys).ToList();
        var removedColumns = sourceData.Keys.Except(targetData.Keys).ToList();

        // Identify renames via manual fixes
        var changedColumns = new List<string>();
        var remainingAdded = new List<string>(addedColumns);
        var remainingRemoved = new List<string>(removedColumns);

        foreach (var added in addedColumns.ToList())
        {
            foreach (var removed in removedColumns.ToList())
            {
                if (SchemaFixProvider.IsKnownRename(added, removed, sourcePath))
                {
                    var colIdx = sourceKeyList.IndexOf(removed) + 1;
                    changedColumns.Add($"(Col {colIdx}) {removed} -> {added}");
                    remainingAdded.Remove(added);
                    remainingRemoved.Remove(removed);
                    break;
                }
            }
        }

        // If equal remaining counts, treat as renames
        if (remainingAdded.Count == remainingRemoved.Count && remainingAdded.Count > 0)
        {
            foreach (var (a, r) in remainingAdded.Zip(remainingRemoved))
            {
                var colIdx = sourceKeyList.IndexOf(r) + 1;
                changedColumns.Add($"(Col {colIdx}) {r} -> {a}");
            }
            remainingAdded.Clear();
            remainingRemoved.Clear();
        }

        var finalAddedColumns = remainingAdded
            .Select(c => $"(Col {targetKeyList.IndexOf(c) + 1}) {c}")
            .ToList();
        var finalRemovedColumns = remainingRemoved
            .Select(c => $"(Col {sourceKeyList.IndexOf(c) + 1}) {c}")
            .ToList();

        // Row diffs
        var sourceRowKeys = RowInstanceKeyHelper.BuildKeys(sourceData[rowHeaderColumn]);
        var targetRowKeys = RowInstanceKeyHelper.BuildKeys(targetData[rowHeaderColumn]);

        var addedRowsTask = Task.Run(() => targetRowKeys.Except(sourceRowKeys).ToList());
        var removedRowsDict = DiffEngine.GetRemovedRows(sourceRowKeys, targetRowKeys);
        var allRemovedRows = DiffEngine.ExpandCounts(removedRowsDict);
        var addedRows = addedRowsTask.Result;

        var sourceRowIndexMap = sourceRowKeys
            .Select((key, index) => new { key, index })
            .ToDictionary(x => x.key, x => x.index);
        var targetRowIndexMap = targetRowKeys
            .Select((key, index) => new { key, index })
            .ToDictionary(x => x.key, x => x.index);

        // Identify row renames
        var changedRows = new List<string>();
        var processedAdded = new HashSet<RowInstanceKey>();
        var processedRemoved = new HashSet<RowInstanceKey>();

        foreach (var added in addedRows)
        {
            foreach (var removed in allRemovedRows)
            {
                if (!processedAdded.Contains(added) && !processedRemoved.Contains(removed) &&
                    SchemaFixProvider.IsKnownRename(added.Name, removed.Name, sourcePath))
                {
                    var srcRow = sourceRowIndexMap[removed] + 1;
                    changedRows.Add($"(Row {srcRow}) {removed.Name} -> {added.Name}");
                    processedAdded.Add(added);
                    processedRemoved.Add(removed);
                }
            }
        }

        if (addedRows.Count == allRemovedRows.Count)
        {
            foreach (var (added, removed) in addedRows.Zip(allRemovedRows))
            {
                if (!processedAdded.Contains(added) && !processedRemoved.Contains(removed))
                {
                    var srcRow = sourceRowIndexMap[removed] + 1;
                    changedRows.Add($"(Row {srcRow}) {removed.Name} -> {added.Name}");
                    processedAdded.Add(added);
                    processedRemoved.Add(removed);
                }
            }
        }

        var finalAddedRows = addedRows
            .Where(r => !processedAdded.Contains(r))
            .Select(r => $"(Row {targetRowIndexMap[r] + 1}) {r.Name}")
            .ToList();
        var finalRemovedRows = allRemovedRows
            .Where(r => !processedRemoved.Contains(r))
            .Select(r => $"(Row {sourceRowIndexMap[r] + 1}) {r.Name}")
            .ToList();

        // Value-level diffs
        var groupedDifferences = DiffEngine.GetGroupedDifferences(
            sourceData, targetData, allHeaders, rowHeaderColumn, includeNewRows);

        return new CompareResult(
            Path.GetFileName(sourcePath),
            finalAddedColumns,
            finalRemovedColumns,
            changedColumns,
            finalAddedRows,
            finalRemovedRows,
            changedRows,
            groupedDifferences);
    }

    public static List<CompareResult> CompareFolder(
        string sourcePath, string targetPath, bool includeNewRows,
        Action<string>? onProgress = null)
    {
        var results = new List<CompareResult>();
        var sourceFiles = Directory.GetFiles(sourcePath, "*.txt");
        var targetFiles = Directory.GetFiles(targetPath, "*.txt");

        foreach (var sourceFile in sourceFiles)
        {
            var fileName = Path.GetFileName(sourceFile);
            var targetFile = Array.Find(targetFiles,
                f => Path.GetFileName(f)!.Equals(fileName, StringComparison.OrdinalIgnoreCase));

            if (targetFile is not null)
            {
                onProgress?.Invoke(fileName);
                results.Add(CompareFile(sourceFile, targetFile, includeNewRows));
            }
        }

        return results;
    }
}

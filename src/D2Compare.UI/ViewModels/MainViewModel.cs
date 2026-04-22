using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Avalonia.Styling;

using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

using D2Compare.Core.Models;
using D2Compare.Core.Services;
using D2Compare.Services;
using D2Compare.Views;

using Microsoft.Win32;

namespace D2Compare.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly TopLevel _topLevel;

    [ObservableProperty] private ObservableCollection<string> _sourceVersions = new();
    [ObservableProperty] private ObservableCollection<string> _targetVersions = new();
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsSourceCustom))]
    private int _selectedSourceIndex = -1;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsTargetCustom))]
    private int _selectedTargetIndex = -1;

    [ObservableProperty] private ObservableCollection<string> _fileList = new();
    [ObservableProperty] private int _selectedFileIndex = -1;

    [ObservableProperty] private bool _includeNewRows;
    [ObservableProperty] private bool _showOnlyNewRows;
    [ObservableProperty] private bool _omitUnchangedFiles = true;
    [ObservableProperty] private bool _batchReloadNeeded;
    [ObservableProperty] private bool _isDarkMode;

    [ObservableProperty] private bool _convertColumns = true;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsRowModeNone))]
    [NotifyPropertyChangedFor(nameof(IsRowModeAppendOriginal))]
    [NotifyPropertyChangedFor(nameof(IsRowModeAppendTarget))]
    private RowConversionMode _rowConversionMode = RowConversionMode.None;

    // No row conversion (keep source rows)
    public bool IsRowModeNone
    {
        get => RowConversionMode == RowConversionMode.None;
        set { if (value) RowConversionMode = RowConversionMode.None; }
    }

    // Append original data at end
    public bool IsRowModeAppendOriginal
    {
        get => RowConversionMode == RowConversionMode.AppendOriginalAtEnd;
        set { if (value) RowConversionMode = RowConversionMode.AppendOriginalAtEnd; }
    }

    // Append target data at end
    public bool IsRowModeAppendTarget
    {
        get => RowConversionMode == RowConversionMode.AppendTargetAtEnd;
        set { if (value) RowConversionMode = RowConversionMode.AppendTargetAtEnd; }
    }

    [ObservableProperty] private string _statusText = "";
    [ObservableProperty] private bool _isStatusWarning;
    [ObservableProperty] private string _searchText = "";
    [ObservableProperty] private string _activeSearchTerm = "";
    [ObservableProperty] private string _matchLabel = "";
    [ObservableProperty] private bool _isLoading;


    [ObservableProperty] private FormattedDocument? _columnsDocument;
    [ObservableProperty] private FormattedDocument? _rowsDocument;
    [ObservableProperty] private FormattedDocument? _valuesDocument;
    [ObservableProperty] private FormattedDocument? _filesDocument;

    [ObservableProperty] private int _columnsAdded;
    [ObservableProperty] private int _columnsRemoved;
    [ObservableProperty] private int _columnsChanged;
    [ObservableProperty] private int _rowsAdded;
    [ObservableProperty] private int _rowsRemoved;
    [ObservableProperty] private int _rowsChanged;
    [ObservableProperty] private int _valuesChanged;
    [ObservableProperty] private int _valuesNew;

    public bool HasNoFileChanges => FilesDocument is null || FilesDocument.Lines.Count == 0;

    partial void OnFilesDocumentChanged(FormattedDocument? value) => OnPropertyChanged(nameof(HasNoFileChanges));

    public AppSettings Settings => _settings;
    private readonly AppSettings _settings;
    private readonly IReadOnlyList<VersionInfo> _availableVersions;

    private string _sourceFolderPath = "";
    private string _targetFolderPath = "";
    private List<CompareResult> _batchResults = new();

    public string AppVersion => UpdateService.GetCurrentVersion();
    public bool IsSourceCustom => SelectedSourceIndex >= _availableVersions.Count;
    public bool IsTargetCustom => SelectedTargetIndex >= _availableVersions.Count;

    // Auto-update
    [ObservableProperty] private UpdateInfo? _pendingUpdate;
    [ObservableProperty] private bool _isDownloadingUpdate;
    [ObservableProperty] private double _updateDownloadProgress;
    [ObservableProperty] private string _updateStatusText = "";

    public bool HasPendingUpdate => PendingUpdate is not null;
    public string UpdateAvailableText =>
        PendingUpdate is null ? "" : $"Update available: {PendingUpdate.TagName}";

    partial void OnPendingUpdateChanged(UpdateInfo? value)
    {
        OnPropertyChanged(nameof(HasPendingUpdate));
        OnPropertyChanged(nameof(UpdateAvailableText));
    }

    public MainViewModel(TopLevel topLevel)
    {
        _topLevel = topLevel;
        _settings = AppSettings.Load();

        IsDarkMode = _settings.IsDarkMode;
        if (IsDarkMode && Avalonia.Application.Current is { } app)
            app.RequestedThemeVariant = ThemeVariant.Dark;

        _availableVersions = VersionInfo.GetAvailableVersions();

        var names = _availableVersions.Select(v => v.DisplayName).ToList();
        names.Add("Custom");

        SourceVersions = new ObservableCollection<string>(names);
        TargetVersions = new ObservableCollection<string>(names);

        var customIndex = _availableVersions.Count;

        // Restore custom paths before setting indices (only if directory still exists)
        if (!string.IsNullOrEmpty(_settings.CustomSourcePath) && Directory.Exists(_settings.CustomSourcePath))
        {
            _sourceFolderPath = _settings.CustomSourcePath;
            SourceVersions[customIndex] = _settings.CustomSourcePath;
        }
        else if (!string.IsNullOrEmpty(_settings.CustomSourcePath))
        {
            _settings.CustomSourcePath = "";
            if (_settings.SelectedSourceIndex == customIndex)
                _settings.SelectedSourceIndex = -1;
            _settings.Save();
            StatusText = $"Custom source folder no longer exists";
            IsStatusWarning = true;
        }

        if (!string.IsNullOrEmpty(_settings.CustomTargetPath) && Directory.Exists(_settings.CustomTargetPath))
        {
            _targetFolderPath = _settings.CustomTargetPath;
            TargetVersions[customIndex] = _settings.CustomTargetPath;
        }
        else if (!string.IsNullOrEmpty(_settings.CustomTargetPath))
        {
            _settings.CustomTargetPath = "";
            if (_settings.SelectedTargetIndex == customIndex)
                _settings.SelectedTargetIndex = -1;
            _settings.Save();
            StatusText = $"Custom target folder no longer exists";
            IsStatusWarning = true;
        }

        if (_settings.SelectedSourceIndex >= 0 && _settings.SelectedSourceIndex < SourceVersions.Count)
        {
            if (_settings.SelectedSourceIndex < customIndex || !string.IsNullOrEmpty(_settings.CustomSourcePath))
                SelectedSourceIndex = _settings.SelectedSourceIndex;
        }
        if (_settings.SelectedTargetIndex >= 0 && _settings.SelectedTargetIndex < TargetVersions.Count)
        {
            if (_settings.SelectedTargetIndex < customIndex || !string.IsNullOrEmpty(_settings.CustomTargetPath))
                SelectedTargetIndex = _settings.SelectedTargetIndex;
        }

        if (!_settings.DisableUpdateCheck)
            _ = CheckForUpdateAsync();
    }

    partial void OnSelectedSourceIndexChanged(int value)
    {
        if (value < 0) return;

        if (value < _availableVersions.Count)
            _sourceFolderPath = _availableVersions[value].GetPath();
        else if (!string.IsNullOrEmpty(_settings.CustomSourcePath))
            _sourceFolderPath = _settings.CustomSourcePath;

        _settings.SelectedSourceIndex = value;
        _settings.Save();
        OnTargetChanged();
    }

    partial void OnSelectedTargetIndexChanged(int value)
    {
        if (value < 0) return;

        if (value < _availableVersions.Count)
            _targetFolderPath = _availableVersions[value].GetPath();
        else if (!string.IsNullOrEmpty(_settings.CustomTargetPath))
            _targetFolderPath = _settings.CustomTargetPath;

        _settings.SelectedTargetIndex = value;
        _settings.Save();
        OnTargetChanged();
    }

    partial void OnSelectedFileIndexChanged(int value)
    {
        OpenSourceFileCommand.NotifyCanExecuteChanged();
        OpenTargetFileCommand.NotifyCanExecuteChanged();
        if (value < 0 || value >= FileList.Count) return;
        BatchReloadNeeded = false;
        _ = RunSingleComparisonAsync();
    }

    partial void OnIncludeNewRowsChanged(bool value)
    {
        if (!value) ShowOnlyNewRows = false;
        if (SelectedFileIndex >= 0 && _sourceFolderPath.Length > 0)
            _ = RunSingleComparisonAsync();
        else if (_batchResults.Count > 0)
            BatchReloadNeeded = true;
    }

    partial void OnShowOnlyNewRowsChanged(bool value)
    {
        if (SelectedFileIndex >= 0 && _sourceFolderPath.Length > 0)
            _ = RunSingleComparisonAsync();
        else if (_batchResults.Count > 0)
            BatchReloadNeeded = true;
    }

    partial void OnOmitUnchangedFilesChanged(bool value)
    {
        if (_batchResults.Count > 0 && SelectedFileIndex < 0)
            BatchReloadNeeded = true;
    }

    [RelayCommand]
    private async Task ExecuteSearch()
    {
        IsLoading = true;
        StatusText = string.IsNullOrEmpty(SearchText) ? "Clearing..." : "Searching...";

        // Yield a frame so the loading indicator renders before the heavy highlight rebuild
        await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(
            () => { }, Avalonia.Threading.DispatcherPriority.Render);

        ActiveSearchTerm = SearchText;

        if (string.IsNullOrEmpty(SearchText))
        {
            MatchLabel = "";
            IsLoading = false;
            StatusText = "";
            return;
        }

        var doc = ValuesDocument;
        if (doc is null)
        {
            MatchLabel = "";
            IsLoading = false;
            StatusText = "";
            return;
        }

        var term = SearchText;
        int count = await Task.Run(() =>
        {
            int c = 0;
            foreach (var line in doc.Lines)
            {
                foreach (var span in line.Spans)
                {
                    int pos = 0;
                    while (pos < span.Text.Length)
                    {
                        int idx = span.Text.IndexOf(term, pos, StringComparison.OrdinalIgnoreCase);
                        if (idx == -1) break;
                        c++;
                        pos = idx + term.Length;
                    }
                }
            }
            return c;
        });

        MatchLabel = count > 0 ? $"{count} matches" : "0 matches";
        IsLoading = false;
        StatusText = "";
    }

    [RelayCommand]
    private async Task BatchLoad()
    {
        if (_sourceFolderPath.Length == 0 || _targetFolderPath.Length == 0) return;
        if (!Directory.Exists(_sourceFolderPath) || !Directory.Exists(_targetFolderPath))
        {
            StatusText = "One or both folder paths no longer exist";
            IsStatusWarning = true;
            return;
        }

        IsLoading = true;
        IsStatusWarning = false;
        StatusText = "Loading...";

        try
        {
            var sourcePath = _sourceFolderPath;
            var targetPath = _targetFolderPath;
            var includeNew = IncludeNewRows;
            var omitUnchanged = OmitUnchangedFiles;

            _batchResults = await Task.Run(() =>
                CompareService.CompareFolder(sourcePath, targetPath, includeNew));

            SelectedFileIndex = -1;
            RebuildBatchDocuments();
            BatchReloadNeeded = false;
            SearchText = "";
        }
        finally
        {
            IsLoading = false;
            StatusText = "";
        }
    }

    [RelayCommand]
    private void BrowseSourceCustom() => _ = SelectCustomFolder(isSource: true);

    [RelayCommand]
    private void BrowseTargetCustom() => _ = SelectCustomFolder(isSource: false);

    [RelayCommand]
    private async Task ConvertToTarget()
    {
        if (_sourceFolderPath.Length == 0 || _targetFolderPath.Length == 0)
        {
            StatusText = "Select source and target folders first";
            IsStatusWarning = true;
            return;
        }
        if (!Directory.Exists(_sourceFolderPath) || !Directory.Exists(_targetFolderPath))
        {
            StatusText = "One or both folder paths no longer exist";
            IsStatusWarning = true;
            return;
        }
        if (!ConvertColumns)
        {
            StatusText = "Select at least one conversion option (columns)";
            IsStatusWarning = true;
            return;
        }

        var storage = _topLevel.StorageProvider;
        if (!storage.CanPickFolder) return;

        var folders = await storage.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = "Select Output Folder for Converted Files",
            AllowMultiple = false,
        });

        if (folders.Count == 0) return;

        var outputPath = folders[0].Path.LocalPath;
        IsLoading = true;
        StatusText = "Converting...";

        try
        {
            ConvertService.ConvertFolder(
                _sourceFolderPath,
                _targetFolderPath,
                outputPath,
                ConvertColumns,
                RowConversionMode,
                name => StatusText = $"Converting {name}...");
            StatusText = $"Done. Files written to {outputPath}";
        }
        catch (Exception ex)
        {
            StatusText = $"Error: {ex.Message}";
        }
        finally
        {
            IsLoading = false;
        }
    }

    [RelayCommand]
    private void ToggleTheme()
    {
        IsDarkMode = !IsDarkMode;

        if (Avalonia.Application.Current is { } app)
            app.RequestedThemeVariant = IsDarkMode ? ThemeVariant.Dark : ThemeVariant.Light;

        _settings.IsDarkMode = IsDarkMode;
        _settings.Save();
    }

    [RelayCommand]
    private static void OpenOriginalProject() =>
        OpenUrl("https://github.com/locbones/D2Compare");

    private bool HasFileSelected() => SelectedFileIndex >= 0 && SelectedFileIndex < FileList.Count;

    [RelayCommand(CanExecute = nameof(HasFileSelected))]
    private void OpenSourceFile()
    {
        if (SelectedFileIndex < 0 || SelectedFileIndex >= FileList.Count)
            return;

        var filePath = Path.Combine(_sourceFolderPath, FileList[SelectedFileIndex]);
        OpenFileViewer(filePath, role: "--source", originPath: _sourceFolderPath);
    }

    [RelayCommand(CanExecute = nameof(HasFileSelected))]
    private void OpenTargetFile()
    {
        if (SelectedFileIndex < 0 || SelectedFileIndex >= FileList.Count)
            return;

        var filePath = Path.Combine(_targetFolderPath, FileList[SelectedFileIndex]);
        OpenFileViewer(filePath, role: "--target", originPath: _targetFolderPath);
    }


    private void OpenFileViewer(string filePath, string role, string originPath)
    {
        if (!File.Exists(filePath))
            return;

        string absFilePath = Path.GetFullPath(filePath);
        string absOriginPath = Path.GetFullPath(originPath);
        string args = $"{role} --path \"{absOriginPath}\" \"{absFilePath}\"";

        ProcessStartInfo psi;

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            var viewerPath = GetWindowsViewerPathFromRegistry();

            if (!string.IsNullOrWhiteSpace(viewerPath) && File.Exists(viewerPath))
            {
                psi = new ProcessStartInfo
                {
                    FileName = viewerPath,
                    Arguments = args,
                    UseShellExecute = false
                };
            }
            else
            {
                psi = new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                };
            }
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            var linuxViewerPath = _settings.LinuxViewerPath;

            psi = !string.IsNullOrWhiteSpace(linuxViewerPath) && File.Exists(linuxViewerPath)
                ? new ProcessStartInfo
                {
                    FileName = linuxViewerPath,
                    Arguments = args,
                    UseShellExecute = false
                }
                : new ProcessStartInfo
                {
                    FileName = "xdg-open",
                    Arguments = $"\"{filePath}\"",
                    UseShellExecute = false
                };
        }
        else
        {
            throw new PlatformNotSupportedException("Unsupported operating system.");
        }

        Process.Start(psi);
    }

    [SupportedOSPlatform("windows")]
    private static string? GetWindowsViewerPathFromRegistry()
    {
        const string valueName = "ExecutablePath";
        const string KeyPath = @"Software\d2_horadrim";
        using (var legacyKey = Registry.CurrentUser.OpenSubKey(KeyPath, writable: false))
        {
            return legacyKey?.GetValue(valueName) as string;
        }
    }

    private void RebuildBatchDocuments()
    {
        var omit = OmitUnchangedFiles;

        var colResults = omit
            ? _batchResults.Where(r => r.ChangedColumns.Count > 0 || r.AddedColumns.Count > 0 || r.RemovedColumns.Count > 0)
            : _batchResults;
        var rowResults = omit
            ? _batchResults.Where(r => r.ChangedRows.Count > 0 || r.AddedRows.Count > 0 || r.RemovedRows.Count > 0)
            : _batchResults;
        var valResults = omit
            ? _batchResults.Where(r => r.GroupedDifferences.Count > 0)
            : _batchResults;

        ColumnsDocument = FormattedTextBuilder.MergeDocuments(
            colResults.Select(r => FormattedTextBuilder.BuildColumnDiffs(r, true)));
        RowsDocument = FormattedTextBuilder.MergeDocuments(
            rowResults.Select(r => FormattedTextBuilder.BuildRowDiffs(r, true)));
        var onlyNew = ShowOnlyNewRows;
        ValuesDocument = FormattedTextBuilder.MergeDocuments(
            valResults.Select(r => FormattedTextBuilder.BuildValueDiffs(r, true, onlyNew)));
        UpdateStats(_batchResults);
    }

    private async Task RunSingleComparisonAsync()
    {
        if (SelectedFileIndex < 0 || SelectedFileIndex >= FileList.Count) return;

        var sourcePath = Path.Combine(_sourceFolderPath, FileList[SelectedFileIndex]);
        var targetPath = Path.Combine(_targetFolderPath, FileList[SelectedFileIndex]);

        if (!File.Exists(sourcePath) || !File.Exists(targetPath))
        {
            StatusText = "File not found";
            IsStatusWarning = true;
            return;
        }

        IsLoading = true;
        IsStatusWarning = false;
        StatusText = "Loading...";
        var includeNew = IncludeNewRows;
        var onlyNew = ShowOnlyNewRows;

        try
        {
            var result = await Task.Run(() =>
                CompareService.CompareFile(sourcePath, targetPath, includeNew));

            ColumnsDocument = FormattedTextBuilder.BuildColumnDiffs(result, false);
            RowsDocument = FormattedTextBuilder.BuildRowDiffs(result, false);
            ValuesDocument = FormattedTextBuilder.BuildValueDiffs(result, false, onlyNew);
            UpdateStats(new[] { result });

            SearchText = "";
        }
        finally
        {
            IsLoading = false;
            StatusText = "";
        }
    }

    private void UpdateStats(IEnumerable<CompareResult> results)
    {
        int ca = 0, cr = 0, cc = 0, ra = 0, rr = 0, rc = 0, vc = 0, vn = 0;
        foreach (var r in results)
        {
            ca += r.AddedColumns.Count;
            cr += r.RemovedColumns.Count;
            cc += r.ChangedColumns.Count;
            ra += r.AddedRows.Count;
            rr += r.RemovedRows.Count;
            rc += r.ChangedRows.Count;
            foreach (var g in r.GroupedDifferences)
            {
                if (g.IsNew) vn++;
                else vc++;
            }
        }
        ColumnsAdded = ca; ColumnsRemoved = cr; ColumnsChanged = cc;
        RowsAdded = ra; RowsRemoved = rr; RowsChanged = rc;
        ValuesChanged = ShowOnlyNewRows ? 0 : vc;
        ValuesNew = vn;
    }

    private void OnTargetChanged()
    {
        if (_sourceFolderPath.Length == 0 || _targetFolderPath.Length == 0) return;
        if (!Directory.Exists(_sourceFolderPath) || !Directory.Exists(_targetFolderPath)) return;

        IsStatusWarning = false;
        StatusText = "";
        var fileResult = FileDiscovery.DiscoverFiles(_sourceFolderPath, _targetFolderPath);

        FileList = new ObservableCollection<string>(fileResult.CommonFiles);
        FilesDocument = FormattedTextBuilder.BuildFileDiffs(fileResult);
        SelectedFileIndex = -1;
        ColumnsDocument = null;
        RowsDocument = null;
        ValuesDocument = null;
    }

    private async Task SelectCustomFolder(bool isSource)
    {
        var storage = _topLevel.StorageProvider;
        if (!storage.CanPickFolder) return;

        var folders = await storage.OpenFolderPickerAsync(new FolderPickerOpenOptions
        {
            Title = isSource ? "Select Source Folder" : "Select Target Folder",
            AllowMultiple = false,
        });

        if (folders.Count > 0)
        {
            var path = folders[0].Path.LocalPath;
            var customIndex = _availableVersions.Count;

            if (isSource)
            {
                _sourceFolderPath = path;
                SourceVersions[customIndex] = path;
                _settings.SelectedSourceIndex = customIndex;
                _settings.CustomSourcePath = path;
                _settings.Save();
                SelectedSourceIndex = customIndex;
                OnPropertyChanged(nameof(IsSourceCustom));
            }
            else
            {
                _targetFolderPath = path;
                TargetVersions[customIndex] = path;
                _settings.SelectedTargetIndex = customIndex;
                _settings.CustomTargetPath = path;
                _settings.Save();
                SelectedTargetIndex = customIndex;
                OnPropertyChanged(nameof(IsTargetCustom));
            }

            OnTargetChanged();
        }
    }

    private async Task CheckForUpdateAsync()
    {
        var info = await UpdateService.CheckForUpdateAsync();
        if (info is not null)
            PendingUpdate = info;
    }

    [RelayCommand]
    private async Task ApplyUpdate()
    {
        if (PendingUpdate is null) return;

        var dialog = new UpdateDialog(PendingUpdate.Version, PendingUpdate.ReleaseNotes);
        var result = await dialog.ShowDialog<bool?>((Window)_topLevel);
        if (result is not true) return;

        IsDownloadingUpdate = true;
        UpdateStatusText = "Downloading...";
        var progress = new Progress<double>(p =>
        {
            UpdateDownloadProgress = p;
            UpdateStatusText = $"Downloading... {p:P0}";
        });

        var stagingDir = await UpdateService.DownloadUpdateAsync(PendingUpdate, progress);
        if (stagingDir is null)
        {
            UpdateStatusText = "Download failed";
            IsDownloadingUpdate = false;
            return;
        }

        UpdateStatusText = "Applying update...";
        UpdateService.LaunchUpdateScriptAndExit(stagingDir);
    }

    [RelayCommand]
    private void DismissUpdate() => PendingUpdate = null;

    [RelayCommand(CanExecute = nameof(HasColumnsContent))]
    private Task SaveColumnsAsText() => SaveDocumentAsText(ColumnsDocument, "Columns.Altered", "Columns Altered");
    private bool HasColumnsContent() => ColumnsDocument is { Lines.Count: > 0 };
    public bool ColumnsHasContent => HasColumnsContent();

    [RelayCommand(CanExecute = nameof(HasRowsContent))]
    private Task SaveRowsAsText() => SaveDocumentAsText(RowsDocument, "Rows.Altered", "Rows Altered");
    private bool HasRowsContent() => RowsDocument is { Lines.Count: > 0 };
    public bool RowsHasContent => HasRowsContent();

    [RelayCommand(CanExecute = nameof(HasValuesContent))]
    private Task SaveValuesAsText() => SaveDocumentAsText(ValuesDocument, "Values", "Values");
    private bool HasValuesContent() => ValuesDocument is { Lines.Count: > 0 };
    public bool ValuesHasContent => HasValuesContent();

    partial void OnColumnsDocumentChanged(FormattedDocument? value)
    {
        SaveColumnsAsTextCommand.NotifyCanExecuteChanged();
        OnPropertyChanged(nameof(ColumnsHasContent));
    }
    partial void OnRowsDocumentChanged(FormattedDocument? value)
    {
        SaveRowsAsTextCommand.NotifyCanExecuteChanged();
        OnPropertyChanged(nameof(RowsHasContent));
    }
    partial void OnValuesDocumentChanged(FormattedDocument? value)
    {
        SaveValuesAsTextCommand.NotifyCanExecuteChanged();
        OnPropertyChanged(nameof(ValuesHasContent));
    }

    private async Task SaveDocumentAsText(FormattedDocument? document, string panelSuffix, string panelName)
    {
        if (document is null || document.Lines.Count == 0) return;

        bool isDisplayAll = _batchResults.Count > 0 && SelectedFileIndex < 0;
        string prefix;
        if (isDisplayAll)
            prefix = "DisplayAll";
        else if (SelectedFileIndex >= 0 && SelectedFileIndex < FileList.Count)
            prefix = Path.GetFileNameWithoutExtension(FileList[SelectedFileIndex]);
        else
            prefix = "output";

        var storage = _topLevel.StorageProvider;
        var file = await storage.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = "Save As Text",
            SuggestedFileName = $"{prefix}-{panelSuffix}.txt",
            FileTypeChoices = [new FilePickerFileType("Text File") { Patterns = ["*.txt"] }],
        });

        if (file is null) return;

        var sourceName = SelectedSourceIndex >= 0 && SelectedSourceIndex < SourceVersions.Count
            ? SourceVersions[SelectedSourceIndex] : "Unknown";
        var targetName = SelectedTargetIndex >= 0 && SelectedTargetIndex < TargetVersions.Count
            ? TargetVersions[SelectedTargetIndex] : "Unknown";

        var headerLines = new List<string>
        {
            $"Source:   {sourceName}",
            $"Target:   {targetName}",
        };
        headerLines.Add($"File:     {(isDisplayAll ? "All Files" : $"{prefix}.txt")}");
        headerLines.Add($"Summary:  {panelName}");
        headerLines.Add(new string('-', 40));
        headerLines.Add("");

        var content = document.Lines.Select(l => string.Concat(l.Spans.Select(s => s.Text)));
        await File.WriteAllLinesAsync(file.Path.LocalPath, headerLines.Concat(content));
    }

    public void InitializeFromArguments(string? sourceFolder, string? targetFolder, string? filePath)
    {
        var customIndex = _availableVersions.Count;

        if (!string.IsNullOrWhiteSpace(sourceFolder) && Directory.Exists(sourceFolder))
        {
            // Treat as custom source folder
            _settings.CustomSourcePath = sourceFolder;
            _settings.SelectedSourceIndex = customIndex;
            _settings.Save();

            if (SourceVersions.Count > customIndex)
                SourceVersions[customIndex] = sourceFolder;

            _sourceFolderPath = sourceFolder;
            SelectedSourceIndex = customIndex;
        }

        if (!string.IsNullOrWhiteSpace(targetFolder) && Directory.Exists(targetFolder))
        {
            // Treat as custom target folder
            _settings.CustomTargetPath = targetFolder;
            _settings.SelectedTargetIndex = customIndex;
            _settings.Save();

            if (TargetVersions.Count > customIndex)
                TargetVersions[customIndex] = targetFolder;

            _targetFolderPath = targetFolder;
            SelectedTargetIndex = customIndex;
        }

        // If a specific file was provided, try to select and compare it
        if (!string.IsNullOrWhiteSpace(filePath))
        {
            // Ensure the file list is built if we have valid folders
            if (FileList.Count == 0 &&
                _sourceFolderPath.Length > 0 &&
                _targetFolderPath.Length > 0 &&
                Directory.Exists(_sourceFolderPath) &&
                Directory.Exists(_targetFolderPath))
            {
                OnTargetChanged();
            }

            if (FileList.Count > 0)
            {
                var fileName = Path.GetFileName(filePath);
                if (string.IsNullOrEmpty(fileName))
                    fileName = filePath;

                var index = FileList.IndexOf(fileName);
                if (index >= 0)
                    SelectedFileIndex = index;
            }
        }
    }

    private static void OpenUrl(string url) =>
        Process.Start(new ProcessStartInfo { FileName = url, UseShellExecute = true });
}
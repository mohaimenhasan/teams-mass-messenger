using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using MSSLTeamsMessenger.Helpers;
using MSSLTeamsMessenger.Models;
using MSSLTeamsMessenger.Services;

namespace MSSLTeamsMessenger;

public partial class MainWindow : Window
{
    private readonly ObservableCollection<AliasEntry> _aliases = new();
    private readonly TeamsMessengerService _teamsService = new();
    private CancellationTokenSource? _cts;

    public MainWindow()
    {
        InitializeComponent();
        MessageBox.Text = LoadMessage();
        LoadAliases();
    }

    private string LoadMessage()
    {
        var exeDir = AppDomain.CurrentDomain.BaseDirectory;
        var candidates = new[]
        {
            Path.Combine(exeDir, "message.txt"),
            Path.Combine(exeDir, "..", "..", "..", "message.txt"), // dev: bin/Debug/net8.0-windows/
        };

        foreach (var c in candidates)
        {
            if (File.Exists(c))
            {
                Log($"Loaded message from {Path.GetFullPath(c)}");
                return File.ReadAllText(c);
            }
        }

        Log("WARNING: message.txt not found. Enter your message manually.");
        return "";
    }

    private void LoadAliases()
    {
        // Look for alias.txt next to the exe, or in the parent (repo root)
        var exeDir = AppDomain.CurrentDomain.BaseDirectory;
        var candidates = new[]
        {
            Path.Combine(exeDir, "alias.txt"),
            Path.Combine(exeDir, "..", "..", "..", "..", "alias.txt"), // dev: bin/Debug/net8.0-windows/ -> repo root
        };

        string? aliasFile = null;
        foreach (var c in candidates)
        {
            if (File.Exists(c)) { aliasFile = c; break; }
        }

        if (aliasFile == null)
        {
            Log("WARNING: alias.txt not found. Use File > Open or place it next to the exe.");
            return;
        }

        var entries = AliasLoader.LoadFromFile(aliasFile);
        _aliases.Clear();
        foreach (var e in entries)
            _aliases.Add(e);

        AliasGrid.ItemsSource = _aliases;
        UpdateCount();
        Log($"Loaded {_aliases.Count} unique aliases from {Path.GetFullPath(aliasFile)}");
    }

    private void UpdateCount()
    {
        var selected = _aliases.Count(a => a.IsSelected);
        CountText.Text = $"{selected}/{_aliases.Count} selected";
    }

    private void SelectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var a in _aliases) a.IsSelected = true;
        UpdateCount();
    }

    private void DeselectAll_Click(object sender, RoutedEventArgs e)
    {
        foreach (var a in _aliases) a.IsSelected = false;
        UpdateCount();
    }

    private void Log(string msg)
    {
        var entry = $"[{DateTime.Now:HH:mm:ss}] {msg}";
        LogList.Items.Add(entry);
        LogList.ScrollIntoView(entry);
    }

    private int GetDelay()
    {
        if (int.TryParse(DelayBox.Text, out int seconds) && seconds >= 1)
            return seconds * 1000;
        return 7000;
    }

    private void SetSendingState(bool isSending)
    {
        TestBtn.IsEnabled = !isSending;
        SendBtn.IsEnabled = !isSending;
        StopBtn.IsEnabled = isSending;
        AliasGrid.IsReadOnly = isSending;
    }

    private async void Test_Click(object sender, RoutedEventArgs e)
    {
        var message = MessageBox.Text;
        if (string.IsNullOrWhiteSpace(message))
        {
            System.Windows.MessageBox.Show("Message is empty.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        SetSendingState(true);
        _cts = new CancellationTokenSource();

        try
        {
            Log("=== TEST MODE: Sending to mohaimenkhan@microsoft.com ===");
            await _teamsService.SendMessageToAlias("mohaimenkhan", message, GetDelay(), _cts.Token,
                msg => Dispatcher.Invoke(() => Log(msg)));
            Log("=== TEST COMPLETE ===");
        }
        catch (OperationCanceledException)
        {
            Log("Test cancelled.");
        }
        catch (Exception ex)
        {
            Log($"ERROR: {ex.Message}");
        }
        finally
        {
            SetSendingState(false);
        }
    }

    private async void Send_Click(object sender, RoutedEventArgs e)
    {
        var message = MessageBox.Text;
        if (string.IsNullOrWhiteSpace(message))
        {
            System.Windows.MessageBox.Show("Message is empty.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var selected = _aliases.Where(a => a.IsSelected).ToList();
        if (selected.Count == 0)
        {
            System.Windows.MessageBox.Show("No aliases selected.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var confirm = System.Windows.MessageBox.Show(
            $"Send message to {selected.Count} people?\n\nMake sure Teams is open and signed in.\nDo NOT use the computer during sending.",
            "Confirm Send",
            MessageBoxButton.YesNo, MessageBoxImage.Question);

        if (confirm != MessageBoxResult.Yes) return;

        SetSendingState(true);
        _cts = new CancellationTokenSource();
        ProgressBar.Maximum = selected.Count;
        ProgressBar.Value = 0;

        int sent = 0, failed = 0;

        try
        {
            for (int i = 0; i < selected.Count; i++)
            {
                _cts.Token.ThrowIfCancellationRequested();

                var alias = selected[i];
                ProgressText.Text = $"Sending {i + 1} of {selected.Count}: {alias.Alias}";

                try
                {
                    alias.Status = "Sending...";
                    await _teamsService.SendMessageToAlias(alias.Alias, message, GetDelay(), _cts.Token,
                        msg => Dispatcher.Invoke(() => Log(msg)));
                    alias.Status = "✅ Sent";
                    sent++;
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex)
                {
                    alias.Status = $"❌ Error: {ex.Message}";
                    Log($"ERROR sending to {alias.Alias}: {ex.Message}");
                    failed++;
                }

                ProgressBar.Value = i + 1;
            }

            Log($"=== BATCH COMPLETE: {sent} sent, {failed} failed ===");
        }
        catch (OperationCanceledException)
        {
            Log($"Batch cancelled. {sent} sent, {failed} failed, {selected.Count - sent - failed} skipped.");
        }
        finally
        {
            SetSendingState(false);
            ProgressText.Text = "";
        }
    }

    private void Stop_Click(object sender, RoutedEventArgs e)
    {
        _cts?.Cancel();
        Log("Stop requested...");
    }
}
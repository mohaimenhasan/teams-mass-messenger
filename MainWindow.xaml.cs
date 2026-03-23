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

    private void ImportClipboard_Click(object sender, RoutedEventArgs e)
    {
        if (!Clipboard.ContainsText())
        {
            System.Windows.MessageBox.Show("Clipboard is empty. Copy the team members list from Teams first.",
                "No Data", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var text = Clipboard.GetText();
        var parsed = ParseTeamsMemberData(text);

        if (parsed.Count == 0)
        {
            System.Windows.MessageBox.Show(
                "No email addresses or aliases found in clipboard.\n\n" +
                "How to copy members from Teams:\n" +
                "1. Open your team in Teams\n" +
                "2. Click \"...\" > Manage team > Members\n" +
                "3. Select all (Ctrl+A) and copy (Ctrl+C)\n" +
                "4. Come back here and click Import again",
                "No Members Found", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        // Merge with existing, deduplicating
        var existing = new HashSet<string>(_aliases.Select(a => a.Alias), StringComparer.OrdinalIgnoreCase);
        int added = 0;
        foreach (var alias in parsed)
        {
            if (existing.Add(alias))
            {
                _aliases.Add(new AliasEntry { Alias = alias });
                added++;
            }
        }

        UpdateCount();
        Log($"Imported {added} new aliases from clipboard ({parsed.Count} total parsed, {parsed.Count - added} duplicates skipped).");
    }

    private static List<string> ParseTeamsMemberData(string text)
    {
        var results = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Split on common delimiters: newlines, tabs, commas, semicolons
        var tokens = text.Split(new[] { '\r', '\n', '\t', ',', ';' }, StringSplitOptions.RemoveEmptyEntries);

        foreach (var raw in tokens)
        {
            var token = raw.Trim();

            // Match email addresses (user@domain)
            var emailMatch = System.Text.RegularExpressions.Regex.Match(token,
                @"[\w.\-]+@[\w.\-]+\.[\w]+");
            if (emailMatch.Success)
            {
                var email = emailMatch.Value;
                // Extract alias (part before @)
                var alias = email.Substring(0, email.IndexOf('@'));
                if (!string.IsNullOrWhiteSpace(alias))
                    results.Add(alias);
                continue;
            }

            // Skip tokens that look like role labels, dates, or UI artifacts
            if (token.Length > 50) continue;
            if (token.Equals("Owner", StringComparison.OrdinalIgnoreCase)) continue;
            if (token.Equals("Member", StringComparison.OrdinalIgnoreCase)) continue;
            if (token.Equals("Guest", StringComparison.OrdinalIgnoreCase)) continue;
            if (token.Equals("Members", StringComparison.OrdinalIgnoreCase)) continue;
            if (token.Equals("Owners", StringComparison.OrdinalIgnoreCase)) continue;
            if (token.Equals("Name", StringComparison.OrdinalIgnoreCase)) continue;
            if (token.Equals("Role", StringComparison.OrdinalIgnoreCase)) continue;
            if (token.Equals("Title", StringComparison.OrdinalIgnoreCase)) continue;
            if (token.Equals("Location", StringComparison.OrdinalIgnoreCase)) continue;
            if (System.Text.RegularExpressions.Regex.IsMatch(token, @"^\d")) continue; // starts with digit
            if (token.Contains(' ') && !token.Contains('@')) continue; // likely a display name, skip

            // Looks like it could be an alias (single word, no spaces)
            if (System.Text.RegularExpressions.Regex.IsMatch(token, @"^[\w.\-]+$") && token.Length >= 2)
            {
                results.Add(token);
            }
        }

        return results.ToList();
    }

    private void ExportAliases_Click(object sender, RoutedEventArgs e)
    {
        if (_aliases.Count == 0)
        {
            System.Windows.MessageBox.Show("No aliases to export.", "Empty", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // Find alias.txt path
        var exeDir = AppDomain.CurrentDomain.BaseDirectory;
        var candidates = new[]
        {
            Path.Combine(exeDir, "alias.txt"),
            Path.Combine(exeDir, "..", "..", "..", "..", "alias.txt"),
        };

        string aliasFile = candidates[0];
        foreach (var c in candidates)
        {
            if (File.Exists(c)) { aliasFile = c; break; }
        }

        var lines = _aliases.Select(a => a.Alias).ToList();
        File.WriteAllLines(aliasFile, lines);
        Log($"Exported {lines.Count} aliases to {Path.GetFullPath(aliasFile)}");
    }
}
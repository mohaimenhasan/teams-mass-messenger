using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using InputSimulatorStandard;
using InputSimulatorStandard.Native;

namespace MSSLTeamsMessenger.Services;

public class TeamsMessengerService
{
    [DllImport("user32.dll")]
    private static extern IntPtr FindWindow(string? lpClassName, string? lpWindowName);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    private readonly InputSimulator _inputSimulator = new();

    /// <summary>
    /// Opens a Teams 1:1 chat with the given email via deep link.
    /// </summary>
    public void OpenChat(string email)
    {
        var url = $"https://teams.microsoft.com/l/chat/0/0?users={Uri.EscapeDataString(email)}";
        Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
    }

    /// <summary>
    /// Attempts to bring the Teams window to the foreground.
    /// Returns true if a Teams window was found and activated.
    /// </summary>
    public bool FocusTeamsWindow()
    {
        IntPtr teamsHwnd = IntPtr.Zero;

        EnumWindows((hWnd, _) =>
        {
            if (!IsWindowVisible(hWnd))
                return true;

            var sb = new System.Text.StringBuilder(512);
            GetWindowText(hWnd, sb, sb.Capacity);
            var title = sb.ToString();

            // Match both new Teams ("Microsoft Teams") and classic Teams
            if (title.Contains("Microsoft Teams", StringComparison.OrdinalIgnoreCase) ||
                title.Contains("Teams", StringComparison.OrdinalIgnoreCase) && title.Contains("Chat", StringComparison.OrdinalIgnoreCase))
            {
                teamsHwnd = hWnd;
                return false; // stop enumeration
            }
            return true;
        }, IntPtr.Zero);

        if (teamsHwnd != IntPtr.Zero)
        {
            SetForegroundWindow(teamsHwnd);
            return true;
        }
        return false;
    }

    /// <summary>
    /// Sets the clipboard to the message, pastes it, and presses Enter.
    /// Must be called from the STA thread (UI thread).
    /// </summary>
    public void PasteAndSend(string message)
    {
        Clipboard.SetText(message);
        Thread.Sleep(300);

        // Ctrl+V to paste
        _inputSimulator.Keyboard.ModifiedKeyStroke(
            VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);

        Thread.Sleep(500);

        // Enter to send
        _inputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
    }

    /// <summary>
    /// Full workflow: open chat, wait for Teams to load it, focus, paste, send.
    /// If the alias contains spaces, it's treated as a display name and searched in Teams.
    /// </summary>
    public async Task<bool> SendMessageToAlias(string alias, string message, int delayMs,
        CancellationToken ct, Action<string>? log = null)
    {
        bool isDisplayName = alias.Contains(' ');

        if (isDisplayName)
        {
            return await SendViaSearch(alias, message, delayMs, ct, log);
        }
        else
        {
            return await SendViaDeepLink(alias, message, delayMs, ct, log);
        }
    }

    private async Task<bool> SendViaDeepLink(string alias, string message, int delayMs,
        CancellationToken ct, Action<string>? log = null)
    {
        var email = alias.Contains('@') ? alias : $"{alias}@microsoft.com";

        log?.Invoke($"Opening chat with {email}...");
        OpenChat(email);

        await Task.Delay(delayMs, ct);

        log?.Invoke("Focusing Teams window...");
        bool focused = FocusTeamsWindow();
        if (!focused)
            log?.Invoke("WARNING: Could not find Teams window. Attempting to send anyway...");

        await Task.Delay(500, ct);

        log?.Invoke("Pasting and sending message...");
        Application.Current.Dispatcher.Invoke(() => PasteAndSend(message));

        await Task.Delay(1000, ct);

        log?.Invoke($"Message sent to {email}.");
        return true;
    }

    /// <summary>
    /// Searches for a person by display name in Teams, opens new chat with first result, pastes and sends.
    /// Flow: Focus Teams → Ctrl+N (new chat) → type name → wait for search → Enter (select first) → paste → send
    /// </summary>
    private async Task<bool> SendViaSearch(string displayName, string message, int delayMs,
        CancellationToken ct, Action<string>? log = null)
    {
        log?.Invoke($"Searching Teams for \"{displayName}\"...");

        // Focus Teams first
        bool focused = FocusTeamsWindow();
        if (!focused)
        {
            log?.Invoke("WARNING: Could not find Teams window.");
            return false;
        }

        await Task.Delay(500, ct);

        // Ctrl+N to open a new chat
        _inputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_N);
        await Task.Delay(1500, ct);

        // Type the display name in the "To:" field
        log?.Invoke($"Typing \"{displayName}\" in search...");
        _inputSimulator.Keyboard.TextEntry(displayName);

        // Wait for search results to populate
        await Task.Delay(delayMs, ct);

        // Press Down arrow to select first result, then Enter to confirm
        log?.Invoke("Selecting first search result...");
        _inputSimulator.Keyboard.KeyPress(VirtualKeyCode.DOWN);
        await Task.Delay(300, ct);
        _inputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);

        // Wait for chat to open
        await Task.Delay(1500, ct);

        // Tab to move focus from "To:" field to the compose box
        log?.Invoke("Moving focus to compose box...");
        _inputSimulator.Keyboard.KeyPress(VirtualKeyCode.TAB);
        await Task.Delay(500, ct);

        // Now paste the message and send
        log?.Invoke("Pasting and sending message...");
        Application.Current.Dispatcher.Invoke(() => PasteAndSend(message));

        await Task.Delay(1000, ct);

        log?.Invoke($"Message sent to \"{displayName}\".");
        return true;
    }
}

using System.ComponentModel;

namespace MSSLTeamsMessenger.Models;

public class AliasEntry : INotifyPropertyChanged
{
    private bool _isSelected = true;
    private string _status = "";

    public string Alias { get; set; } = "";
    public string Email => $"{Alias}@microsoft.com";

    public bool IsSelected
    {
        get => _isSelected;
        set { _isSelected = value; OnPropertyChanged(nameof(IsSelected)); }
    }

    public string Status
    {
        get => _status;
        set { _status = value; OnPropertyChanged(nameof(Status)); }
    }

    public event PropertyChangedEventHandler? PropertyChanged;
    private void OnPropertyChanged(string name) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}

using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace SpreadsheetEditor.App.ViewModels;

/// <summary>The base of the view models: <see cref="INotifyPropertyChanged"/> by hand, without a framework (CLAUDE.md).</summary>
public abstract partial class ObservableObject : INotifyPropertyChanged
{
    public event PropertyChangedEventHandler? PropertyChanged;

    /// <summary>Raises <see cref="PropertyChanged"/> for <paramref name="propertyName"/>.</summary>
    protected void OnPropertyChanged([CallerMemberName] string? propertyName = null) =>
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}

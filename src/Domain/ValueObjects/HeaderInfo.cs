namespace SpreadsheetCLI.Domain.ValueObjects;

using System;

/// <summary>
/// Header information including name and row index
/// </summary>
public sealed class HeaderInfo
{
    public string Name { get; }
    public int RowIndex { get; }

    public HeaderInfo(string name, int rowIndex)
    {
        Name = name ?? "";
        RowIndex = rowIndex;
    }

    public override bool Equals(object? obj)
    {
        if (obj is not HeaderInfo other)
            return false;

        return Name == other.Name && RowIndex == other.RowIndex;
    }

    public override int GetHashCode()
    {
        return HashCode.Combine(Name, RowIndex);
    }
}
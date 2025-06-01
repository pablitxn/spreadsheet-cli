#!/usr/bin/env dotnet-script
#r "nuget: Aspose.Cells, 24.10.0"

using System;
using System.Linq;
using Aspose.Cells;

var filePath = "dataset/expanded_dataset_moved.xlsx";
var workbook = new Workbook(filePath);
var worksheet = workbook.Worksheets["Data"];

Console.WriteLine("=== ANÁLISIS DEL DATASET ===");
Console.WriteLine($"Archivo: {filePath}");
Console.WriteLine($"Hoja: {worksheet.Name}");
Console.WriteLine($"Filas totales: {worksheet.Cells.MaxDataRow + 1}");
Console.WriteLine($"Columnas totales: {worksheet.Cells.MaxDataColumn + 1}");

// Buscar la columna TotalBaseIncome
int totalBaseIncomeCol = -1;
for (int col = 0; col <= worksheet.Cells.MaxDataColumn; col++)
{
    var cell = worksheet.Cells[2, col]; // Row 3 (index 2) contiene headers
    if (cell.StringValue == "TotalBaseIncome")
    {
        totalBaseIncomeCol = col;
        break;
    }
}

if (totalBaseIncomeCol >= 0)
{
    Console.WriteLine($"\nColumna TotalBaseIncome encontrada en: {(char)('A' + totalBaseIncomeCol)}");
    
    var values = new List<double>();
    
    // Leer valores desde fila 4 (index 3) hasta el final
    for (int row = 3; row <= worksheet.Cells.MaxDataRow; row++)
    {
        var cell = worksheet.Cells[row, totalBaseIncomeCol];
        if (cell.Type == CellValueType.IsNumeric)
        {
            values.Add(cell.DoubleValue);
        }
    }
    
    Console.WriteLine($"\n=== ESTADÍSTICAS DE TotalBaseIncome ===");
    Console.WriteLine($"Cantidad de valores: {values.Count}");
    Console.WriteLine($"Suma total: {values.Sum():F2}");
    Console.WriteLine($"Promedio: {values.Average():F2}");
    Console.WriteLine($"Mínimo: {values.Min():F2}");
    Console.WriteLine($"Máximo: {values.Max():F2}");
    
    Console.WriteLine($"\nPrimeros 10 valores:");
    for (int i = 0; i < Math.Min(10, values.Count); i++)
    {
        Console.WriteLine($"  Fila {i + 4}: {values[i]:F2}");
    }
    
    Console.WriteLine($"\nÚltimos 10 valores:");
    for (int i = Math.Max(0, values.Count - 10); i < values.Count; i++)
    {
        Console.WriteLine($"  Fila {i + 4}: {values[i]:F2}");
    }
}
else
{
    Console.WriteLine("ERROR: No se encontró la columna TotalBaseIncome");
}
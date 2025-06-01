using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main(string[] args)
    {
        var groundTruthFile = args.Length > 0 ? args[0] : "./scripts/dataset/ground_truth_expanded_dataset_moved.xlsx";
        
        if (!File.Exists(groundTruthFile))
        {
            Console.Error.WriteLine($"Error: Ground truth file not found: {groundTruthFile}");
            Environment.Exit(1);
        }
        
        try
        {
            var workbook = new Workbook(groundTruthFile);
            var worksheet = workbook.Worksheets[0];
            
            int questionCol = -1;
            int answerCol = -1;
            
            // Find Question and Answer columns
            for (int col = 0; col < worksheet.Cells.MaxColumn + 1; col++)
            {
                var cell = worksheet.Cells[0, col];
                var value = cell.StringValue?.Trim();
                
                if (value == "Question")
                    questionCol = col;
                else if (value == "Answer")
                    answerCol = col;
            }
            
            if (questionCol == -1 || answerCol == -1)
            {
                Console.Error.WriteLine("Error: Could not find 'Question' or 'Answer' columns");
                Environment.Exit(1);
            }
            
            // Extract questions and answers
            for (int row = 1; row <= worksheet.Cells.MaxRow; row++)
            {
                var questionCell = worksheet.Cells[row, questionCol];
                var answerCell = worksheet.Cells[row, answerCol];
                
                var question = questionCell.StringValue;
                var answer = answerCell.StringValue;
                
                if (!string.IsNullOrWhiteSpace(question) && !string.IsNullOrWhiteSpace(answer))
                {
                    Console.WriteLine($"{question}|||{answer}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error reading Excel file: {ex.Message}");
            Environment.Exit(1);
        }
    }
}
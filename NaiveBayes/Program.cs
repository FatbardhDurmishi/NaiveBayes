using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

//Open Excel file
List<string[]> records = new List<string[]>();

using (SpreadsheetDocument document = SpreadsheetDocument.Open("C:\\Riinvest\\Viti III\\Semestri 6\\AI\\Detyra\\NaiveBayes\\BuyComputerDataSet.xlsx", false))
{
    // Get the first worksheet
    WorksheetPart worksheetPart = document.WorkbookPart.WorksheetParts.First();
    Worksheet worksheet = worksheetPart.Worksheet;

    // Get the shared string table
    SharedStringTablePart sharedStringTablePart = document.WorkbookPart.SharedStringTablePart;
    SharedStringTable sharedStringTable = sharedStringTablePart.SharedStringTable;

    // Read the rows
    foreach (Row row in worksheet.GetFirstChild<SheetData>().Elements<Row>())
    {
        List<string> columns = new List<string>();
        foreach (Cell cell in row.Elements<Cell>())
        {
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                // Cell value is stored as a shared string
                int index = int.Parse(cell.InnerText);
                columns.Add(sharedStringTable.ElementAt(index).InnerText);
            }
            else
            {
                // Cell value is stored as a simple value
                columns.Add(cell.InnerText);
            }
        }
        records.Add(columns.ToArray());
    }
}

// Split records into training and test data
int splitIndex = (int)(0.9 * records.Count);
List<string[]> trainingData = records.GetRange(0, splitIndex);
List<string[]> testData = records.GetRange(splitIndex, records.Count - splitIndex);

// Train model
Dictionary<string, Dictionary<string, int>> wordCountsByLabel = new Dictionary<string, Dictionary<string, int>>();
Dictionary<string, int> labelCounts = new Dictionary<string, int>();
foreach (string[] record in trainingData)
{
    string label = record[record.Length - 1];
    if (!wordCountsByLabel.ContainsKey(label))
    {
        wordCountsByLabel[label] = new Dictionary<string, int>();
        labelCounts[label] = 0;
    }
    for (int i = 0; i < record.Length - 1; i++)
    {
        string word = record[i];
        if (!wordCountsByLabel[label].ContainsKey(word))
        {
            wordCountsByLabel[label][word] = 0;
        }
        wordCountsByLabel[label][word]++;
    }
    labelCounts[label]++;
}


// Test model
//int correctPredictions = 0;
int totalPredictions = testData.Count;

//List<string> predictions = new List<string>();
//Give Predictions
foreach (string[] record in testData)
{
    //string currentLabel;
    //string actualLabel = record[record.Length - 1];
    Dictionary<string, double> scoresByLabel = new Dictionary<string, double>();
    foreach (string label in labelCounts.Keys)
    {
        double score=1;
        for (int i = 0; i < record.Length - 1; i++)
        {
            string word = record[i];

            if (wordCountsByLabel.TryGetValue(label, out Dictionary<string, int> inner))
            {
                if (inner.TryGetValue(word, out int value))
                {
                    score *= (double)value/labelCounts[label];
                }
            }
        }
        score*=trainingData.Count;
        scoresByLabel[label] = score;
    }
    record[record.Length - 1] = scoresByLabel.Aggregate((x, y) => x.Value > y.Value ? x : y).Key;
    for(int i = 0; i < record.Length; i++)
    {
        Console.Write($"{record[i]}\t");

    }
    Console.WriteLine();

    //predictions.Add(scoresByLabel.Aggregate((x, y) => x.Value > y.Value ? x : y).Key);
    //string predictedLabel = scoresByLabel.Aggregate((x, y) => x.Value > y.Value ? x : y).Key;
    //if (predictedLabel == currentLabel)
    //{
    //    correctPredictions++;
    //}
}
//foreach(string prediction in predictions)
//{
//    Console.WriteLine(prediction);
//}

// Print accuracy
//double accuracy = (double)correctPredictions / totalPredictions;
//Console.WriteLine($"Accuracy: {accuracy:P}");
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml; // Add this line
using DocumentFormat.OpenXml.Office2013.ExcelAc;


namespace ExcelAnalyzerApp
{
    public partial class MainWindow : Window
    {
        private WorkbookPart? _workbookPart; // Declare as nullable
        private SpreadsheetDocument? _spreadsheetDocument; // Declare as nullable
        private bool _isComboBoxPopulated = false; // Flag to track ComboBox population

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xlsm",
                Title = "Select an Excel File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                AnalyzeExcelFile(filePath);
            }
        }

        private void AnalyzeExcelFile(string filePath)
        {
            // Dispose the previous document if it's open
            if (_spreadsheetDocument != null)
            {
                _spreadsheetDocument.Dispose(); // Use Dispose instead of Close
                _spreadsheetDocument = null;
            }

            ResultsListBox.Items.Clear();
            WorksheetComboBox.Items.Clear();
            _isComboBoxPopulated = false; // Reset the flag

            try
            {
                // Open the Excel file and keep the document open
                _spreadsheetDocument = SpreadsheetDocument.Open(filePath, false);
                _workbookPart = _spreadsheetDocument.WorkbookPart;

                if (_workbookPart == null)
                {
                    ResultsListBox.Items.Add("Error: Workbook part is null.");
                    return;
                }

                // Populate the ComboBox with worksheet names
                foreach (var worksheetPart in _workbookPart.WorksheetParts)
                {
                    Sheet sheet = _workbookPart.Workbook.Descendants<Sheet>()
                        .FirstOrDefault(s => s.Id.Value == _workbookPart.GetIdOfPart(worksheetPart));

                    if (sheet != null)
                    {
                        WorksheetComboBox.Items.Add(sheet.Name);
                    }
                }

                // Enable the SelectionChanged event after populating the ComboBox
                _isComboBoxPopulated = true;

                // Select the first worksheet by default
                if (WorksheetComboBox.Items.Count > 0)
                {
                    WorksheetComboBox.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                ResultsListBox.Items.Add($"Error loading Excel file: {ex.Message}");
            }
        }

        private void WorksheetComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Only trigger if the ComboBox is fully populated
            if (_isComboBoxPopulated && WorksheetComboBox.SelectedItem != null)
            {
                string selectedWorksheet = WorksheetComboBox.SelectedItem.ToString();
                AnalyzeWorksheet(selectedWorksheet);
            }
        }


private void AnalyzeWorksheet(string worksheetName)
{
    ResultsListBox.Items.Clear();

    if (_workbookPart == null)
    {
        ResultsListBox.Items.Add("No workbook loaded.");
        return;
    }

    try
    {
        Sheet sheet = _workbookPart.Workbook.Descendants<Sheet>()
            .FirstOrDefault(s => s.Name == worksheetName);

        if (sheet?.Id?.Value == null)
        {
            ResultsListBox.Items.Add($"Worksheet '{worksheetName}' not found.");
            return;
        }

        WorksheetPart worksheetPart = (WorksheetPart)_workbookPart.GetPartById(sheet.Id.Value);

        // Only call DetectFormulas (formatting is now merged into this method)
        DetectFormulas(worksheetPart, _workbookPart);
        DetectCharts(worksheetPart);
        DetectPivotTables(_workbookPart);
    }
    catch (Exception ex)
    {
        ResultsListBox.Items.Add($"Error analyzing worksheet '{worksheetName}': {ex.Message}");
    }
}


private void DetectFormulas(WorksheetPart worksheetPart, WorkbookPart workbookPart)
{
    try
    {
        var cellsWithFormulas = worksheetPart.Worksheet.Descendants<Cell>()
            .Where(c => c.CellFormula != null);

        ResultsListBox.Items.Add($"Found {cellsWithFormulas.Count()} cells with formulas:");

        foreach (var cell in cellsWithFormulas)
        {
            string cellReference = cell.CellReference?.Value ?? "Unknown";
            string formula = cell.CellFormula?.Text ?? "No Formula";
            string calculatedValue = GetCellValue(cell, workbookPart);

            ResultsListBox.Items.Add($"- Cell {cellReference}:");
            ResultsListBox.Items.Add($"  Formula: {formula}");
            ResultsListBox.Items.Add($"  Calculated Value: {calculatedValue}");

            // Formatting check ONLY for formula cells
            if (cell.StyleIndex != null)
            {
                uint styleIndex = cell.StyleIndex.Value;
                var cellFormat = workbookPart.WorkbookStylesPart?.Stylesheet
                    .CellFormats?.Elements<CellFormat>()
                    .ElementAt((int)styleIndex);

                if (cellFormat != null)
                {
                    // Number Format
                    if (cellFormat.NumberFormatId != null)
                    {
                        var numberFormat = workbookPart.WorkbookStylesPart?.Stylesheet
                            .NumberingFormats?.Elements<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat>()
                            .FirstOrDefault(nf => nf.NumberFormatId == cellFormat.NumberFormatId);
                        
                        ResultsListBox.Items.Add($"  Number Format: {numberFormat?.FormatCode ?? "General"}");
                    }

                    // Font
                    if (cellFormat.FontId != null)
                    {
                        var font = workbookPart.WorkbookStylesPart?.Stylesheet
                            .Fonts?.Elements<Font>()
                            .ElementAt((int)cellFormat.FontId.Value);
                        
                        ResultsListBox.Items.Add($"  Font: {font?.FontName?.Val ?? "Calibri"}, Size: {font?.FontSize?.Val ?? 11}");
                    }

                    // Fill Color
                    if (cellFormat.FillId != null)
                    {
                        var fill = workbookPart.WorkbookStylesPart?.Stylesheet
                            .Fills?.Elements<Fill>()
                            .ElementAt((int)cellFormat.FillId.Value);
                        
                        ResultsListBox.Items.Add($"  Fill Color: {fill?.PatternFill?.ForegroundColor?.Rgb ?? "None"}");
                    }

                    // Borders
                    if (cellFormat.BorderId != null)
                    {
                        var border = workbookPart.WorkbookStylesPart?.Stylesheet
                            .Borders?.Elements<DocumentFormat.OpenXml.Spreadsheet.Border>()
                            .ElementAt((int)cellFormat.BorderId.Value);
                        
                        ResultsListBox.Items.Add($"  Borders: Top={border?.TopBorder != null}, Bottom={border?.BottomBorder != null}");
                    }
                }
            }
        }
    }
    catch (Exception ex)
    {
        ResultsListBox.Items.Add($"Error detecting formulas: {ex.Message}");
    }
}



private void DetectFormatting(WorksheetPart worksheetPart, WorkbookPart workbookPart)
{
    try
    {
        var styledCells = worksheetPart.Worksheet.Descendants<Cell>()
            .Where(c => c.StyleIndex != null);

        ResultsListBox.Items.Add($"Found {styledCells.Count()} cells with custom formatting:");

        foreach (var cell in styledCells)
        {
            string cellReference = cell.CellReference?.Value ?? "Unknown";
            uint styleIndex = cell.StyleIndex?.Value ?? 0;

            var cellFormat = workbookPart.WorkbookStylesPart?.Stylesheet
                .CellFormats?.Elements<CellFormat>()
                .ElementAt((int)styleIndex);

            if (cellFormat != null)
            {
                ResultsListBox.Items.Add($"- Cell {cellReference}:");

                // Fix: Fully qualify NumberingFormat
                if (cellFormat.NumberFormatId != null)
                {
                    var numberFormat = workbookPart.WorkbookStylesPart?.Stylesheet
                        .NumberingFormats?.Elements<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat>()
                        .FirstOrDefault(nf => nf.NumberFormatId == cellFormat.NumberFormatId);

                    ResultsListBox.Items.Add($"  Number Format: {numberFormat?.FormatCode ?? "General"}");
                }

                // Fix: Fully qualify Border
                if (cellFormat.BorderId != null)
                    {
                        var border = workbookPart.WorkbookStylesPart?.Stylesheet
                            .Borders?.Elements<DocumentFormat.OpenXml.Spreadsheet.Border>()
                            .ElementAt((int)cellFormat.BorderId.Value);

                        if (border != null)
                        {
                            // Check if TopBorder and BottomBorder exist
                            bool hasTop = border.TopBorder != null;
                            bool hasBottom = border.BottomBorder != null;

                            ResultsListBox.Items.Add($"  Borders: Top={hasTop}, Bottom={hasBottom}");
                        }
                    }
            }
        }
    }
    catch (Exception ex)
    {
        ResultsListBox.Items.Add($"Error detecting formatting: {ex.Message}");
    }
}


private void DetectCharts(WorksheetPart worksheetPart)
{
    try
    {
        int chartCount = 0;

        // Ensure the DrawingsPart exists
        if (worksheetPart.DrawingsPart != null)
        {
            var worksheetDrawing = worksheetPart.DrawingsPart.WorksheetDrawing;

            if (worksheetDrawing != null)
            {
                // Iterate over all TwoCellAnchor elements
                foreach (var twoCellAnchor in worksheetDrawing.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>())
                {
                    // Check for regular charts
                    if (twoCellAnchor.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().Any())
                    {
                        string chartPosition = GetCellPositionFromTwoCellAnchor(twoCellAnchor);
                        ResultsListBox.Items.Add($"Regular Chart {++chartCount} located at {chartPosition}.");

                        // Extract elements of the detected regular chart
                        DetectChartElementsForAnchor(twoCellAnchor, worksheetPart, chartCount, chartPosition);
                    }

                    // Check for extended charts (cx:chart)
                    if (twoCellAnchor.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>()
                        .Any(e => e.LocalName == "chart" && e.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2014/chartex"))
                    {
                        string chartPosition = GetCellPositionFromTwoCellAnchor(twoCellAnchor);
                        ResultsListBox.Items.Add($"Extended Chart {++chartCount} located at {chartPosition}.");

                        // Extract elements for the detected extended chart
                        var chartReference = twoCellAnchor.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>()
                            .FirstOrDefault(e => e.LocalName == "chart" && e.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2014/chartex");

                        if (chartReference != null)
                        {
                            string chartId = chartReference.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value;

                            var chartExPart = worksheetPart.DrawingsPart.GetPartById(chartId);
                            if (chartExPart != null)
                            {
                                DetectChartElementsForExtendedChart(chartExPart, chartId, worksheetPart);
                            }
                        }
                    }
                }
            }
            else
            {
                ResultsListBox.Items.Add("No WorksheetDrawing found in the DrawingsPart.");
            }
        }
        else
        {
            ResultsListBox.Items.Add("No DrawingsPart found in the worksheet.");
        }

        if (chartCount == 0)
        {
            ResultsListBox.Items.Add("No charts found in the worksheet.");
        }
        else
        {
            ResultsListBox.Items.Add($"Found {chartCount} charts in total.");
        }
    }
    catch (Exception ex)
    {
        ResultsListBox.Items.Add($"Error detecting charts: {ex.Message}");
    }
}


private void DetectChartElementsForAnchor(DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor twoCellAnchor, WorksheetPart worksheetPart, int chartCount, string chartPosition)
{
    var chartReference = twoCellAnchor.Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>().FirstOrDefault();

    if (chartReference != null)
    {
        string chartId = chartReference.Id;

        if (worksheetPart.DrawingsPart.ChartParts.Any(cp => worksheetPart.DrawingsPart.GetIdOfPart(cp) == chartId))
        {
            var chartPart = worksheetPart.DrawingsPart.GetPartById(chartId) as ChartPart;

            if (chartPart != null)
            {
                string chartName = GetChartName(twoCellAnchor);
                string chartTitle = GetChartTitle(chartPart);
                string chartType = GetChartType(chartPart);
                string chartAxes = GetChartAxes(chartPart);
                bool hasLegend = HasChartLegend(chartPart);
                string chartData = GetRegularChartData(chartPart);

                ResultsListBox.Items.Add($"Chart {chartCount} Details:");
                ResultsListBox.Items.Add($"- Name: {chartName}");
                ResultsListBox.Items.Add($"- Position: {chartPosition}");
                ResultsListBox.Items.Add($"- Title: {chartTitle}");
                ResultsListBox.Items.Add($"- Type: {chartType}");
                ResultsListBox.Items.Add($"- Axes: {chartAxes}");
                ResultsListBox.Items.Add($"- Legend: {(hasLegend ? "Yes" : "No")}");
                ResultsListBox.Items.Add($"- Data Source: {chartData}");
            }
        }
    }
}


private string GetChartName(DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor twoCellAnchor)
{
    var graphicFrame = twoCellAnchor.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>().FirstOrDefault();
    if (graphicFrame != null && graphicFrame.NonVisualGraphicFrameProperties != null)
    {
        return graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties?.Name ?? "Unnamed Chart";
    }
    return "Unnamed Chart";
}

private string GetChartTitle(ChartPart chartPart)
{
    var title = chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Title>().FirstOrDefault();
    if (title?.ChartText?.RichText != null)
    {
        var textElement = title.ChartText.RichText.Descendants<DocumentFormat.OpenXml.Drawing.Run>().FirstOrDefault()?.Text;
        return textElement?.Text ?? "No Title";
    }
    return "No Title";
}


private string GetChartType(ChartPart chartPart)
{
    try
    {
        // Get the PlotArea element from the ChartSpace
        var plotArea = chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().FirstOrDefault();
        if (plotArea == null)
        {
            return "Unknown Chart Type";
        }

        // Find the first chart element inside PlotArea
        var chartElement = plotArea.Elements().FirstOrDefault(e => e.LocalName.EndsWith("Chart"));
        if (chartElement != null)
        {
            // Return the dynamic LocalName of the chart element
            return chartElement.LocalName;
        }

        return "Unknown Chart Type";
    }
    catch (Exception ex)
    {
        // Handle errors gracefully
        return $"Error detecting chart type: {ex.Message}";
    }
}



private string GetChartAxes(ChartPart chartPart)
{
    try
    {
        var axisTitles = new List<string>();

        // X-Axis (Primary Horizontal Axis)
        var xAxis = chartPart.ChartSpace.Descendants<CategoryAxis>().FirstOrDefault();
        if (xAxis != null)
        {
            var xAxisTitle = xAxis.Descendants<Title>().FirstOrDefault();
            var xAxisText = xAxisTitle?.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault()?.Text;
            axisTitles.Add($"X-Axis: {xAxisText ?? "No Title"}");
        }

        // Y-Axis (Primary Vertical Axis)
        var yAxis = chartPart.ChartSpace.Descendants<ValueAxis>().FirstOrDefault();
        if (yAxis != null)
        {
            var yAxisTitle = yAxis.Descendants<Title>().FirstOrDefault();
            var yAxisText = yAxisTitle?.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault()?.Text;
            axisTitles.Add($"Y-Axis: {yAxisText ?? "No Title"}");
        }

        // Secondary Y-Axis (Secondary Vertical Axis)
        var secondaryYAxis = chartPart.ChartSpace.Descendants<ValueAxis>()
            .Skip(1) // Assume secondary Y-Axis comes second in the list of ValueAxis
            .FirstOrDefault();
        if (secondaryYAxis != null)
        {
            var secondaryYAxisTitle = secondaryYAxis.Descendants<Title>().FirstOrDefault();
            var secondaryYAxisText = secondaryYAxisTitle?.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault()?.Text;
            axisTitles.Add($"Secondary Y-Axis: {secondaryYAxisText ?? "No Title"}");
        }

        // Z-Axis (3D Charts)
        var zAxis = chartPart.ChartSpace.Descendants<SeriesAxis>().FirstOrDefault();
        if (zAxis != null)
        {
            var zAxisTitle = zAxis.Descendants<Title>().FirstOrDefault();
            var zAxisText = zAxisTitle?.Descendants<DocumentFormat.OpenXml.Drawing.Text>().FirstOrDefault()?.Text;
            axisTitles.Add($"Z-Axis: {zAxisText ?? "No Title"}");
        }

        return axisTitles.Count > 0 ? string.Join(", ", axisTitles) : "No Axes Titles Found";
    }
    catch (Exception ex)
    {
        return $"Error detecting axis titles: {ex.Message}";
    }
}


private bool HasChartLegend(ChartPart chartPart)
{
    return chartPart.ChartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Legend>().Any();
}


private string GetExtendedChartTitle(OpenXmlPart chartExPart)
{
    // Locate the <cx:title> or any title-equivalent element
    var titleElement = chartExPart.RootElement.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>()
        .FirstOrDefault(p => p.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Any());

    if (titleElement != null)
    {
        return titleElement.InnerText;
    }

    return "No Title";
}



private List<string> GetExtendedChartAxesTitles(OpenXmlPart chartExPart)
{
    var axisTitles = new List<string>();

    try
    {
        // Locate all axis elements with titles in the extended chart
        var axisElements = chartExPart.RootElement
            .Descendants()
            .Where(el => el.LocalName == "axis" && el.Descendants().Any(d => d.LocalName == "title"));

        foreach (var axis in axisElements)
        {
            var titleElement = axis.Descendants()
                .FirstOrDefault(d => d.LocalName == "title");

            if (titleElement != null)
            {
                var textElement = titleElement.Descendants()
                    .FirstOrDefault(t => t.LocalName == "t"); // Looks for <a:t> text tag
                if (textElement != null)
                {
                    axisTitles.Add(textElement.InnerText);
                }
            }
        }
    }
    catch (Exception ex)
    {
        axisTitles.Add($"Error extracting axis titles: {ex.Message}");
    }

    return axisTitles.Count > 0 ? axisTitles : new List<string> { "No Axis Titles Found" };
}


private string GetExtendedChartType(OpenXmlPart chartExPart)
{
    try
    {
        // Locate the first <cx:series> element and extract its "layoutId" attribute
        var layoutIdElement = chartExPart.RootElement.Descendants<DocumentFormat.OpenXml.OpenXmlElement>()
            .FirstOrDefault(e => e.LocalName == "series" && e.HasAttributes);

        if (layoutIdElement != null)
        {
            var layoutId = layoutIdElement.GetAttribute("layoutId", "").Value;
            return !string.IsNullOrEmpty(layoutId) ? layoutId : "Unknown Chart Type";
        }

        return "Unknown Chart Type";
    }
    catch (Exception ex)
    {
        return $"Error detecting chart type: {ex.Message}";
    }
}






private bool IsLegendPresentInExtendedChart(OpenXmlPart chartExPart)
{
    // Check if any <legend> or equivalent exists in the extended chart XML
    var legendElement = chartExPart.RootElement
        .Descendants<DocumentFormat.OpenXml.OpenXmlElement>()
        .FirstOrDefault(e => e.LocalName == "legend");

    return legendElement != null;
}



private string GetExtendedChartName(WorksheetPart worksheetPart, string chartId)
{
    try
    {
        // Locate the GraphicFrame with the matching r:id in drawing1.xml
        var graphicFrame = worksheetPart.DrawingsPart.WorksheetDrawing
            .Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>()
            .FirstOrDefault(gf =>
            {
                var chartReference = gf.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties>()
                    .FirstOrDefault()?.NonVisualDrawingProperties;
                if (chartReference != null)
                {
                    var chartElement = gf.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>()
                        .FirstOrDefault(e => e.LocalName == "chart" && e.NamespaceUri == "http://schemas.microsoft.com/office/drawing/2014/chartex");
                    if (chartElement != null)
                    {
                        return chartElement.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value == chartId;
                    }
                }
                return false;
            });

        // Extract the name attribute from the matching GraphicFrame
        if (graphicFrame != null)
        {
            var nonVisualProps = graphicFrame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;
            if (nonVisualProps != null)
            {
                return nonVisualProps.Name ?? "Unnamed Extended Chart";
            }
        }

        return "Unnamed Extended Chart";
    }
    catch (Exception ex)
    {
        return $"Error detecting extended chart name: {ex.Message}";
    }
}


private List<string> GetExtendedChartData(OpenXmlPart chartExPart, WorkbookPart workbookPart)
{
    var chartDataLocations = new List<string>();

    try
    {
        // Locate the <cx:numDim> element in the chartEx XML
        var numDimElement = chartExPart.RootElement.Descendants()
            .FirstOrDefault(e => e.LocalName == "numDim" && e.HasAttributes);

        if (numDimElement != null)
        {
            // Find the formula reference (e.g., "_xlchart.v2.0") in the <cx:f> tag
            var formulaReference = numDimElement.Descendants()
                .FirstOrDefault(e => e.LocalName == "f")?.InnerText;

            if (!string.IsNullOrEmpty(formulaReference))
            {
                // Locate the corresponding defined name in workbook.xml
                var definedName = workbookPart.Workbook.DefinedNames
                    .Descendants<DocumentFormat.OpenXml.Spreadsheet.DefinedName>()
                    .FirstOrDefault(dn => dn.Name == formulaReference);

                if (definedName != null)
                {
                    // Extract the range (e.g., "SheetName!A1:B10")
                    chartDataLocations.Add(definedName.InnerText);
                }
                else
                {
                    chartDataLocations.Add($"Defined name '{formulaReference}' not found in workbook.");
                }
            }
            else
            {
                chartDataLocations.Add("No formula reference found for chart data.");
            }
        }
        else
        {
            chartDataLocations.Add("No <cx:numDim> element found in chartEx XML.");
        }
    }
    catch (Exception ex)
    {
        chartDataLocations.Add($"Error retrieving extended chart data: {ex.Message}");
    }

    return chartDataLocations;
}


// Helper function to parse a cell range (e.g., "D1029:D1037") and return individual cell references


private string GetRegularChartData(ChartPart chartPart)
{
    try
    {
        // For regular charts, look for <c:cat> (categories) and <c:val> (values)
        var categoryRangeElement = chartPart.ChartSpace.Descendants()
            .FirstOrDefault(e => e.LocalName == "f" && e.Parent.LocalName == "strRef" && e.Parent.Parent.LocalName == "cat");

        var valueRangeElement = chartPart.ChartSpace.Descendants()
            .FirstOrDefault(e => e.LocalName == "f" && e.Parent.LocalName == "numRef" && e.Parent.Parent.LocalName == "val");

        if (categoryRangeElement != null && valueRangeElement != null)
        {
            // Extract the ranges
            string categoryRange = categoryRangeElement.InnerText; // Example: QuestionPizza!$C$876:$C$885
            string valueRange = valueRangeElement.InnerText;       // Example: QuestionPizza!$D$876:$D$885

            // Combine the ranges
            string sheetName = categoryRange.Split('!')[0];
            string categoryStart = categoryRange.Split('!')[1].Split(':')[0];
            string categoryEnd = categoryRange.Split('!')[1].Split(':')[1];
            string valueStart = valueRange.Split('!')[1].Split(':')[0];
            string valueEnd = valueRange.Split('!')[1].Split(':')[1];

            // Return combined range
            return $"{sheetName}!{categoryStart}:{valueEnd}";
        }

        // Handle pivot charts: Check for <c:pivotSource> and associated ranges
        var pivotSourceElement = chartPart.ChartSpace.Descendants()
            .FirstOrDefault(e => e.LocalName == "pivotSource");

        if (pivotSourceElement != null)
        {
            // Extract the pivot table name
            var pivotNameElement = pivotSourceElement.Descendants()
                .FirstOrDefault(e => e.LocalName == "name");

            if (pivotNameElement != null)
            {
                string pivotName = pivotNameElement.InnerText; // Example: QuestionPizza!PivotTable1

                // Locate category and value ranges for the pivot chart
                var pivotCategoryRange = chartPart.ChartSpace.Descendants()
                    .FirstOrDefault(e => e.LocalName == "f" && e.Parent.LocalName == "strRef" && e.Parent.Parent.LocalName == "cat");

                var pivotValueRange = chartPart.ChartSpace.Descendants()
                    .FirstOrDefault(e => e.LocalName == "f" && e.Parent.LocalName == "numRef" && e.Parent.Parent.LocalName == "val");

                if (pivotCategoryRange != null && pivotValueRange != null)
                {
                    // Extract the ranges
                    string categoryRange = pivotCategoryRange.InnerText; // Example: QuestionPizza!$C$876:$C$885
                    string valueRange = pivotValueRange.InnerText;       // Example: QuestionPizza!$D$876:$D$885

                    // Combine the ranges
                    string sheetName = categoryRange.Split('!')[0];
                    string categoryStart = categoryRange.Split('!')[1].Split(':')[0];
                    string categoryEnd = categoryRange.Split('!')[1].Split(':')[1];
                    string valueStart = valueRange.Split('!')[1].Split(':')[0];
                    string valueEnd = valueRange.Split('!')[1].Split(':')[1];

                    // Return combined range
                    return $"{sheetName}!{categoryStart}:{valueEnd}";
                }

                // If ranges not found, return just the pivot name
                return pivotName;
            }
        }

        // Fallback: If no data source found
        return "Data source not found.";
    }
    catch (Exception ex)
    {
        return $"Error retrieving data source: {ex.Message}";
    }
}


private void DetectChartElementsForExtendedChart(OpenXmlPart chartExPart, string chartId, WorksheetPart worksheetPart)
{
    try
    {
        string title = "No Title";
        List<string> axesTitles = new List<string>();
        string chartType = "Unknown";
        bool hasLegend = false;
        string chartName = "Unnamed Chart";
        List<string> chartData = new List<string>();

        // Attempt to retrieve each property
        try { title = GetExtendedChartTitle(chartExPart); } catch (Exception ex) { ResultsListBox.Items.Add($"Error detecting title: {ex.Message}"); }
        try { axesTitles = GetExtendedChartAxesTitles(chartExPart); } catch (Exception ex) { ResultsListBox.Items.Add($"Error detecting axes titles: {ex.Message}"); }
        try { chartType = GetExtendedChartType(chartExPart); } catch (Exception ex) { ResultsListBox.Items.Add($"Error detecting chart type: {ex.Message}"); }
        try { hasLegend = IsLegendPresentInExtendedChart(chartExPart); } catch (Exception ex) { ResultsListBox.Items.Add($"Error detecting legend: {ex.Message}"); }
        try { chartName = GetExtendedChartName(worksheetPart, chartId); } catch (Exception ex) { ResultsListBox.Items.Add($"Error detecting chart name: {ex.Message}"); }
        try { chartData = GetExtendedChartData(chartExPart, _workbookPart); } catch (Exception ex) { ResultsListBox.Items.Add($"Error retrieving chart data: {ex.Message}"); }

        // Output detected details
        ResultsListBox.Items.Add($"Chart Name: {chartName}");
        ResultsListBox.Items.Add($"Chart Title: {title}");
        ResultsListBox.Items.Add($"Chart Type: {chartType}");
        ResultsListBox.Items.Add($"Legend Present: {(hasLegend ? "Yes" : "No")}");
        ResultsListBox.Items.Add($"Axes Titles: {string.Join(", ", axesTitles)}");
        ResultsListBox.Items.Add($"Chart Data: {string.Join(", ", chartData)}");
    }
    catch (Exception ex)
    {
        ResultsListBox.Items.Add($"Error detecting chart elements: {ex.Message}");
    }
}




private string GetCellPositionFromTwoCellAnchor(DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor twoCellAnchor)
{
    // Get starting position (FromMarker)
    var fromMarker = twoCellAnchor.FromMarker;
    string fromCell = GetExcelCellReference(
        int.Parse(fromMarker.ColumnId.Text),
        int.Parse(fromMarker.RowId.Text)
    );

    // Get ending position (ToMarker)
    var toMarker = twoCellAnchor.ToMarker;
    string toCell = GetExcelCellReference(
        int.Parse(toMarker.ColumnId.Text),
        int.Parse(toMarker.RowId.Text)
    );

    // Return as Excel-style range
    return $"{fromCell}:{toCell}";
}

private string GetExcelCellReference(int columnIndex, int rowIndex)
{
    // Convert column index to Excel-style letters (0-based to 1-based)
    string columnLetter = GetColumnLetters(columnIndex);
    int excelRow = rowIndex + 1; // Convert 0-based row to 1-based

    return $"{columnLetter}{excelRow}";
}

private string GetColumnLetters(int columnIndex)
{
    string columnLetters = string.Empty;

    while (columnIndex >= 0)
    {
        columnLetters = (char)('A' + (columnIndex % 26)) + columnLetters;
        columnIndex = (columnIndex / 26) - 1;
    }

    return columnLetters;
}


private void DetectPivotTables(WorkbookPart workbookPart)
{
    try
    {
        int pivotTableCount = 0;

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            foreach (var pivotTablePart in worksheetPart.PivotTableParts)
            {
                var pivotTableDef = pivotTablePart.PivotTableDefinition;
                string pivotTableName = pivotTableDef?.Name?.Value ?? "Unnamed Pivot Table";
                string location = pivotTableDef?.Location?.Reference?.Value ?? "Unknown Location";

                ResultsListBox.Items.Add($"Pivot Table '{pivotTableName}' found at {location}");

                // Pass workbookPart as the third parameter
                ExtractPivotTableValues(worksheetPart, location, workbookPart);
                pivotTableCount++;
            }
        }

        ResultsListBox.Items.Add($"Total Pivot Tables Found: {pivotTableCount}");
    }
    catch (Exception ex)
    {
        ResultsListBox.Items.Add($"Error detecting pivot tables: {ex.Message}");
    }
}


private void ExtractPivotTableValues(WorksheetPart worksheetPart, string cellRange, WorkbookPart workbookPart)
{
    try
    {
        Worksheet worksheet = worksheetPart.Worksheet;
        SheetData sheetData = worksheet.GetFirstChild<SheetData>();

        // Parse cell range (e.g., "A1:D10")
        string[] rangeParts = cellRange.Split(':');
        if (rangeParts.Length != 2)
        {
            ResultsListBox.Items.Add($"Invalid range format: {cellRange}");
            return;
        }

        string startCell = rangeParts[0];
        string endCell = rangeParts[1];

        // Get cells in range
        var cellsInRange = sheetData.Descendants<Cell>()
            .Where(c => IsCellInRange(c.CellReference?.Value, startCell, endCell))
            .OrderBy(c => GetColumnIndex(c.CellReference.Value))
            .ThenBy(c => GetRowIndex(c.CellReference.Value));

        ResultsListBox.Items.Add($"Pivot Table Values:");

        // Read values
        foreach (var cell in cellsInRange)
        {
            string cellValue = GetCellValue(cell, workbookPart); // Now uses passed workbookPart
            ResultsListBox.Items.Add($"- {cell.CellReference}: {cellValue}");
        }
    }
    catch (Exception ex)
    {
        ResultsListBox.Items.Add($"Error extracting pivot table values: {ex.Message}");
    }
}


// Helper: Check if a cell is within the specified range
private bool IsCellInRange(string cellReference, string startCell, string endCell)
{
    if (string.IsNullOrEmpty(cellReference)) return false;

    int currentCol = GetColumnIndex(cellReference);
    int currentRow = GetRowIndex(cellReference);
    
    int startCol = GetColumnIndex(startCell);
    int startRow = GetRowIndex(startCell);
    
    int endCol = GetColumnIndex(endCell);
    int endRow = GetRowIndex(endCell);

    return currentCol >= startCol && 
           currentCol <= endCol && 
           currentRow >= startRow && 
           currentRow <= endRow;
}

// Helper: Get cell value (reuse from formula detection logic)

private string GetCellValue(Cell cell, WorkbookPart workbookPart)
{
    if (cell.DataType?.Value == CellValues.SharedString)
    {
        // Use workbookPart from parameter
        var sharedStringTable = workbookPart?.SharedStringTablePart?.SharedStringTable;
        if (sharedStringTable != null && int.TryParse(cell.InnerText, out int index))
        {
            return sharedStringTable.ElementAt(index).InnerText;
        }
    }
    return cell.CellValue?.Text ?? "Empty";
}

// Helper: Convert column letters to index (e.g., "AA" -> 27)
private int GetColumnIndex(string cellReference)
{
    string columnLetters = new string(cellReference.Where(char.IsLetter).ToArray());
    int index = 0;
    foreach (char c in columnLetters)
    {
        index = index * 26 + (c - 'A' + 1);
    }
    return index - 1; // Zero-based
}

// Helper: Get row index from cell reference
private int GetRowIndex(string cellReference)
{
    string rowNumber = new string(cellReference.Where(char.IsDigit).ToArray());
    return int.TryParse(rowNumber, out int result) ? result - 1 : -1; // Zero-based
}

// Modified ExtractPivotTableData()
private void ExtractPivotTableData(PivotTableCacheDefinitionPart cacheDefinitionPart)
{
    try
    {
        var cacheDefinition = cacheDefinitionPart?.PivotCacheDefinition;
        var cacheFields = cacheDefinition?.CacheFields;

        if (cacheFields == null || !cacheFields.Any())
        {
            ResultsListBox.Items.Add("No fields found in the pivot cache.");
            return;
        }

        ResultsListBox.Items.Add($"Extracting data for Pivot Cache:");
        foreach (var element in cacheFields)
        {
            if (element is not CacheField field) continue;

            string fieldName = field.Name?.Value ?? "Unnamed Field";
            var sharedItems = field.SharedItems;
            var values = new List<string>();

            if (sharedItems != null)
            {
                foreach (var item in sharedItems.ChildElements)
                {
                    switch (item)
                    {
                        case StringItem stringItem:
                            values.Add(stringItem.Val ?? "Empty");
                            break;
                        case NumberItem numberItem:
                            values.Add(numberItem.Val?.ToString() ?? "0");
                            break;
                        case BooleanItem booleanItem:
                            values.Add(booleanItem.Val.ToString());
                            break;
                        default:
                            values.Add("Unsupported Type");
                            break;
                    }
                }
            }

            ResultsListBox.Items.Add($"- Field: {fieldName}, Values: {string.Join(", ", values)}");
        }
    }
    catch (Exception ex)
    {
        ResultsListBox.Items.Add($"Error extracting pivot cache data: {ex.Message}");
    }
}


private PivotTableCacheDefinitionPart GetPivotCacheDefinitionPartById(WorkbookPart workbookPart, uint cacheId)
{
    try
    {
        // Find the PivotCache with matching CacheId
        var pivotCache = workbookPart.Workbook.PivotCaches?
            .Elements<PivotCache>()
            .FirstOrDefault(pc => pc.CacheId?.Value == cacheId);

        if (pivotCache == null)
        {
            ResultsListBox.Items.Add($"No PivotCache found with ID {cacheId}.");
            return null;
        }

        // Get the related PivotTableCacheDefinitionPart
        return workbookPart.GetPartById(pivotCache.Id!) as PivotTableCacheDefinitionPart;
    }
    catch (Exception ex)
    {
        ResultsListBox.Items.Add($"Error finding PivotCacheDefinitionPart: {ex.Message}");
        return null;
    }
}

private Dictionary<string, List<string>> ExtractPivotCacheData(PivotTableCacheDefinitionPart cachePart)
{
    var data = new Dictionary<string, List<string>>();

    try
    {
        var cacheDefinition = cachePart.PivotCacheDefinition;
        var fields = cacheDefinition?.CacheFields?.Elements<CacheField>();

        if (fields == null || !fields.Any())
        {
            ResultsListBox.Items.Add("No cache fields found.");
            return data;
        }

        foreach (var field in fields)
        {
            var fieldName = field.Name?.Value ?? "Unnamed Field";
            var values = new List<string>();

            var sharedItems = field.SharedItems;
            if (sharedItems != null)
            {
                foreach (var child in sharedItems.ChildElements)
                {
                    switch (child)
                    {
                        case StringItem stringItem:
                            values.Add(stringItem.Val ?? "Empty");
                            break;
                        case NumberItem numberItem:
                            values.Add(numberItem.Val?.ToString() ?? "0");
                            break;
                        case BooleanItem booleanItem:
                            values.Add(booleanItem.Val.ToString());
                            break;
                        default:
                            values.Add("Unsupported Type");
                            break;
                    }
                }
            }

            data[fieldName] = values;
        }
    }
    catch (Exception ex)
    {
        ResultsListBox.Items.Add($"Error extracting pivot cache data: {ex.Message}");
    }

    return data;
}


private PivotTableCacheDefinitionPart GetPivotCachePart(PivotTablePart pivotTablePart, WorkbookPart workbookPart)
{
    // Get the CacheId from the PivotTableDefinition
    var cacheId = pivotTablePart.PivotTableDefinition?.CacheId?.Value;
    if (cacheId == null) return null;

    // Find the PivotCache with matching CacheId
    var pivotCache = workbookPart.Workbook.PivotCaches?
        .Elements<PivotCache>()
        .FirstOrDefault(pc => pc.CacheId?.Value == cacheId);

    if (pivotCache == null) return null;

    // Get the PivotTableCacheDefinitionPart using the PivotCache's ID
    return workbookPart.GetPartById(pivotCache.Id!) as PivotTableCacheDefinitionPart;
}





        // Dispose the document when the window is closed
        protected override void OnClosed(EventArgs e)
        {
            if (_spreadsheetDocument != null)
            {
                _spreadsheetDocument.Dispose(); // Use Dispose instead of Close
            }
            base.OnClosed(e);
        }
    }
}
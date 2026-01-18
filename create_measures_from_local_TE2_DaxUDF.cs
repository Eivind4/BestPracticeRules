// ================================================================================================
// Title: Create Time Intelligence Measures for Selected Measures with Functions
// Author: Eivind Haugen
// ================================================================================================
//
// DESCRIPTION:
// This script automates the creation of time intelligence (or other calculated) measures by 
// applying single-parameter functions to selected base measures. It generates new measures with
// properly formatted names, display folders, and metadata based on function annotations.
//
// HOW TO USE:
// 1. Select one or more base measures in your model
// 2. Run the script
// 3. Select one or more single-parameter functions from the dialog (functions starting with 
//    "Local" will be shown)
// 4. If functions are missing annotations, you'll be prompted to provide measure and folder 
//    suffixes (once per function)
// 5. New measures will be created for each combination of selected measure + function
//
// REQUIREMENTS:
// - Select at least one measure before running the script
// - Functions must have exactly ONE parameter (multi-parameter functions are filtered out)
// - Functions should ideally have the required annotations (see below)
//         , but is not applicable for all as FormatString(Expression) from the base measure will be applied 
//
// FUNCTION ANNOTATIONS USED:
// The script reads the following annotations from functions to configure the generated measures:
//
// • MeasurePrefix         - Prefix to add to the measure name (e.g., "PY" → "PY Sales")
// • MeasureSuffix         - Suffix to add to the measure name (e.g., "LY" → "Sales LY")
// • FolderPrefix          - Prefix for the display folder path
// • FolderSuffix          - Suffix for the display folder path (e.g., "Last Year")
// • FormatString          - Format string for the new measure (e.g., "#,0.0%")
// • FormatStringExpression - Dynamic format string expression
// • Description           - Description to append to the base measure's description
//
// If MeasurePrefix/MeasureSuffix are missing:  You'll be prompted to enter a suffix for each function
// If FolderPrefix/FolderSuffix are missing: You'll be prompted to enter a folder suffix for each function
// If FormatString is missing: The script uses the base measure's format (or percentage format for 
//                             functions containing "%", "PCT", "IDX", or "INDEX" in their name)
//
// VALIDATION CHECKS:
// - Warns if base measures are missing FormatString/FormatStringExpression
// - Warns if measures or functions are missing descriptions
// - Prevents duplicate measure creation (checks if measure with same expression already exists)
// - Shows a summary of any measures that already exist and were skipped
//
// EXAMPLE:
// Base measure: [Sales] with FormatString "#,0"
// Function: Local_PY with annotation MeasureSuffix = "PY"
// Result: New measure [Sales PY] with expression "Local_PY([Sales])" and format "#,0"
//
// ================================================================================================


using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Drawing;
using System.Linq;
using System.Collections.Generic;

// Hide script dialog and cursor
ScriptHelper.WaitFormVisible = false;
Application.UseWaitCursor = false;

// Helper function to count parameters in a function
Func<Function, int> CountFunctionParameters = (function) =>
{
    string expression = function.Expression;
    
    if (string.IsNullOrWhiteSpace(expression))
        return 0;
    
    int openParenIndex = expression.IndexOf('(');
    int closeParenIndex = -1;
    
    if (openParenIndex == -1)
        return 0;
    
    for (int i = openParenIndex + 1; i < expression.Length - 2; i++)
    {
        if (expression[i] == ')')
        {
            string remaining = expression.Substring(i + 1).TrimStart();
            if (remaining.StartsWith("=>"))
            {
                closeParenIndex = i;
                break;
            }
        }
    }
    
    if (closeParenIndex == -1)
        return 0;
    
    string parameterSection = expression.Substring(openParenIndex + 1, closeParenIndex - openParenIndex - 1);
    int colonCount = parameterSection.Count(c => c == ':');
    
    return colonCount;
};

// Helper method for input dialog
Func<string, string, string, string> ShowInputDialog = (text, caption, defaultValue) =>
{
    Form prompt = new Form();
    prompt.Width = 500;
    prompt. Height = 200;
    prompt.FormBorderStyle = FormBorderStyle.FixedDialog;
    prompt.Text = caption;
    prompt.StartPosition = FormStartPosition.CenterScreen;
    prompt.MaximizeBox = false;
    prompt.MinimizeBox = false;
    
    Label textLabel = new Label();
    textLabel.Left = 20;
    textLabel.Top = 20;
    textLabel.Width = 450;
    textLabel.Height = 60;
    textLabel.Text = text;
    textLabel.AutoSize = false;
    
    TextBox textBox = new TextBox();
    textBox.Left = 20;
    textBox. Top = 90;
    textBox.Width = 450;
    textBox.Text = defaultValue;
    
    Button confirmation = new Button();
    confirmation.Text = "OK";
    confirmation.Left = 320;
    confirmation.Width = 75;
    confirmation.Top = 125;
    confirmation.DialogResult = DialogResult.OK;
    
    Button cancel = new Button();
    cancel.Text = "Cancel";
    cancel.Left = 405;
    cancel.Width = 75;
    cancel.Top = 125;
    cancel.DialogResult = DialogResult.Cancel;
    
    confirmation.Click += (sender, e) => { prompt.Close(); };
    cancel.Click += (sender, e) => { prompt.Close(); };
    
    prompt.Controls.Add(textLabel);
    prompt.Controls.Add(textBox);
    prompt.Controls.Add(confirmation);
    prompt.Controls.Add(cancel);
    prompt.AcceptButton = confirmation;
    prompt.CancelButton = cancel;
    
    return prompt. ShowDialog() == DialogResult.OK ?  textBox.Text : string.Empty;
};

List<Function> selectedFunctions = new List<Function>();
bool onlyMeasuresSelected = Selected.Measures.Any() && ! Selected.Functions.Any();

if (onlyMeasuresSelected)
{
    var allFunctions = Model.Functions.ToList();
    var localFunctions = allFunctions.Where(f => f.Name.StartsWith("Local")).ToList();
    
    var singleParameterFunctions = localFunctions
            .Where(f => CountFunctionParameters(f) == 1)
            .OrderBy(f => f.Name)
            .ToList();

    if (allFunctions.Count == 0)
    {
        Error("No functions found in the model.");
        return;
    }

    if (! singleParameterFunctions.Any())
    {
        Error("No functions with exactly one input parameter found.");
        return;
    }

    Form functionForm = new Form();
    functionForm.Text = "Select Single-Parameter Functions";
    functionForm.Width = 450;
    functionForm.Height = 580;
    functionForm.StartPosition = FormStartPosition.CenterScreen;
    
    Label functionLabel = new Label();
    functionLabel.Text = "Select one or more single-parameter functions:";
    functionLabel.Left = 20;
    functionLabel.Top = 15;
    functionLabel.Width = 400;
    functionLabel.Height = 30;

    ListBox functionListBox = new ListBox();
    functionListBox.Items.AddRange(singleParameterFunctions.Select(f => f.Name).ToArray());
    functionListBox.SelectionMode = SelectionMode.MultiExtended;
    functionListBox. Left = 20;
    functionListBox.Top = 50;
    functionListBox.Width = 400;
    functionListBox.Height = 430;

    int buttonY = 500;
    int buttonWidth = 80;
    int buttonSpacing = 20;
    int totalButtonWidth = (buttonWidth * 2) + buttonSpacing;
    int startX = (functionForm.Width - totalButtonWidth) / 2;

    Button functionOkButton = new Button();
    functionOkButton.Text = "OK";
    functionOkButton. Left = startX;
    functionOkButton.Top = buttonY;
    functionOkButton.Width = buttonWidth;
    functionOkButton.Height = 30;
    functionOkButton.DialogResult = DialogResult.OK;

    Button functionCancelButton = new Button();
    functionCancelButton.Text = "Cancel";
    functionCancelButton.Left = startX + buttonWidth + buttonSpacing;
    functionCancelButton.Top = buttonY;
    functionCancelButton.Width = buttonWidth;
    functionCancelButton.Height = 30;
    functionCancelButton.DialogResult = DialogResult.Cancel;

    functionForm.Controls.Add(functionLabel);
    functionForm.Controls.Add(functionListBox);
    functionForm.Controls.Add(functionOkButton);
    functionForm.Controls.Add(functionCancelButton);

    functionForm.AcceptButton = functionOkButton;
    functionForm.CancelButton = functionCancelButton;

    var dialogResult = functionForm.ShowDialog();

    if (dialogResult == DialogResult.Cancel)
    {
        return;
    }

    var selectedFunctionNames = functionListBox. SelectedItems.Cast<string>().ToList();
    selectedFunctions = singleParameterFunctions.Where(f => selectedFunctionNames.Contains(f.Name)).ToList();
}
else
{
    if (! Selected.Functions.Any())
    {
        Error("No Functions selected.");
        return;
    }
    
    selectedFunctions = Selected.Functions.Where(f => CountFunctionParameters(f) == 1).ToList();
    
    if (!selectedFunctions.Any())
    {
        Error("None of the selected functions have exactly one input parameter.");
        return;
    }
}

if (selectedFunctions == null || selectedFunctions.Count == 0)
{
    Error("No single-parameter functions selected.");
    return;
}

var missingFormatString = new List<string>();
var missingMeasureDescription = new List<string>();
var missingFunctionDescription = new List<string>();
var alreadyExistingMeasures = new List<string>();

foreach (var f in selectedFunctions)
{
    if (string.IsNullOrWhiteSpace(f.Description))
        missingFunctionDescription.Add(f.Name);
}

foreach (var m in Selected.Measures)
{
    bool hasFormatString = ! string.IsNullOrWhiteSpace(m.FormatString);
    bool hasFormatStringExpression = !string.IsNullOrWhiteSpace(m.FormatStringExpression);
    
    if (! hasFormatString && !hasFormatStringExpression)
    {
        missingFormatString.Add(m.Name);
    }
}

if (missingFormatString.Any() || missingMeasureDescription.Any() || missingFunctionDescription.Any())
{
    string infoMessage = "Missing metadata detected:\n" +
                     "- Option 1: Press Cancel, fix items, then run script again\n" +
                     "- Option 2: Continue and validate created format string and descriptions\n";
    if (missingFormatString.Any())
    {
        infoMessage += "\n• FormatString or FormatStringExpression missing in measures:\n";
        foreach (var name in missingFormatString)
            infoMessage += "   - " + name + "\n";
    }

    if (missingMeasureDescription.Any())
    {
        infoMessage += "\n• Description missing in measures:\n";
        foreach (var name in missingMeasureDescription)
            infoMessage += "   - " + name + "\n";
    }

    if (missingFunctionDescription.Any())
    {
        infoMessage += "\n• Description missing in functions:\n";
        foreach (var name in missingFunctionDescription)
            infoMessage += "   - " + name + "\n";
    }

    var result = MessageBox.Show(infoMessage, "Missing Metadata", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
    if (result == DialogResult.Cancel)
    {
        return;
    }
}

// Pre-process functions to get user input for missing suffixes ONCE per function
var functionSuffixes = new Dictionary<string, string>();

foreach (var f in selectedFunctions)
{
    string measureSuffix = f.GetAnnotation("MeasureSuffix");
    string measurePrefix = f.GetAnnotation("MeasurePrefix");
    
    // Only ask for suffix if both prefix and suffix annotations are missing
    if (string. IsNullOrWhiteSpace(measurePrefix) && string.IsNullOrWhiteSpace(measureSuffix))
    {
        string suffix = ShowInputDialog(
            "Enter suffix for the measure name, from the selected function:",
            "Define Measure Suffix",
            f. Name  // Default value
        );
        
        functionSuffixes[f.Name] = suffix ??  "";
    }
}

// Now loop through measures and functions
foreach (var m in Selected.Measures)
{
    foreach (var f in selectedFunctions)
    {
        string formatString;
        string formatStringExpression;
    
        string annotationFormatString = f.GetAnnotation("FormatString");
        string annotationFormatStringExpression = f.GetAnnotation("FormatStringExpression");
        
        if (! string.IsNullOrWhiteSpace(annotationFormatString))
        {
            formatString = annotationFormatString;
        }
        else
        {
            bool isPercentage = f.Name.ToUpper().Contains("%") ||
                               f.Name.ToUpper().Contains("PCT") ||
                               f.Name.ToUpper().Contains("IDX") ||
                               f. Name.ToUpper().Contains("INDEX");
            
            formatString = isPercentage ? "#,0.0%\u003B-#,0.0%\u003B#,0.0%" : m.FormatString;
        }

        if (!string.IsNullOrWhiteSpace(annotationFormatStringExpression))
        {
            formatStringExpression = annotationFormatStringExpression;
        }
        else
        {        
            formatStringExpression = m.FormatStringExpression;
        }

        string expectedDaxExpression = f.Name + "(" + m.DaxObjectName + ")";

        bool measureExists = Model.AllMeasures.Any(existingMeasure =>
            existingMeasure. Expression. Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "") 
            == expectedDaxExpression. Replace(" ", "").Replace("\n", "").Replace("\r", "").Replace("\t", "")
            && existingMeasure. Table.Name == m.Table. Name);

        string measurePrefix = f.GetAnnotation("MeasurePrefix");
        string measureSuffix = f.GetAnnotation("MeasureSuffix");

        string newMeasureName;

        if (!string.IsNullOrWhiteSpace(measurePrefix) || !string.IsNullOrWhiteSpace(measureSuffix))
        {
            string prefix = !string.IsNullOrWhiteSpace(measurePrefix) ? measurePrefix.Trim(' ', '\'') + " " : "";
            string suffix = !string.IsNullOrWhiteSpace(measureSuffix) ? " " + measureSuffix. Trim(' ', '\'') : "";
            
            newMeasureName = (prefix + m.Name + suffix).Trim(' ', '\'');
        }
        else
        {
            // Use the pre-collected suffix for this function
            string suffix = functionSuffixes[f.Name];
            
            newMeasureName = string.IsNullOrWhiteSpace(suffix) 
                ? m.Name. Trim(' ', '\'')
                : string.Format("{0} {1}", m.Name, suffix).Trim(' ', '\'');
        }

        if (measureExists)
        {
            alreadyExistingMeasures.Add(string.Format("{0} (in {1})", newMeasureName, m.Table.Name));
        }
        else
        {
            string folderPrefix = f.GetAnnotation("FolderPrefix");
            string folderSuffix = f.GetAnnotation("FolderSuffix");

            string displayFolder;

            if (!string. IsNullOrWhiteSpace(folderPrefix) || !string.IsNullOrWhiteSpace(folderSuffix))
            {
                string prefix = !string.IsNullOrWhiteSpace(folderPrefix) ? folderPrefix.Trim() + " " : "";
                string suffix = ! string.IsNullOrWhiteSpace(folderSuffix) ? " - " + folderSuffix.Trim() : "";
                
                displayFolder = (prefix + m.DisplayFolder + "\\" + m.Name + suffix).Trim();
            }
            else
            {
                displayFolder = m.DisplayFolder + "\\" + m.Name;
            }

            var newMeasure = m.Table.AddMeasure(newMeasureName, expectedDaxExpression, displayFolder);

            newMeasure.FormatString = formatString;
            newMeasure.FormatStringExpression = formatStringExpression;
            newMeasure.Description = string.IsNullOrWhiteSpace(m.Description) || string.IsNullOrWhiteSpace(f.Description)
                ? ""
                : m.Description + ", " + f.Description;
            
            newMeasure.FormatDax();
            newMeasure.Expression = "\n" + newMeasure. Expression;
        }
    }
}

if (alreadyExistingMeasures.Any())
{
    string message = "The following measures already exist and were not created:\n\n";
    foreach (var name in alreadyExistingMeasures)
        message += " - " + name + "\n";
    MessageBox.Show(message, "Existing Measures", MessageBoxButtons. OK, MessageBoxIcon.Information);
}
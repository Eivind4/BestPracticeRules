// Title: Create TI measures for selected measures with Functions
// Author: Eivind Haugen

using System;
using System. Text.RegularExpressions;
using System.Windows.Forms;
using System.Drawing;
using System.Linq;
using System. Collections.Generic;

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

foreach (var m in Selected.Measures)
{
    foreach (var f in selectedFunctions)
    {
        string formatString;
        string formatStringExpression;
    
        string annotationFormatString = f. GetAnnotation("FormatString");
        string annotationFormatStringExpression = f.GetAnnotation("FormatStringExpression");
        
        if (!string.IsNullOrWhiteSpace(annotationFormatString))
        {
            formatString = annotationFormatString;
        }
        else
        {
            bool isPercentage = f.Name.ToUpper().Contains("%") ||
                               f.Name.ToUpper().Contains("PCT") ||
                               f.Name. ToUpper().Contains("IDX") ||
                               f.Name. ToUpper().Contains("INDEX");
            
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
            && existingMeasure.Table.Name == m.Table.Name);

        string measurePrefix = f.GetAnnotation("MeasurePrefix");
        string measureSuffix = f.GetAnnotation("MeasureSuffix");

        string newMeasureName;

        if (!string.IsNullOrWhiteSpace(measurePrefix) || !string.IsNullOrWhiteSpace(measureSuffix))
        {
            string prefix = !string.IsNullOrWhiteSpace(measurePrefix) ? measurePrefix.Trim(' ', '\'') + " " : "";
            string suffix = !string.IsNullOrWhiteSpace(measureSuffix) ? " " + measureSuffix.Trim(' ', '\'') : "";
            
            newMeasureName = (prefix + m.Name + suffix).Trim(' ', '\'');
        }
        else
        {
            string textToExclude = ShowInputDialog(
                "Enter text to exclude from function name.\nThe remaining text will be used as suffix:",
                "Define Measure Suffix",
                f.Name
            );
            
            string suffix = f.Name;
            int idx = f.Name.ToUpper().IndexOf(textToExclude. ToUpper());
            if (idx >= 0)
            {
                suffix = f.Name.Remove(idx, textToExclude.Length).Trim();
            }
            
            newMeasureName = string.IsNullOrWhiteSpace(suffix) 
                ? m.Name. Trim(' ', '\'')
                : (m.Name + " " + suffix).Trim(' ', '\'');
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

            if (!string.IsNullOrWhiteSpace(folderPrefix) || !string.IsNullOrWhiteSpace(folderSuffix))
            {
                string prefix = !string.IsNullOrWhiteSpace(folderPrefix) ? folderPrefix.Trim() + " " : "";
                string suffix = !string.IsNullOrWhiteSpace(folderSuffix) ? " - " + folderSuffix. Trim() : "";
                
                displayFolder = (prefix + m.DisplayFolder + "\\" + m.Name + suffix).Trim();
            }
            else
            {
                displayFolder = m. DisplayFolder + "\\" + m. Name;
            }

            var newMeasure = m.Table.AddMeasure(newMeasureName, expectedDaxExpression, displayFolder);

            newMeasure.FormatString = formatString;
            newMeasure.FormatStringExpression = formatStringExpression;
            newMeasure.Description = m.Description + " - using function (" + f.Name + "): " + f.Description;
            
            newMeasure.Expression = FormatDax(newMeasure.Expression);
            newMeasure.Expression = "\n" + newMeasure.Expression;
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
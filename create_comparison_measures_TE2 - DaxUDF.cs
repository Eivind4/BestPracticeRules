// Title: Create TI measures for selected measures with Functions
// Author: Eivind Haugen

#r "Microsoft.VisualBasic"
using System;
using Microsoft.VisualBasic;
using System.Windows.Forms;
using System.Drawing;
using System.Linq;
using System. Collections.Generic;

// Hide script execution dialog and cursor
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
    
    for (int i = openParenIndex + 1; i < expression. Length - 2; i++)
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
    
    string parameterSection = expression.Substring(openParenIndex + 1, closeParenIndex - openParenIndex - 1).Trim();
    
    if (string.IsNullOrWhiteSpace(parameterSection))
        return 0;
    
    int colonCount = parameterSection.Count(c => c == ':');
    
    if (colonCount > 0)
    {
        return colonCount;
    }
    
    string cleanParameterSection = parameterSection;
    var lines = parameterSection.Split('\n');
    var cleanLines = new List<string>();
    
    foreach (var line in lines)
    {
        string cleanLine = line. Trim();
        
        if (cleanLine.StartsWith("//"))
            continue;
        
        int commentIndex = cleanLine.IndexOf("//");
        if (commentIndex >= 0)
        {
            cleanLine = cleanLine. Substring(0, commentIndex).Trim();
        }
        
        if (!string.IsNullOrWhiteSpace(cleanLine))
        {
            cleanLines.Add(cleanLine);
        }
    }
    
    cleanParameterSection = string.Join(" ", cleanLines. ToArray());
    
    if (string.IsNullOrWhiteSpace(cleanParameterSection))
        return 0;
    
    int commaCount = cleanParameterSection.Count(c => c == ',');
    
    return commaCount > 0 ?  commaCount + 1 :  1;
};

// Check if exactly two measures are selected
if (Selected.Measures.Count != 2)
    throw new Exception("You must select exactly two measures!");

// Retrieve the selected measures
var measure1 = Selected. Measures.ElementAt(0);
var measure2 = Selected. Measures.ElementAt(1);

List<Function> selectedFunctions = new List<Function>();

// Check if only measures are selected (no functions)
bool onlyMeasuresSelected = Selected.Measures.Any() && !Selected.Functions.Any();

if (onlyMeasuresSelected)
{
    var allComparisonFunctions = Model.Functions.Where(f => f.Name.StartsWith("Comparison")).ToList();
    var twoParameterComparisonFunctions = allComparisonFunctions. Where(f => CountFunctionParameters(f) == 2).ToList();

    if (allComparisonFunctions.Count == 0)
    {
        Error("No Comparison functions found in the model.");
        return;
    }

    if (twoParameterComparisonFunctions.Count == 0)
    {
        Error(string.Format("Found {0} Comparison functions, but none have exactly two input parameters.", allComparisonFunctions.Count));
        return;
    }

    Form functionForm = new Form();
    functionForm.Text = "Select Two-Parameter Comparison Functions";
    functionForm.Width = 450;
    functionForm.Height = 580;
    functionForm.StartPosition = FormStartPosition.CenterScreen;

    Label functionLabel = new Label();
    functionLabel.Text = string.Format("Select one or more two-parameter comparison functions ({0} available):", twoParameterComparisonFunctions.Count);
    functionLabel. Left = 20;
    functionLabel.Top = 15;
    functionLabel.Width = 400;
    functionLabel.Height = 30;

    ListBox functionListBox = new ListBox();
    functionListBox.Items.AddRange(twoParameterComparisonFunctions.Select(f => f. Name).ToArray());
    functionListBox.SelectionMode = SelectionMode.MultiExtended;
    functionListBox.Left = 20;
    functionListBox. Top = 50;
    functionListBox.Width = 400;
    functionListBox.Height = 430;

    int buttonY = 500;
    int buttonWidth = 80;
    int buttonSpacing = 20;
    int totalButtonWidth = (buttonWidth * 2) + buttonSpacing;
    int startX = (functionForm.Width - totalButtonWidth) / 2;

    Button functionOkButton = new Button();
    functionOkButton.Text = "OK";
    functionOkButton.Left = startX;
    functionOkButton.Top = buttonY;
    functionOkButton.Width = buttonWidth;
    functionOkButton.Height = 30;
    functionOkButton.DialogResult = DialogResult.OK;

    Button functionCancelButton = new Button();
    functionCancelButton. Text = "Cancel";
    functionCancelButton.Left = startX + buttonWidth + buttonSpacing;
    functionCancelButton.Top = buttonY;
    functionCancelButton.Width = buttonWidth;
    functionCancelButton.Height = 30;
    functionCancelButton.DialogResult = DialogResult.Cancel;

    functionForm. Controls.Add(functionLabel);
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

    var selectedFunctionNames = functionListBox.SelectedItems.Cast<string>().ToList();
    selectedFunctions = twoParameterComparisonFunctions.Where(f => selectedFunctionNames.Contains(f.Name)).ToList();
}
else
{
    if (!Selected.Functions.Any())
    {
        Error("No Functions selected.");
        return;
    }
    selectedFunctions = Selected.Functions.Where(f => f.Name.StartsWith("Comparison") && CountFunctionParameters(f) == 2).ToList();
    
    if (!selectedFunctions.Any())
    {
        Error("None of the selected functions are Comparison functions with exactly two input parameters.");
        return;
    }
}

if (selectedFunctions == null || selectedFunctions.Count == 0)
{
    Error("No two-parameter comparison functions selected.");
    return;
}

var missingFormatString = new List<string>();
var missingMeasureDescription = new List<string>();
var missingFunctionDescription = new List<string>();
var alreadyExistingMeasures = new List<string>();

foreach (var m in Selected.Measures)
{
    if (string.IsNullOrWhiteSpace(m. FormatString))
        missingFormatString.Add(m. Name);
    if (string.IsNullOrWhiteSpace(m.Description))
        missingMeasureDescription.Add(m.Name);
}

foreach (var f in selectedFunctions)
{
    if (string. IsNullOrWhiteSpace(f.Description))
        missingFunctionDescription.Add(f. Name);
}

if (missingFormatString.Any() || missingMeasureDescription.Any() || missingFunctionDescription. Any())
{
    string infoMessage = "Missing metadata detected:\n" +
                     "- Option 1: Press Cancel, fix items, then run script again\n" +
                     "- Option 2: Continue and validate created format string and descriptions\n";

    if (missingFormatString. Any())
    {
        infoMessage += "\n• FormatString missing in measures:\n";
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

var finalTargetTable = measure1.Table;

foreach (var function in selectedFunctions)
{
    string measureName;
    string functionNameLower = function.Name.ToLower();
    
    if (function.Name.EndsWith("Deviation"))
    {
        measureName = measure1.Name + " Dev";
    }
    else if (functionNameLower. Contains("pct") || 
             functionNameLower.Contains("idx") || 
             functionNameLower.Contains("index"))
    {
        measureName = measure1.Name + " % change";
    }
    else
    {
        measureName = function.Name.Replace("Comparison", "").Trim();
        if (string.IsNullOrEmpty(measureName))
            measureName = function.Name;
    }

    if (finalTargetTable.Measures.Any(m => m.Name == measureName))
    {
        alreadyExistingMeasures.Add(measureName);
        continue;
    }

    var daxExpression = function.Name + "(" + measure1.DaxObjectName + ", " + measure2.DaxObjectName + ")";

    var newMeasure = finalTargetTable. AddMeasure(
        measureName,
        daxExpression,
        measure1.DisplayFolder
    );

    string formatString;
    if (functionNameLower.Contains("pct") || 
        functionNameLower. Contains("idx") || 
        functionNameLower.Contains("index"))
    {
        formatString = "#,0.0%\u003B-#,0.0%\u003B#,0.0%";
    }
    else
    {
        formatString = measure1.FormatString;
    }
    newMeasure.FormatString = formatString;

    string description = function.Description ?? ("Comparison function " + function.Name);
    description = description.Replace("measure1", measure1.DaxObjectFullName);
    description = description.Replace("measure2", measure2.DaxObjectFullName);
    newMeasure.Description = description;

    newMeasure.Expression = FormatDax(newMeasure.Expression);
}

if (alreadyExistingMeasures.Any())
{
    string existingMessage = "The following measures already exist and were skipped:\n";
    foreach (var name in alreadyExistingMeasures)
        existingMessage += "   - " + name + "\n";
    
    MessageBox.Show(existingMessage, "Already Existing Measures", MessageBoxButtons. OK, MessageBoxIcon.Information);
}
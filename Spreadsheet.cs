using Newtonsoft.Json;
using System;
using System.Runtime.ConstrainedExecution;
using System.Text.RegularExpressions;
using SpreadsheetUtilities;
using static System.Net.Mime.MediaTypeNames;
namespace SS
{

    [JsonObject(MemberSerialization.OptIn)]
    public class Spreadsheet : AbstractSpreadsheet
    {
        private bool changed;

        private DependencyGraph graph;

        private Dictionary<string, Cell> nonemptyCells = new();

        [JsonProperty(PropertyName ="cells")]
        private Dictionary<string, Cell> nonemptyCellsStringForm = new();

        private string variablePattern = "^[a-zA-Z]+[a-zA-z0-9]*$";

        public override bool Changed { get => changed; protected set => changed = value; }

        /// <summary>
        /// // Empty zero parameter constructor to create an empty Spreadsheet with no specific validator, normalizer or version.
        /// </summary>
        public Spreadsheet() : base(s => true, s => s, "default")
        {
            graph = new DependencyGraph(); 
        }

        /// <summary>
        /// // Three parameter constructor to create an empty Spreadsheet with a validator, normalizer and version provided by the user.
        /// </summary>
        /// <param name="isValid"></param> - Validator to determine what a valid variable name is.
        /// <param name="normalize"></param> - Normalizer to normalize all variables to a standard format.
        /// <param name="version"></param> - The specific version that the Spreadsheet was created with.
        public Spreadsheet(Func<string, bool> isValid, Func<string, string> normalize, string version) : base(isValid, normalize, version)
        {
            graph = new DependencyGraph();
        }

        /// <summary>
        /// // Four parameter constructor to create a new Spreadsheet from an existing spreadsheet with a validator, normalizer and version provided by the user.
        /// </summary>
        /// <param name="fileToPath"></param> - Filepath of an existing spreadsheet to be used to create this speradsheet.
        /// <param name="isValid"></param> - Validator to determine what a valid variable name is.
        /// <param name="normalize"></param> - Normalizer to normalize all variables to a standard format.
        /// <param name="version"></param> - The specific version that the Spreadsheet was created with.
        public Spreadsheet(string fileToPath, Func<string, bool> isValid, Func<string, string> normalize, string version) : base(isValid, normalize, version)
        {
            try
            {
                graph = new DependencyGraph();
                Spreadsheet? restored;
                string? file = File.ReadAllText(fileToPath);
                if (file != null)
                {
                    restored = JsonConvert.DeserializeObject<Spreadsheet>(file);
                    if (restored != null)
                    {
                        if (!restored.Version.Equals(version))
                            throw new SpreadsheetReadWriteException("Version of saved spreadsheet is different than current version.");
                        foreach (KeyValuePair<string, Cell> cell in restored.nonemptyCellsStringForm)
                        {
                            if (!Regex.IsMatch(cell.Key, variablePattern))
                                throw new SpreadsheetReadWriteException("Saved spreadsheet contains an invalid variable name.");
                            if (nonemptyCellsStringForm != null)
                            {
                                string? content = cell.Value.ContentString;
                                SetContentsOfCell(cell.Key, content);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (ex is CircularException)
                    throw new SpreadsheetReadWriteException("A circular dependency has been encountered.");
                else if (ex is FormulaFormatException)
                    throw new SpreadsheetReadWriteException("A syntactically incorrect formula has been encountered.");
                else
                {
                    throw new SpreadsheetReadWriteException("An error has been encountered in the process of reading the saved spreadsheet.");
                }
            }
        }

        /// <summary>
        /// // Saves the contents of this spreadsheet to a JSON file with the provided filename.
        /// </summary>
        /// <param name="filename"></param> - Filename to be saved to.
        /// <exception cref="SpreadsheetReadWriteException"></exception>
        public override void Save(string filename)
        {
            string jsonAsString = JsonConvert.SerializeObject(this, Newtonsoft.Json.Formatting.Indented);
            try
            {
                File.WriteAllText(filename, jsonAsString);
            }
            catch
            {
                throw new SpreadsheetReadWriteException("Spreadsheet was unable to save to file.");
            }
            Changed = false;
        }

        /// <summary>
        /// // Calculates the values of a cell. If the name is not a valid cell name then an InvalidName
        /// </summary>
        /// <param name="name"></param> - Name of cells value to be calculated.
        /// <returns></returns>
        /// <exception cref="InvalidNameException"></exception> - An exception that is thrown if the provided name is invalid.
        public override object GetCellValue(string name)
        {
            string normalizedName = Normalize(name);
            // Checks to see if name is a valid variable name, i.e. it is any amount of letters followed by any amount of numbers.
            if (!Regex.IsMatch(normalizedName, variablePattern))
                throw new InvalidNameException();
            else
            {
                // Checks to see if cell exists in nonemptyCells.
                if (nonemptyCells.TryGetValue(normalizedName, out Cell? cell))
                {

                    // Checks to see if content inside cell is a Formula.
                    if (cell.Content.GetType() == typeof(Formula))
                    {
                        Formula formula = (Formula)cell.Content;
                        object finalValue = formula.Evaluate(LookupMethod);
                        cell.Value = finalValue;
                        return finalValue;
                    }

                    // If cell.Content is a double or string then that it is returned.
                    else
                    {
                        return cell.Content;
                    }
                }
                else
                {
                    return "";
                }
            }
        }

        /// <summary>
        /// Gets and returns a dictionary of all non-empty cells in spreadsheet.
        /// </summary>
        /// <returns></returns> - All non-empty cells in the spreadsheet.
        public override IEnumerable<string> GetNamesOfAllNonemptyCells()
        {
            HashSet<string> enumeratedCells = new HashSet<string>();
            foreach (KeyValuePair<string, Cell> cell in nonemptyCells)
            {
                string cellName = cell.Key;
                enumeratedCells.Add(cellName);
            }
            return enumeratedCells;
        }

        /// <summary>
        /// Gets the contents of a specific cell using name.
        /// </summary>
        /// <param name="name"></param> - Name of cell.
        /// <returns></returns> - Contents of the desired cell.
        /// <exception cref="InvalidNameException"></exception>
        public override object GetCellContents(string name)
        {
            string normalizedName = Normalize(name);
            if (!Regex.IsMatch(normalizedName, variablePattern))
                throw new InvalidNameException();
            else
            {
                if (nonemptyCells.TryGetValue(normalizedName, out Cell? content))
                    return content.Content;
                else
                {
                    return "";
                }
            }
        }

        public override IList<string> SetContentsOfCell(string name, string content)
        {
            IsValid(name);
            string normalizedName = Normalize(name);
            // If name is not a valid cell name then an InvalidNameException is thrown,
            if (!Regex.IsMatch(normalizedName, variablePattern))
                throw new InvalidNameException();

            if (double.TryParse(content, out double num))
            {
                SetCellContents(normalizedName, num);
            }
            else if (content[0].Equals('='))
            {
                string remainderOfFormula = content.Remove(0, 1);
                Formula f = new Formula(remainderOfFormula);
                SetCellContents(normalizedName, f);
            }
            else
            {
                SetCellContents(normalizedName, content);
            }

            IList<string> cellsToRecalculate = GetCellsToRecalculate(normalizedName).ToList();

            foreach (string cell in cellsToRecalculate)
            {
                if (!cell.Equals(normalizedName))
                {
                    Object contents = nonemptyCells[cell].Content;
                    if (contents.GetType() == typeof(Formula))
                    {
                        SetCellContents(cell, (Formula)contents);
                    }
                }

            }

            Changed = true;

            //Returns a list of the current cell, along with all of the cells that are dependent on that cell.
            return cellsToRecalculate;
        }

        /// <summary>
        /// Sets the contents of a celll using a double.
        /// </summary>
        /// <param name="name"></param> - Name of cell.
        /// <param name="number"></param> - Double to be added to cell.
        /// <returns></returns> - A list containing the name of the cell along with all dependents of that cell.
        /// Whether they are directly or indirectly dependent on the cell.
        /// <exception cref="InvalidNameException"></exception>
        protected override IList<string> SetCellContents(string name, double number)
        {
            string normalizedName = Normalize(name);
            // Checks to see if name already exists as a cell.
            if (nonemptyCells.TryGetValue(normalizedName, out Cell? cell))
            {

                // If the previous content of name was a formula then the dependees of name are reset.
                if (cell.Content.GetType() == typeof(Formula))
                {
                    IEnumerable<string> emptyDependees = new List<string>();
                    graph.ReplaceDependees(normalizedName, emptyDependees);
                }

                // Value of cell is changed to text.
                cell.Content = number;
            }

            // If name was not already contained in nonemptyCells then a new cell is made and number is assigned to cell.
            else if (!nonemptyCells.ContainsKey(normalizedName))
            {
                Cell newCell = new Cell(number);
                nonemptyCells.Add(normalizedName, newCell);
                nonemptyCellsStringForm.Add(normalizedName, newCell);
            }

            // Converts IEnumberable<string> to IList<string.
            IList<string> dependents = GetCellsToRecalculate(normalizedName).ToList();

            // Returns all direct and indirect dependents of name.
            return dependents;
        }

        /// <summary>
        /// Sets the contents of a cell using a string.
        /// </summary>
        /// <param name="name"></param> - Name of cell.
        /// <param name="text"></param> - String to be added to cell.
        /// <returns></returns> - A list containing the name of the cell along with all dependents of that cell.
        /// Whether they are directly or indirectly dependent on the cell.
        /// <exception cref="InvalidNameException"></exception>
        protected override IList<string> SetCellContents(string name, string text)
        {
            string normalizedName = Normalize(name);
            // Checks to see if name already exists as a cell.
            if (nonemptyCells.TryGetValue(normalizedName, out Cell? cell))
            {

                // If the previous content of name was a formula then the dependees of name are reset.
                if (cell.Content.GetType() == typeof(Formula))
                {
                    IEnumerable<string> emptyDependees = new List<string>();
                    graph.ReplaceDependees(normalizedName, emptyDependees);
                }

                // Value of cell is changed to text.
                cell.Content = text;
            }

            // If name was not already contained in nonemptyCells then a new cell is made and text is assigned to cell.
            else if (!nonemptyCells.ContainsKey(normalizedName))
            {
                Cell newCell = new Cell(text);
                nonemptyCells.Add(normalizedName, newCell);
                nonemptyCellsStringForm.Add(normalizedName, newCell);
            }

            // Converts IEnumberable<string> to IList<string.
            IList<string> dependents = GetCellsToRecalculate(normalizedName).ToList();

            // Returns all direct and indirect dependents of name.
            return dependents;
        }

        /// <summary>
        /// Sets the contents of a cell using a formula.
        /// </summary>
        /// <param name="name"></param> - Name of cell.
        /// <param name="formula"></param> - Formula to be added to cell.
        /// <returns></returns> - A list containing the name of the cell along with all dependents of that cell.
        /// Whether they are directly or indirectly dependent on the cell.
        /// <exception cref="InvalidNameException"></exception>
        protected override IList<string> SetCellContents(string name, Formula formula)
        {
            string normalizedName = Normalize(name);
            // Tries to run code within try block.
            try
            {
                // Checks to see if formula contains variables.
                if (formula.GetVariables().Count() > 0)
                {

                    // Adds a dependency for each variable in formula.
                    foreach (string variable in formula.GetVariables())
                    {
                        graph.AddDependency(variable, normalizedName);
                    }
                }

                // GetCellsToRecalculate is called. If a circular dependency is gonna be added then a CircularException is thrown before the content of the
                // designated cell can be changed.
                GetCellsToRecalculate(normalizedName).ToList();
            }
            
            // If exception is thrown then code is reversed and previously added dependencies are removed.
            catch (CircularException)
            {
                foreach (string variable in formula.GetVariables())
                {
                    graph.RemoveDependency(variable, normalizedName);
                }
                throw new CircularException();
            }

            // Checks to see if name already exists as a cell.
            if (nonemptyCells.TryGetValue(normalizedName, out Cell? cell))
            {

                // If the previous content of name was a formula then the dependees of name are reset.
                if (cell.Content.GetType() == typeof(Formula))
                {
                    IEnumerable<string> emptyDependees = new List<string>();
                    graph.ReplaceDependees(normalizedName, emptyDependees);
                }

                // Value of cell is changed to text.
                cell.Content = formula;
            }
            // If name was not already contained in nonemptyCells then a new cell is made and formula is assigned to cell.
            else if (!nonemptyCells.ContainsKey(normalizedName))
            {
                Cell newCell = new Cell(formula);
                nonemptyCells.Add(normalizedName, newCell);
                nonemptyCellsStringForm.Add(normalizedName, newCell);
            }

            // Returns all direct and indirect dependents of name.
            return GetCellsToRecalculate(normalizedName).ToList();

        }

        /// <summary>
        /// Returns the direct dependents of a cell.
        /// </summary>
        /// <param name="name"></param> - The name of the cell.
        /// <returns></returns> - The direct dependents of a cell.
        /// <exception cref="NotImplementedException"></exception>
        protected override IEnumerable<string> GetDirectDependents(string name)
        {
                IEnumerable<string> directDependents = graph.GetDependents(name);
                return directDependents;
        }


        private double LookupMethod(string name)
        {
            if (nonemptyCells.TryGetValue(name, out Cell? value))
            {
                if (value.Content.GetType() == typeof(double))
                {
                    return (double)value.Content;
                }
                else if (value.Content.GetType() == typeof(Formula))
                {
                    return (double)value.Value;
                }
                else
                {
                    throw new ArgumentException();
                }
            }
            else
            {
                throw new ArgumentException();
            }
        }

        [JsonObject(MemberSerialization.OptIn)]
        private class Cell
        {
            
            private object contents;
            private object? cellValue;

            [JsonProperty(PropertyName = "stringForm")]
            private string contentsString;

            public Cell()
            {
                this.contents = "";
                this.contentsString = "";
            }
            // Creates a constructor for a double.
            public Cell(double content)
            {
                this.contents = content;
                this.contentsString = content.ToString();
            }

            // Creates a constructor for a string.
            public Cell(string content)
            {
                this.contents = content;
                this.contentsString = content;
            }

            // Creates a constructor for a formula.
            public Cell(Formula content)
            {
                this.contents = content;
                this.contentsString = "=" + content.ToString();
            }

            public object Content
            {
                get
                {
                    return contents;
                }
                set
                {
                    contents = value;
                }
            }

            public object Value
            {
                get
                {
                    if (cellValue != null)
                    {
                        return cellValue;
                    }
                    return "";
                }
                set
                {
                    cellValue = value;
                }
            }

            public string ContentString
            {
                get
                {
                    return contentsString;
                }
            }
        }
    }

}
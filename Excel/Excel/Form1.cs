using System;
using Exc = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;


namespace Excel
{
    public partial class Form1 : Form
    {
        Dictionary<string, List<string>> finishedItems = new Dictionary<string, List<string>>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            { e.Effect = DragDropEffects.Copy; }
            else if (e.Data.GetDataPresent("FileGroupDescriptor"))
            { e.Effect = DragDropEffects.Copy; }
            else
            { e.Effect = DragDropEffects.None; }
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop,false) == true)
                {
                    Exc.Application xlApp = new Exc.Application();
                    Exc.Range range = null;
                    string[] fileName = (string[])e.Data.GetData(DataFormats.FileDrop);
                    string location = fileName[0].Trim();
                    Exc.Workbook xlWorkbook = xlApp.Workbooks.Open(location);
                    Exc._Worksheet selection = null;

                    selection = xlWorkbook.Sheets[1];
                    
                    range = selection.UsedRange;

                    int rrCount = 0;                   
                    int rows = range.Rows.Count;
                    string styleNumber = null; //The actual style number as opposed to Cat#
                    string type = null;
                    string catNumber = null;
                    string yardage = null;
                    string fullDescription = null;
                    string fullOperationInfo = null;
                    int onRowNumber = 0;
                    // It's stringed because there is a list of strings that will need
                    // this value later on
                    string railroaded = "False";

                    // Getting railroad info
                    bool[] railRoad = PopulateRailroadInfo(selection, rows);
                    for (int i = 1; i <= rows; i++)
                    {
                        
                        bool found = false;
                        var val = (selection.Cells[i, 1] as Exc.Range).Value;
                        
                        // for some reason, the values that seem to be integers are stored as doubles, 
                        // so a record is selected when dealing with any type of number
                        if ((val != null) && Utils.IsNumber(val))
                        {
                            string info = null;
                            string styleName = (selection.Cells[i, 2] as Exc.Range).Value.Trim();
                            type = Utils.GetProperType(((selection.Cells[i, 10] as Exc.Range).Value).ToString());
                            // if perentheses are involved, then there is information aside form just the style number
                            if (styleName.Contains('('))
                            {
                                string[] nameAndInfo = styleName.Split('(');
                                styleName = nameAndInfo[0].Trim();
                                info = nameAndInfo[1].Trim();
                                if(info[info.Length - 1] == ')'){ info = info.Remove(info.Length - 1, 1); }
                            }

                            bool differentPart = false;
                            if(info != null)
                            {
                                string parsedInfo = "";
                                bool railRoadToBool = railroaded.ToLower() == "true" ? true : false;
                                List<string> itemParts = new List<string>();
                                itemParts = Utils.GetPartsFromItem(info, type, railRoadToBool);
                                if (itemParts.Count > 0) { parsedInfo = Utils.ConvertPartsToQuery(itemParts); }
                                // If the style already exists, but the parts for it are different, a different query
                                // must be implemented
                                if ((finishedItems.ContainsKey(styleName) && (finishedItems[styleName].ElementAt(5) != parsedInfo || finishedItems[styleName].ElementAt(4).ToLower() != railroaded.ToLower())))
                                {
                                    differentPart = true;
                                }
                            }
                            // Signifies that there's already an item by that style existing, but the part provided part(s)
                            // differs. So we might've gotten a seat already, but now, it's a back, so the yardages will differ 
                            if (differentPart){
                                bool rr = railRoad[rrCount];
                                IndividualItem item = new IndividualItem(styleName, rr, info, type);

                                if (item.fullDescription != null)
                                {
                                    found = true;
                                    styleNumber = item.fullDescription;
                                    type = item.type;
                                    yardage = item.yardage.ToString();
                                    catNumber = item.catNumber;
                                    fullDescription = item.fullDescription;
                                    railroaded = item.railroaded.ToString();
                                    info = item.itemQuery;
                                    fullOperationInfo = item.operationDetails;
                                }

                            }
                            // If the same item has already been querried, there's no reason to run a query again, so 
                            // a dictionary is used to store previously-used models
                            else if (finishedItems.ContainsKey(styleName))
                            {
                                found = true;
                                yardage = finishedItems[styleName].ElementAt(0);
                                type = finishedItems[styleName].ElementAt(1);
                                catNumber = finishedItems[styleName].ElementAt(2);
                                fullDescription = finishedItems[styleName].ElementAt(3);
                                railroaded = finishedItems[styleName].ElementAt(4);
                                fullOperationInfo = finishedItems[styleName].ElementAt(6);
                            }
                            else
                            {
                                bool rr = railRoad[rrCount];
                                // add the matching option
                                IndividualItem item = new IndividualItem(styleName, rr, info, type);
                                if (item.fullDescription != null)
                                {
                                    found = true;
                                    styleNumber = item.fullDescription;
                                    type = item.type;
                                    yardage = item.yardage.ToString();
                                    catNumber = item.catNumber;
                                    fullDescription = item.fullDescription;
                                    railroaded = item.railroaded.ToString();
                                    info = item.itemQuery;
                                    fullOperationInfo = item.operationDetails;

                                    // If the sheet contains multiple instances of the same item, there's no reason to run the same
                                    // query multiple times, so all the information about that model is stored in a dictionary                                    
                                    finishedItems.Add(styleName, new List<string>() { yardage, type, catNumber, fullDescription, railroaded, info, fullOperationInfo });
                                }
                            }
                            if (found)
                            {
                                double orderTotalYardage;
                                double pricePerYard;
                                double valuePerUnit;
                                double yar;
                                yar = Double.Parse(yardage);
                                orderTotalYardage = val * yar;

                                string possiblePrice = ((selection.Cells[i, 19] as Exc.Range).Value) == null ? "" : ((selection.Cells[i, 19] as Exc.Range).Value).ToString();
                                // If the value in the yardage field exists and is a number 
                                if (possiblePrice != ""){
                                    if (double.TryParse(possiblePrice, out double num)) {
                                        pricePerYard = num;
                                        valuePerUnit = orderTotalYardage * pricePerYard / val;
                                        (selection.Cells[i, 20] as Exc.Range).Value = valuePerUnit;
                                    }
                                }

                                info = (info != null && info.Length != 0) ? " (" + info.ToUpper() + ')' : "";
                                fullOperationInfo = fullOperationInfo.Remove(fullOperationInfo.Length - 2, 2);

                                // If cal'ed
                                if (((selection.Cells[i, 6] as Exc.Range).Value) != null) { // Had to make nested ifs or it was throwing an error
                                    if (((selection.Cells[i, 6] as Exc.Range).Value) != "") { fullOperationInfo += "\ncal = 3"; yar += 3; } 
                                }

                                if (type == "NM") {
                                    (selection.Cells[i, 13] as Exc.Range).Value = "";
                                    (selection.Cells[i, 15] as Exc.Range).Value = yar;
                                }
                                else {
                                    (selection.Cells[i, 13] as Exc.Range).Value = yar;
                                    (selection.Cells[i, 15] as Exc.Range).Value = "";
                                }

                                if (railroaded == "True") {
                                    (selection.Cells[i, 8] as Exc.Range).Value = "x";
                                    (selection.Cells[i, 7] as Exc.Range).Value = "";
                                }
                                else
                                {
                                    (selection.Cells[i, 8] as Exc.Range).Value = "";
                                    (selection.Cells[i, 7] as Exc.Range).Value = "x";
                                }

                                (selection.Cells[i, 23] as Exc.Range).Value = fullOperationInfo;
                                (selection.Cells[i, 2] as Exc.Range).Value = (fullDescription + info);
                                (selection.Cells[i, 17] as Exc.Range).Value = orderTotalYardage;

                            }

                            string measurement;
                            var measurementInCell = ((selection.Cells[i, 11] as Exc.Range).Value);
                            if (measurementInCell == null) { measurementInCell = ""; }
                            measurement = Utils.InchesToMilimeters(measurementInCell.ToString().Trim());                         
                            if (measurement != null)
                            {
                                (selection.Cells[i, 11] as Exc.Range).Value += "\n" + measurement;
                                (selection.Cells[i, 11] as Exc.Range).NumberFormat = "#,##0.00";
                            }

                            measurementInCell = ((selection.Cells[i, 12] as Exc.Range).Value);
                            if (measurementInCell == null) { measurementInCell = ""; }
                            measurement = Utils.InchesToMilimeters(measurementInCell.ToString().Trim());
                            if (measurement != null)
                            {
                                (selection.Cells[i, 12] as Exc.Range).Value += "\n" + measurement;
                                (selection.Cells[i, 12] as Exc.Range).NumberFormat = "#,##0.00";
                            }
                            (selection.Cells[i, 10] as Exc.Range).Value = type;
                            rrCount++;
                            onRowNumber = i; // Used for printing purposes
                        }
                    }
                    selection.PageSetup.PrintArea = $"A1:V{onRowNumber}";
                    range.Columns[2].AutoFit();
                    range.Columns[22].AutoFit();

                    xlWorkbook.Save();
                    xlWorkbook.Close();
                    xlApp.Quit();

                    System.Diagnostics.Process.Start($"{location}");
                }
            }
            
            catch(Exception ex)
            {
                MessageBox.Show(ex + "");
            }
        }

        private bool[] PopulateRailroadInfo(Exc._Worksheet selection, int rows)
        {
            string[] styleNameAndItsParts = null;
            string itemParts = "";
            string providedStyleName = null;
            string type = null;
            bool railroaded;

            RailoadForm rlf = new RailoadForm();
            for (int j = 1; j <= rows; j++)
            {
                var val = (selection.Cells[j, 1] as Exc.Range).Value;
                List<string> partsList = new List<string>();
                // Checking to see whether the row is part of the data
                if (val != null && Utils.IsNumber(val))
                {
                    if ((selection.Cells[j, 8] as Exc.Range).Value != null) { railroaded = true; } else { railroaded = false; }
                    type = ((selection.Cells[j, 10] as Exc.Range).Value).ToString();
                    providedStyleName = (selection.Cells[j, 2] as Exc.Range).Value.Trim();

                    // Ideally, if there is any information beside the style or cat names provided, it will be seperated by a '('
                    if (providedStyleName.Contains('(')) {
                        styleNameAndItsParts = providedStyleName.Split('(');
                        providedStyleName = styleNameAndItsParts[0];
                        itemParts = styleNameAndItsParts[1];
                        partsList = Utils.GetPartsFromItem(itemParts, type, railroaded);
                    }
                    string fabricInfo = (selection.Cells[j, 4] as Exc.Range).Value + '-' + (selection.Cells[j, 5] as Exc.Range).Value;

                    if (partsList.Count > 0)
                    {
                        rlf.PopulateItems(providedStyleName, fabricInfo, railroaded, partsList);
                    }
                    else
                    {
                        rlf.PopulateItems(providedStyleName, fabricInfo, railroaded);
                    }
                }
            }
            rlf.ShowDialog();

            return rlf.railroaded;
        }


        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}

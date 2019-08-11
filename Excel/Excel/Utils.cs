using System;
using System.Collections.Generic;
using System.Linq;
using System.Data.OleDb;
using System.Text.RegularExpressions;
using System.IO;

namespace Excel
{
    static class Utils
    {
        public static List<string> GetPartsFromItem(string info, string type, bool railroaded)
        {
            string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\Yardage.accdb");
            String s = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath}; Persist Security Info=False;";

            List<string> allTableColumns = new List<string>();
            string railRoadedPrefix = railroaded == true ? "RR" : "";
            string query = $"SELECT * FROM {railRoadedPrefix}{type}Yardage";

            OleDbConnection connection = new OleDbConnection(s);
            connection.Open();

            OleDbCommand command = new OleDbCommand(query, connection);
            OleDbDataReader reader = command.ExecuteReader();

            string name = null;
            for (int i = 0; i < reader.FieldCount; i++)
            {
                name = reader.GetName(i);
                // If this method is implemented, then we're working with yardage parts. So the columns
                // that describe the fabric or the whole item yardage are not to be considered
                if (!(name == "styleName" || name == "type" || name == "whole item")) { allTableColumns.Add(name); }
            }

            reader.Close();
            connection.Close();

            List<string> modelItems = new List<string>();
            List<string> infoList = info.ToLower().Split(new string[] { ','.ToString(), '&'.ToString(), "and", "only" }, StringSplitOptions.RemoveEmptyEntries).ToList();

            // If the string includes "except", then we need to get all the values but the one(s)
            // that is specified
            bool exceptPresent = false;
            List<string> partsLookedThrough = new List<string>();
            foreach (string item in infoList)
            {
                foreach (string possibility in allTableColumns)
                {
                    if (item.Contains("except")) { exceptPresent = true; }
                    if (item.Length >= possibility.Length && item.Contains(possibility)) {
                        if (!(modelItems.Contains(possibility)))
                        {
                            string pos = possibility;
                            if (pos.Contains(' ')) { pos = '[' + pos + ']'; }
                            modelItems.Add(pos);
                        }
                    }
                }

            }
            if (modelItems.Count > 0)
            {
                // Get the opposite of values that are in possibleItemParts but not in modelItems
                if (exceptPresent) { modelItems = allTableColumns.ToList().Except(modelItems).ToList(); }

                modelItems.Sort((a, b) => b.Length.CompareTo(a.Length));
            }
            return modelItems;
        }

        public static string ConvertPartsToQuery(List<string> parts)
        {
            string partsQuery = "";
            foreach (string item in parts)
            {
                partsQuery += item + ',';
            }
            // there's a superfulous "," at the end, so we get rid of it
            partsQuery = partsQuery.Remove(partsQuery.Length - 1, 1);

            return partsQuery;
        }

        public static string InchesToMilimeters(string measurement)
        {
            measurement = measurement.Trim();
            if (measurement.Length > 0 && measurement[measurement.Length - 1] == '"') { measurement = measurement.Remove(measurement.Length - 1); }
            Regex measurementRegex = new Regex(@"\s*0+\.?0*[^0-9]\s*");
            MatchCollection match = measurementRegex.Matches(measurement);
            System.Windows.Forms.MessageBox.Show(match.Count + "");
            if (match.Count > 0 || measurement == "" || measurement == null) { return null; }

            string wholeNumber = null;
            string remainder = null;
            double milimeters = 0;
            if (measurement != null || measurement != "0")
            {
                measurement = measurement.Trim();
                // The inches will usually in two forms -- '123.12' or '123 1/3"'. 
                // Both are accounted for
                // This is the '123.12' case
                if (Double.TryParse(measurement, out double number))
                {
                    if (number > 0)
                    {
                        milimeters = 25.4 * number;
                    }
                }
                else
                {
                    if (measurement.Contains(' '))
                    {
                        wholeNumber = measurement.Split(' ')[0];
                        remainder = measurement.Split(' ')[1];
                        // If the string is something like 25 1/3
                        if (double.TryParse(wholeNumber, out double whole) && int.TryParse(remainder.Split('/')[0], out int n) && int.TryParse(remainder.Split('/')[1], out int d))
                        {
                            milimeters = (whole + (double)n / d) * 25.4;
                        }
                    }
                    else
                    {
                        // If the string is something like 1/3
                        if (int.TryParse(measurement.Split('/')[0], out int numerator) && int.TryParse(measurement.Split('/')[1], out int denominator))
                        {
                            milimeters = ((double)numerator / denominator) * 25.4;
                        }
                    }
                }
            }
            return milimeters.ToString();
        }

        public static bool IsNumber(object val)
        {
            bool isnumber = false;
            if (val.GetType() == typeof(double) || val.GetType() == typeof(int) || val.GetType() == typeof(float))
            {
                isnumber = true;
            }
            return isnumber;
        }

        public static string GetProperType(string type)
        {
            if(type != null) //Splitting the ifs for safety
            {
                if(type != "")
                {
                    if (type.Contains("2")) { type = "2M"; }
                    else if (type.Contains("1")) { type = "1M"; }
                    else { type = "NM"; }
                    return type;
                }
                return "NM";
            }
            return "NM"; //Default Type
        }
    }
}

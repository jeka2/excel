using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.IO;

namespace Excel
{
    class IndividualItem
    {
        public string fullDescription = null;
        public string catNumber = null;
        public string type = null;
        public double yardage;
        public bool railroaded = false;
        string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\..\Yardage.accdb");

        public string itemQuery = "";
        Dictionary<string, double> partAndCorrespondingYardage = new Dictionary<string, double>();
        List<string> parts = new List<string>();
        // This will contain all the parts used and their yardage displayed
        // for easy understanding of where the values came from
        public string operationDetails = null;

        OleDbConnection connection = new OleDbConnection();

        public IndividualItem(string style, bool railroaded, string partsInfo = null, string type = null)
        {
            this.type = type.ToUpper().Trim();
            this.railroaded = railroaded;

            PopulateUnitInfo(style, partsInfo);
        }

        private void PopulateUnitInfo(string style, string potentialUnitParts)
        {
            GetFullDescription(style);
            // fullDescription will be null if it wasn't found within the database
            // and if it wasn't, nothing else needs to be done with this item
            if (this.fullDescription != null)
            {
                if (potentialUnitParts != null)
                {
                    parts = Utils.GetPartsFromItem(potentialUnitParts, this.type, this.railroaded);
                    this.itemQuery = parts.Count > 0 ? Utils.ConvertPartsToQuery(parts) : "";

                    // specialQuery refers to querrying an item's parts as opposed
                    // to the whole item
                    int numPartsToQuery = parts.Count;
                    GetTypeAndYardage(numPartsToQuery);
                }
                else { GetTypeAndYardage(); }
            }
        }

        public void GetFullDescription(string val)
        {
            string fullField = null;
            try
            {
                String s = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\TRA-FS1\AccessFiles\2007 Master Production Database_New.accdb;
Persist Security Info=False;";
                connection.ConnectionString = s;
                connection.Open();

                OleDbCommand command = new OleDbCommand();
                string query = $"SELECT [Master Style Table].[Style Number], [Master Style Table].[Cat #] FROM [Master Style Table] WHERE [Master Style Table].[Style Number] = '{val}' OR [Master Style Table].[Cat #] = '{val}';";
                command.Connection = connection;
                command.CommandText = query;
                OleDbDataReader reader = command.ExecuteReader();
                while (reader.Read())
                {
                    string styleNumber = (string)reader[0];
                    string catNumber = (string)reader[1];

                    // Saving the cat number for saving into yardage database, as cat number is most frequently used
                    this.catNumber = catNumber;

                    fullField = styleNumber + '/' + catNumber;
                    break;
                }
                reader.Close();
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("Error" + ex);
            }
            connection.Close();

            this.fullDescription = fullField;
        }

        private void GetTypeAndYardage(int partsToQuery = 0)
        {
            string style = this.catNumber;
            string type = null;
            string query = "";
            try
            {
                String s = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath}; Persist Security Info=False;";
                connection.ConnectionString = s;
                connection.Open();

                OleDbCommand command = new OleDbCommand();
                string railRoadPrefix = this.railroaded == true ? "RR" : "";
                if (partsToQuery > 0) {
                    query = $"SELECT styleName, type, {this.itemQuery} FROM {railRoadPrefix}{this.type}Yardage WHERE styleName = '{style}';";
                }
                else {
                    query = $"SELECT styleName, type, [whole item] FROM {railRoadPrefix}{this.type}Yardage WHERE styleName = '{style}';";
                    parts.Add("[whole item]");
                    partsToQuery++; // Accounting for the whole item
                }

                command.Connection = connection;
                command.CommandText = query;
                bool recordExists = false;
                bool updateInOrder = false;
                double totalYardage = 0;

                OleDbDataReader reader = command.ExecuteReader();
                // The loop will run only once as the styles are a primary key, but 
                // having it in this format is convenient in several ways
                while (reader.Read())
                {
                    recordExists = true;
                    // If the style exists
                    if ((string)reader[0] == style)
                    {
                        // the + 2 is to account for styleName and type which are not yardages
                        for (int i = 2; i < partsToQuery + 2; i++)
                        {
                            
                            double.TryParse(reader[i].ToString(), out double num);
                            // If a part is provided and is 0 in the database, it needs to be updated
                            if (num == 0) { updateInOrder = true; operationDetails = null; break; }
                            partAndCorrespondingYardage.Add(parts.ElementAt(i - 2), num);
                            totalYardage += num;


                            operationDetails += parts.ElementAt(i - 2) + '=' + num + "/\n";
                        }
                        // If the style exists with no obvious errors, we'll go ahead with the information received
                        string matchType = reader[1].ToString().ToUpper();
                        if ((matchType == "NM" || matchType == "1M" || matchType == "2M") && totalYardage != 0)
                        {
                            type = (string)reader[1];
                            this.yardage = totalYardage;
                            this.type = type;
                        }
                        // If the style exists, but doesn't adhere to some basic principles, it needs to be updated
                        else
                        {
                            // The values will need to be modified if something associated with the style isn't as expected
                            updateInOrder = true;
                        }
                    }
                }

                reader.Close();
                connection.Close();
                
                if (updateInOrder)
                {
                    UpdateRecord();
                }
                if (!recordExists)
                {
                    AddRecord();
                }
            }
            catch (Exception ex)
            {
                System.Console.WriteLine("Error" + ex);
            }
        }

        // This method adds a record to the yardage database. The database keeps track of yardages and types
        // The method also sets the yardage and the type
        private void AddRecord()
        {
            string style = this.catNumber;
            string q = "";
            String s = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath}; Persist Security Info=False;";

            connection.ConnectionString = s;
            connection.Open();

            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;

            string railRoadPrefix = this.railroaded == true ? "RR" : "";

            YardageForm yf = new YardageForm(style, this.railroaded, this.type);
            if (parts.Count == 0)
            {
                yf.ShowDialog();
                q = $"INSERT INTO {railRoadPrefix}{this.type}Yardage (styleName, type, [whole item]) VALUES ('{style}', '{yf.type}', {yf.yardage})";

                this.type = yf.type;
                this.yardage = yf.yardage;

                operationDetails += "whole item" + '=' + yf.yardage + "/\n";
            }
            else
            {
                string augmentedStyle = "";
                string queryType = "";
                string queryValue = "";
                string fabricType = null;
                double totalYardage = 0;
                foreach(string part in parts)
                {
                    augmentedStyle = style + '(' + part + ')';
                    yf.Reinitialize(augmentedStyle);
                    yf.ShowDialog();
                    queryType += part + ',';
                    queryValue += yf.yardage + ", ";
                    totalYardage += yf.yardage;
                    fabricType = yf.type;

                    operationDetails += part + '=' + yf.yardage + "/\n";
                }

                this.type = fabricType;
                this.yardage = totalYardage;

                queryType = queryType.Remove(queryType.Length - 1, 1);
                queryValue = queryValue.Remove(queryValue.Length - 2, 2);
                q = $"INSERT INTO {railRoadPrefix}{this.type}Yardage (styleName, type, {queryType}) VALUES ('{style}', '{fabricType}', {queryValue})";
            }

            command.CommandText = q;
            command.ExecuteNonQuery();

            connection.Close();
        }

        private void UpdateRecord()
        {
            string style = this.catNumber;
            string q = "";
            String s = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbPath}; Persist Security Info=False;";

            connection.ConnectionString = s;
            connection.Open();

            OleDbCommand command = new OleDbCommand();
            command.Connection = connection;

            string railRoadPrefix = this.railroaded == true ? "RR" : "";

            YardageForm yf = new YardageForm(style, this.railroaded, this.type);
            string fabricType = null;
            string augmentedStyle = "";
            string queryCondition = "";
            double totalYardage = 0;
            if (parts.Count == 0)
            {
                yf.ShowDialog();

                this.type = yf.type;
                this.yardage = yf.yardage;

                q = $"UPDATE {railRoadPrefix}{this.type}Yardage SET [whole item]={yf.yardage}, type = '{yf.type}' WHERE styleName = '{style}'";

                operationDetails += "whole item" + '=' + yf.yardage + "/\n";

            }
            else
            {
                foreach (string part in parts)
                {
                    augmentedStyle = style + '(' + part + ')';
                    yf.Reinitialize(augmentedStyle);
                    yf.ShowDialog();
                    queryCondition += $"{part}={yf.yardage},";
                    fabricType = yf.type;
                    totalYardage += yf.yardage;

                    operationDetails += part + '=' + yf.yardage + "/\n";
                }
                queryCondition = queryCondition.Remove(queryCondition.Length - 1, 1);

                this.type = fabricType;
                this.yardage = totalYardage;
                q = $"UPDATE {railRoadPrefix}{fabricType}Yardage SET {queryCondition} WHERE styleName='{style}'";
            }


            command.CommandText = q;
            command.ExecuteNonQuery();

            connection.Close();
        }
    }   
}

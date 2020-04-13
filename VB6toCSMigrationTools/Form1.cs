using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VB6toCSMigrationTools
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnConvertVB6toCS_Click(object sender, EventArgs e)
        {
            txtDest.Text = ConvertVB6toCS(txtSource.Text);
        }

        private string ConvertVB6toCS(string vb6text)
        {
            // first of all, break the instructions one per line

            List<string> lines = Regex.Split(vb6text, "\r\n|\r|\n").ToList(); //vb6text.Split(Environment.NewLine.ToCharArray()).ToList();
            StringBuilder resultLines = new StringBuilder();
            bool isMakingAnEnum = false;

            foreach (var line in lines)
            {
                string resultingLine = string.Empty;


                // Clean the line from messy stuff
                string cleanLine = Regex.Replace(line, @"\s+", " ").Trim();

                // Break the line apart
                string[] arrLine = cleanLine.Split(' ');

                // Get what this command means
                if (arrLine[0].ToLower().Equals("public") || arrLine[0].ToLower().Equals("private"))
                {
                    // It's a declaration
                    string declarationProtectionLevel = arrLine[0].ToLower();

                    if (arrLine[1].ToLower().Equals("enum"))
                    {
                        // Public declaration of enum
                        isMakingAnEnum = true;
                        string enumName = arrLine[2];

                        // If there's more content, it's probably comments
                        string theRestOfLine = GetRestOfLineIfComments(arrLine, 3);

                        // Write down the command
                        resultingLine = declarationProtectionLevel + " enum " + enumName + " " + theRestOfLine + 
                            Environment.NewLine + "{";

                    }
                    else if (arrLine[1].ToLower().Equals("property"))
                    {

                    }
                    else
                    {
                        // it's a field. The second argument is the field name
                        string propertyName = arrLine[1];
                        string propertyDataType = string.Empty;

                        // Try to get the type if available
                        if (arrLine.Length > 2)
                        {
                            if (arrLine[2].ToLower().Equals("as"))
                            {
                                // Here comes the data type
                                propertyDataType = ConvertDataTypeVB6toCS(arrLine[3]);
                            }
                        }
                        else
                        {
                            // No data type
                        }

                        // If there's more content, it's probably comments
                        string theRestOfLine = GetRestOfLineIfComments(arrLine, 4);

                        // Write down the command
                        resultingLine = declarationProtectionLevel + " " + propertyDataType + " " + propertyName + "; " + theRestOfLine;
                    }
                }
                else if (arrLine[0].StartsWith(@"'"))
                {
                    // It's a comment, just write it
                    resultingLine = GetRestOfLineIfComments(arrLine, 0);
                }
                else if (arrLine[0].ToLower().StartsWith(@"end"))
                {
                    // It's ending something
                    if (arrLine.Length > 1)
                    {
                        // Figure out what it is ending
                        if (arrLine[1].ToLower().StartsWith(@"enum"))
                        {
                            // It's ending a enum
                            resultingLine = "}";
                            isMakingAnEnum = false;
                        }
                        resultingLine += GetRestOfLineIfComments(arrLine, 2);
                    }
                }
                else if (isMakingAnEnum)
                {
                    // It's one of the enum options. Just add it as it is
                    bool addedComma = false;
                    for (int i = 0; i < arrLine.Length; i++)
                    {
                        if (arrLine[i].StartsWith(@"'"))
                        {
                            resultingLine += ",";
                            addedComma = true;
                            // Comment started. We can give up here
                            resultingLine += GetRestOfLineIfComments(arrLine, i);
                            break;
                        }
                        resultingLine += arrLine[i] + " ";
                    }
                    if (!addedComma)
                    {
                        resultingLine += ",";
                        addedComma = true;
                    }
                }

                resultLines.AppendLine(resultingLine);
            }

            return resultLines.ToString();
        }

        private string GetRestOfLineIfComments(string[] arrLine, int startIndex)
        {
            if (arrLine.Length < startIndex)
                return string.Empty;

            // If there's more content, it's probably comments
            string theRestOfLine = string.Empty;
            if (arrLine.Length > startIndex)
            {
                if (arrLine[startIndex].StartsWith(@"'"))
                {
                    for (int i = startIndex; i < arrLine.Length; i++)
                    {
                        theRestOfLine += arrLine[i] + " ";
                    }
                    theRestOfLine = "// " + theRestOfLine.Trim();
                }
            }
            // Write down the command
            return theRestOfLine;

        }

        private string ConvertDataTypeVB6toCS(string VB6Type)
        {
            switch (VB6Type.ToLower())
            {
                case "double":
                    return "double";
                case "byte":
                    return "byte";
                case "integer":
                    return "short";
                case "long":
                    return "int";
                case "string":
                    return "string";
                case "boolean":
                    return "bool";
                default:
                    return VB6Type;
            }
        }

    }
}

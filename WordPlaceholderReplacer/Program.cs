using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class Program
{
    static void Main(string[] args)
    {
        string templatePath = "REGISTRATION FORM.docx";  // keep your template here
        string outputPath = "Output.docx";

        // Copy template so original stays untouched
        File.Copy(templatePath, outputPath, true);

        // Ask user for input
        Console.Write("Enter Name: ");
        string name = Console.ReadLine();

        Console.Write("Enter Email Address: ");
        string email = Console.ReadLine();

        Console.Write("Enter Home Phone: ");
        string homePhone = Console.ReadLine();

        Console.Write("Enter Cell Phone: ");
        string cellPhone = Console.ReadLine();

        Console.Write("Enter Address Line One: ");
        string address1 = Console.ReadLine();

        Console.Write("Enter Address Line Two: ");
        string address2 = Console.ReadLine();

        Console.Write("Enter Class: ");
        string className = Console.ReadLine();

        Console.Write("Enter Permit Number: ");
        string permit = Console.ReadLine();

        // Open doc and replace placeholders
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputPath, true))
        {
            var body = wordDoc.MainDocumentPart.Document.Body;

            ReplaceText(body, "NAME", name);
            ReplaceText(body, "EMAIL ADDRESS", email);
            ReplaceText(body, "HOME PHONE", homePhone);
            ReplaceText(body, "CELL PHONE", cellPhone);
            ReplaceText(body, "ADDRESS LINE ONE", address1);
            ReplaceText(body, "ADDRESS LINE TWO", address2);
            ReplaceText(body, "CLASS", className);
            ReplaceText(body, "PERMIT NUMBER", permit);

            wordDoc.MainDocumentPart.Document.Save();
        }

        Console.WriteLine("✅ Document generated: " + outputPath);
    }

    static void ReplaceText(Body body, string placeholder, string newValue)
    {
        foreach (var text in body.Descendants<Text>())
        {
            if (text.Text.Contains(placeholder))
            {
                text.Text = text.Text.Replace(placeholder, newValue);
            }
        }
    }
}

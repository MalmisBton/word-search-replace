using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Text.RegularExpressions;
using System.IO;
using System.Windows;

/// <todo>
/// 1. Funktionen "ListFiles" kan endast lista filer i vald map + EN undermapp. Om det finns fler undermappar kommer dessa inte med.
/// 2. Hitta bättre sätt att filtrera dolda filer än att filtrera bort filer som innehåller '$' (lol). Fungerar inte på alla filer.
/// 3. Många mallar har dubbelparenteser av någon anledning (EX: "Henrik Malmsjö ((19880707))"). Dessa bör bytas ut mot endast en.
/// 4. Fältkodlistan måste uppdateras. Vissa fältkoder har formen [[x.???]]. Detta betyder att jag inte hitttat korresponderade kod i nya ÖFS. 
/// 5. Finns troligtvis ett snabbare sätt att byta ut fältkoder i dokumentet:
/// Nuvarande funktion använder sig av "document.Content.Find.Execute" och loopar igenom för varje fältkod-par.
/// En snabbare variant vore troligtvis att 
/// (1) kopiera innehållet i dokumnetet till en sträng.
/// (2) söka igenom strängen efter fältkoder.
/// (3) ladda om nyckel och värdelistorna.
/// (4) Använda sig av document.Content.Find.Execute med listorna som ges av (3)
/// Detta eftersom fältkodlistan nu innehåller cirka 90 fältkoder och de flesta mallar använder ~10. 
/// document.Content.Find.Execute har även visat sig vara rätt segt...
/// 6. Det måste finnas ett sätt att extrahera fältkoderna från ett textdokument som INTE involverar tre List<string>-objekt...
/// Kanske dictionary?
/// 7. "Extrahera fältkoder från textfil" är inte implementerad än.
/// </todo>

namespace WordReplaceV2
{

    public class Menu
    {
        private static void Text(string[] menuList)
        {
            int counter = 0;
            foreach (var item in menuList)
            {
                counter++;
                Console.WriteLine("[" + counter + "] " + item);
            }
        }

        private static int Selector(string[] menuList)
        {
            int selection = 0;
            Console.Write("Val: ");
            while (selection == 0 || selection < 0 || selection > menuList.Length + 1)
            {
                Int32.TryParse(Console.ReadLine(), out selection);
            }
            return selection;
        }

        public static int Display(string[] menuList)
        {
            Menu.Text(menuList);
            int selection = Menu.Selector(menuList);
            return selection;
        }

    }

    class Program
    {


        static void Main(string[] args)
        {
            // Lokaliserar vart programmet körs från och sparar till sträng.
            string programPath = Assembly.GetExecutingAssembly().CodeBase.ToString();
            programPath = Path.GetDirectoryName(programPath);
            programPath = new Uri(programPath).LocalPath;

            // Fillista som används i programmet.
            #region File and folder list
            string documentList = programPath + @"\Mallar";
            string textDumpOutput = programPath + @"\Text-dump.txt";
            string fieldCodeOutput = programPath + @"\Fieldcode-dump.txt";
            string translationTableTxt = programPath + @"\Translation-table.txt";
            #endregion

            #region Menus


            int menuSelector = 0;

            while (true)
            {
                Console.Clear();
                menuSelector = Menu.Display(new string[] {
                            "Updatera fältkoder.",
                            "Skapa textdump av dokumentlista.",
                            "Lista alla fältkoder i dokumentlista.",
                            "Extrahera alla fältkoder i textfil.",
                            "Inställningar.",
                            "Avsluta"});

                switch (menuSelector)
                {
                    //Menyval: Uppdatera fältkoder.
                    case 1:
                        {
                            Console.Clear();
                            Console.WriteLine("Uppdatera fältkoder.");
                            Console.WriteLine("\nVald mapp: " + documentList);
                            Console.WriteLine("\nFältkodlista: " + translationTableTxt);
                            Console.WriteLine("\nMapp och fältkodlista kan ändras under inställningar.\n");

                            menuSelector = Menu.Display(new string[] { "Starta", "Bakåt" });

                            switch (menuSelector)
                            {
                                case 1:
                                    {
                                        Console.WriteLine("Uppdaterar fältkoder...\n\n");
                                        ReplaceWords(ListFiles(documentList), translationTableTxt);
                                        break;
                                    }

                                case 2:
                                    {

                                        break;
                                    }
                            }
                            break;
                        }

                        // Menyval: Skapa textdump av dokumentlista.
                    case 2:
                        {
                            Console.Clear();
                            Console.WriteLine("Skapa textdump av dokumentlista.");
                            Console.WriteLine("\nVald mapp: " + documentList);
                            Console.WriteLine("\nTextdump sparas som: " + textDumpOutput);
                            Console.WriteLine("\nMapp och filnamn kan ändras under inställningar.");

                            menuSelector = Menu.Display(new string[] { "Starta", "Bakåt" });

                            switch (menuSelector)
                            {
                                case 1:
                                    {
                                        Console.WriteLine("Skapar textdump...\n\n");
                                        TextDump(ListFiles(documentList), textDumpOutput);
                                        break;
                                    }

                                case 2:
                                    {
                                        break;
                                    }

                            }
                            break;
                        }


                        // Menyval: Lista alla fältkoder i en textfil.
                    case 3:
                        {
                            Console.Clear();
                            Console.WriteLine("Lista alla fältkoder i alla dokument i en mapp.");
                            Console.WriteLine("\nVald mapp: " + documentList);
                            Console.WriteLine("\nFältkodlistan sparas som: " + fieldCodeOutput);
                            Console.WriteLine("\nMapp och fältkodlista kan ändras under inställningar.");

                            menuSelector = Menu.Display(new string[] { "Starta", "Bakåt" });

                            switch (menuSelector)
                            {
                                case 1:
                                    {
                                        Console.WriteLine("Extraherar fältkoder...\n\n");
                                        FieldCodeWriter(FieldCodeExtractor(ListFiles(documentList)), fieldCodeOutput);
                                        break;
                                    }
                                case 2:
                                    {

                                        break;
                                    }
                            }
                            break;
                        }
                        // Menyval: Extrahera fältkoder från textfil.
                    case 4:
                        {
                            Console.Clear();
                            Console.WriteLine("Extrahera fältkoder från textfil.");
                            Console.WriteLine("\nVald textfil:" + textDumpOutput);
                            Console.WriteLine("\nFältkodlistan sparas som: " + fieldCodeOutput);
                            Console.WriteLine("\nVald texfil och fil som listan sparas som kan ändras under inställningar.");

                            menuSelector = Menu.Display(new string[] { "Starta", "Bakåt" });

                            switch (menuSelector)
                            {
                                case 1:
                                    {
                                        Console.WriteLine("Extraherar fältkoder...\n\n");
                                        FieldCodeExtractorTXT(textDumpOutput);
                                        Console.ReadLine();
                                        break;
                                    }

                                case 2:
                                    {

                                        break;
                                    }

                            }
                            break;
                        }

                        //Menyval: Inställningar.
                    case 5:
                        {
                            Console.Clear();
                            Console.WriteLine("Inställningar.\n");
                            Console.WriteLine("Mapp att konvertera: \n" + documentList + "\n\n");
                            Console.WriteLine("Fältkodlista som används vid konvertering: \n" + translationTableTxt + "\n\n");
                            Console.WriteLine("Textdump sparas som: \n" + textDumpOutput + "\n\n");
                            Console.WriteLine("Fältkodlistan sparas som: \n" + fieldCodeOutput + "\n\n");

                            menuSelector = Menu.Display(new string[]
                            {   "Ändra konveteringsmapp",
                                "Ändra översättningstabell",
                                "Ändra vart textdumpsfilen sparas",
                                "Ändra vart fältkodlistan sparas",
                                "Bakåt"});


                            switch (menuSelector)
                            {
                                case 1:
                                    {
                                        Console.Clear();
                                        Console.WriteLine("Ändra konveteringsmapp (skriv \"b\" för att backa utan att ändra");
                                        Console.WriteLine("Nuvarande konverteringsmapp: " + documentList);
                                        string input = Console.ReadLine();

                                        if (input.ToLower() == "b")
                                        {
                                            break;
                                        }

                                        else
                                        {
                                            documentList = input;
                                        }

                                        break;
                                    }

                                case 2:
                                    {
                                        Console.Clear();
                                        Console.WriteLine("Ändra översättningstabell (skriv \"b\" för att backa utan att ändra");
                                        Console.WriteLine("Nuvarande översättningstabell: " + translationTableTxt);
                                        string input = Console.ReadLine();

                                        if (input.ToLower() == "b")
                                        {
                                            break;
                                        }

                                        else
                                        {
                                            translationTableTxt = input;
                                        }

                                        break;

                                    }

                                case 3:
                                    {
                                        Console.Clear();
                                        Console.WriteLine("Ändra vart textdumpen sparas (skriv \"b\" för att backa utan att ändra");
                                        Console.WriteLine("Nuvarande plats: " + textDumpOutput);
                                        string input = Console.ReadLine();

                                        if (input.ToLower() == "b")
                                        {
                                            break;
                                        }

                                        else
                                        {
                                            textDumpOutput = input;
                                        }

                                        break;
                                    }

                                case 4:
                                    {
                                        Console.Clear();
                                        Console.WriteLine("Ändra vart översättningstabellen sparas (skriv \"b\" för att backa utan att ändra");
                                        Console.WriteLine("Nuvarande plats: " + fieldCodeOutput);
                                        string input = Console.ReadLine();

                                        if (input.ToLower() == "b")
                                        {
                                            break;
                                        }

                                        else
                                        {
                                            fieldCodeOutput = input;
                                        }

                                        break;
                                    }

                                case 5:
                                    {
                                        break;
                                    }



                            }

                            break;
                        }

                    case 6:
                        {
                            Environment.Exit(0);
                            break;
                        }

                }

            }


        }

        #endregion

        //List<string> input = ListFiles(@"C:\Users\Henrik\Random\Arbete\Mallar");
        //ReplaceWords(input);

        // List<string> documentList = ListFiles(@"C:\Users\Henrik\Random\Arbete\Mallar men alla filer ligger i en map");
        // TextDump(documentList);

        // Writes the exctraced field codes to a document
        public static void FieldCodeWriter(List<string> input, string outputDirectory)
        {

            File.WriteAllText(outputDirectory, "");

            File.WriteAllLines(outputDirectory, input);
        }

        // Takes all the text contained all documents in a directory and prinits it do a .txt document
        // Useful for debugging the replacer

        // Lägg till argument så användaren kan välja vilken mapp som ska användas.
        public static void TextDump(List<string> filelist, string outputDirectory)
        {

            File.WriteAllText(outputDirectory, "");

            List<string> dumpList = new List<string>();

            for (int i = 0; i < filelist.Count; i++)
            {
                Console.WriteLine("Extracting text from document {0} ...", i + 1);

                string temp = TextExtractor(filelist[i]);
                dumpList.Add("\n----------------------------------------------------------------------------------------\n"
                    + "TITLE:"
                    + filelist[i]
                    + "\n----------------------------------------------------------------------------------------\n"
                    + temp
                    + "\n----------------------------------------------------------------------------------------\n"
                    + "END OF FILE"
                    + "\n----------------------------------------------------------------------------------------\n");

                Console.WriteLine(">> Exctraction successful!\n");
            }

            File.WriteAllLines(outputDirectory, dumpList);
        }

        public static void FieldCodeExtractorTXT(string filePath)
        {
            string serachObjetct = File.ReadAllText(filePath);
            List<string> fieldCodeList = FieldCodeFinder(serachObjetct);

            for (int i = 0; i < fieldCodeList.Count - 3; i++)
            {
                Console.WriteLine("{0}      -       {1}        -        {2}         -       {3}",
                    fieldCodeList[i], fieldCodeList[i + 1], fieldCodeList[i + 2], fieldCodeList[i + 3]);
            }
            Console.WriteLine(fieldCodeList.Count);
        }

        // Takes a document list and returns a list with all of the fieldcodes contained in the documents
        public static List<string> FieldCodeExtractor(List<string> documentList)
        {
            List<string> fieldCodeList = new List<string>();
            List<string> mergerList = new List<string>();

            for (int i = 0; i < documentList.Count; i++)
            {
                mergerList = FieldCodeFinder(TextExtractor(documentList[i]));

                ListMerger(fieldCodeList, mergerList);
            }

            foreach (var item in fieldCodeList)
            {
                Console.WriteLine(item);
            }

            return fieldCodeList;
        }

        // Prints all of the items in a list to console
        public static void ListPriner(List<string> input)
        {
            foreach (var item in input)
            {
                Console.WriteLine(item);
            }
        }

        // Merges two lists while paying attention to any duplicates
        public static List<string> ListMerger(List<string> list1, List<string> list2)
        {
            for (int i = 0; i < list2.Count; i++)
            {
                if (list1.Contains(list2[i]) == false)
                {
                    list1.Add(list2[i]);
                }
            }
            return list1;
        }

        // Searches a string for field codes and returns a list with all contained field codes
        public static List<string> FieldCodeFinder(string input)
        {
            // List<string> MatchList = new List<string>();

            string pattern = @"\[\[.+?(?=\])..";

            MatchCollection matchCollection = Regex.Matches(input, @pattern);

            var matchList = matchCollection.Cast<Match>().Select(match => match.Value).Distinct().ToList();

            foreach (string s in matchList)
            {
                Console.WriteLine(s);
            }

            return matchList;

        }

        // Opens a document and returns all text as a string.
        public static string TextExtractor(string filepath)
        {
            Application application = new Application();
            application.Visible = true;

            Document document = application.Documents.Open(filepath);
            string output = document.Content.Text.ToString();
            //Console.WriteLine(output);

            //File.WriteAllText(@"C:\Users\Henrik\Random\Arbete\output.txt", output);

            document.Save();
            application.Quit();
            ((_Document)document).Close();

            return output;

        }

        // Returns a list of all files in a directory (includes one level of subdirectories)
        public static List<string> ListFiles(string dir)
        {

            var dirList = new List<string>();
            var fileList = new List<string>();

            dirList = Directory.EnumerateDirectories(dir).ToList<string>();
            dirList.Add(dir);


            foreach (var item in dirList)
            {
                Console.WriteLine(item);
            }

            for (int i = 0; i < dirList.Count; i++)
            {
                List<string> tempList = new List<string>();
                tempList = Directory.GetFiles(dirList[i]).ToList<string>();
                fileList = fileList.Union(tempList).ToList();
            }

            int ListLength = fileList.Count();

            fileList.RemoveAll(x => x.Contains('$'));
            fileList.RemoveAll(x => (x.Contains(".dot") || 
                                   x.Contains(".dotx") || 
                                   x.Contains(".doc") || 
                                   x.Contains(".docx")) == false);


            foreach (var item in fileList)
            {
                Console.WriteLine(item);
            }

            Console.WriteLine("\n" + "Number of documents in directory: " + fileList.Count() + "\n");

            return fileList;

        }



        // Replaces words in a document list.
        public static void ReplaceWords(List<string> documentList, string translationTable)
        {
            List<string> KeyValuePair = new List<string>();
            List<string> KeyList = new List<string>();
            List<string> ValueList = new List<string>();

            KeyValuePair = File.ReadAllLines(translationTable).ToList();

            KeyList.Add("[[Rotel]]");
            ValueList.Add("");

            KeyList.Add("Rotel");
            ValueList.Add("");

            //First split 
            // \[\[.+?(?=\]\])\]\]?.{0,}(=>) 

            //Second split 
            // =>?.{0,}\[\[(.+?(?=\])\]\]){1,}

            for (int i = 0; i < KeyValuePair.Count; i++)
            {
                if (KeyValuePair[i].Contains("=>") == true)
                {
                    string untrimmedKey = Regex.Match(KeyValuePair[i], @"\[\[.+?(?=\]\])\]\]?.{0,}(=>)").ToString();
                    untrimmedKey = untrimmedKey.Replace("=>", "");
                    untrimmedKey = untrimmedKey.Trim();

                    string untrimmedValue = Regex.Match(KeyValuePair[i], @"=>?.{0,}\[\[(.+?(?=\])\]\]){1,}").ToString();
                    untrimmedValue = untrimmedValue.Replace("=>", "");
                    untrimmedValue = untrimmedValue.Trim();

                    KeyList.Add(untrimmedKey);
                    ValueList.Add(untrimmedValue);
                }
            }

            Application application = new Application();
            application.Visible = true;


            for (int i = 0; i < documentList.Count(); i++)
            {

                Document document = application.Documents.Open(documentList[i]);
                int testcounter = 1;

                Console.WriteLine("Word replacement in:/n{0}", documentList[i]);

                for (int j = 0; j < KeyList.Count; j++)
                {
                    double percent = ((double)testcounter / (double)KeyList.Count);

                    testcounter++;
                    Console.WriteLine(Math.Round(percent, 2) * 100 + "%");
                    //document.Content.Find.Execute("[[Rotel]]", false, true, false, false, false, true, 1, false, "", 2, false, false, false, false);
                    //document.Content.Find.Execute("Rotel", false, true, false, false, false, true, 1, false, "", 2, false, false, false, false);
                    document.Content.Find.Execute(KeyList[j], false, true, false, false, false, true, 1, false, ValueList[j], 2, false, false, false, false);

                }


                document.Save();
                //document.SaveAs2(FileName: documentList[i], FileFormat: WdSaveFormat.wdFormatTemplate);

                /// application.Quit();
                ///((_Document)document).Close();

                document.Close();
            }
            application.Quit();

        }
    }
}


using CsvHelper;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using MixERP.Net.VCards.Serializer;
using MixERP.Net.VCards;

namespace OutlookComToVcard
{
    class Program
    {
        static void Main(string[] args)
        {
            if(args.Length < 2)
            {
                Console.WriteLine("You need to specify the path of the CSV file and where to put the VCards!");
                Console.WriteLine("OutlookComToVcard.exe <Path to CSV file> <Output Folder> [boolean add 1 day to dates (default true)]");
                return;
            }

            String csvFilePath = args[0];
            String outFolder = args[1];
            Boolean correctDates = true;
            if (args.Length == 3)
            {
                correctDates = Boolean.Parse(args[2]);
            }

            TextReader textReader = File.OpenText(csvFilePath);
        
            var csv = new CsvReader(textReader);
            csv.Configuration.RegisterClassMap<OutlookComCsvMap>();
            var records = csv.GetRecords<OutlookComCsvContact>();

            OutlookComToVcardConverter comToVcardConverter = new OutlookComToVcardConverter();

            IEnumerable<OutlookComCsvContact> l = records.AsEnumerable();       
            IEnumerable<VCard> vcards = comToVcardConverter.ConvertOutlookComCsvContacts(l, true);

            string serialized = vcards.Serialize();
            string path = Path.Combine(outFolder, "OutlookConverted.vcf");
            File.WriteAllText(path, serialized);

            Console.WriteLine("finished.");
        }
    }
}

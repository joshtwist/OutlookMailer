using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using CsvHelper;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using CommandLineParser.Arguments;


namespace OutlookMailer
{
    class Program
    {
        static void Main(string[] args)
        {
            var parser = new CommandLineParser.CommandLineParser();
            var send = new SwitchArgument('s', "send", "Whether to automatically send the messages or not", false);
            var template = new ValueArgument<string>('t', "template", "The path to the html template (title = subject)");
            var csv = new ValueArgument<string>('c', "csv", "The path to the CSV file");
            parser.Arguments.Add(template);
            parser.Arguments.Add(csv);
            parser.Arguments.Add(send);

            try
            {
                parser.ParseCommandLine(args);
                parser.ShowParsedArguments();
            }
            catch (System.Exception exc)
            {
                Console.WriteLine(exc.Message);
                return;
            }


            var reader = new CsvReader(new StreamReader(new FileStream(csv.Value, FileMode.Open)));
            var content = File.ReadAllText(template.Value);

            var title = XElement.Parse(content).Descendants("title").First().Value;

            while (reader.Read())
            {
                var body = content;
                foreach (var f in reader.FieldHeaders)
                {
                    body = body.Replace("{" + f.ToLower() + "}", reader.GetField<string>(f));
                }

                var app = new Application();
                var ns = app.GetNamespace("MAPI");
                ns.Logon(null, null, false, false);
                var outbox = ns.GetDefaultFolder(OlDefaultFolders.olFolderOutbox);
                _MailItem message = app.CreateItem(OlItemType.olMailItem);
                message.To = reader.GetField<string>(reader.FieldHeaders.Single(fh => string.Equals(fh, "email", StringComparison.InvariantCultureIgnoreCase)));
                message.Subject = title;
                message.BodyFormat = OlBodyFormat.olFormatHTML;
                message.HTMLBody = body;
                message.SaveSentMessageFolder = outbox;
                if (send.Value)
                {
                    message.Send();
                }
                else
                {
                    message.Save();
                }
            }
        }
    }
}

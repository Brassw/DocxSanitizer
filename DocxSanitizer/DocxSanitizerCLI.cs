using System;
using System.IO;
using System.Collections.Generic;
using Serilog;

namespace DocxSanitizer
{
    class DocxSanitizerCLI
    {
        static void Main(string[] args)
        {
            var log = new LoggerConfiguration()
                .MinimumLevel.Fatal()
                .WriteTo.Console()
                .CreateLogger();
            Log.Logger = log;

            if(args.Length > 0)
            {
                try
                {
                    var files = parseArgs(args);

                    if(log.IsEnabled(Serilog.Events.LogEventLevel.Debug)) {
                        log.Debug("Parsed list of files:");
                        foreach(var file in files)
                        {
                            log.Debug("  " + file);
                        }
                    }

                    Console.WriteLine("Sanitizing total " + files.Count + " files.");
                    DocxSanitizer sanitizer = new DocxSanitizer();
                    sanitizer.Sanitize(files);
                    Console.Write("Finished sanitation. ");
                    Console.WriteLine(sanitizer.Summary.formatSummary());
                }
                catch(Exception e)
                {
                    Console.Error.WriteLine(e.Message);
                    Console.Error.WriteLine("Halting process.\n");
                    printUsage();
                }
            }
            else
            {
                printUsage();
            }
        }

        static List<string> parseArgs(string[] files)
        {
            var list = new List<string>();
            foreach(var file in files)
            {
                // Parse wildcard expressions
                if(file.Contains("*"))
                {
                    int lastIndex = file.LastIndexOfAny(new char[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar });
                    var path = file.Substring(0, lastIndex);
                    var wildcard = file.Substring(lastIndex + 1);
                    Log.Verbose("Split expression '{file}' into '{path}' and '{wildcard}'", file, path, wildcard);
                    string[] globbedFiles = Directory.GetFiles(@path, wildcard);
                    list.AddRange(globbedFiles);
                }
                // Use single file
                else
                {
                    if(File.Exists(file))
                    {
                        list.Add(file);
                    }
                    else
                    {
                        throw new Exception("Unable to parse expression: " + file);
                    }
                }
            }
            return list;
        }

        static void printUsage()
        {
            Console.WriteLine("Docx Sanitizer removes metadata from .docx files.\n");
            Console.WriteLine("Usage:");
            Console.WriteLine("    DocxSanitizer [files]\n");
            Console.WriteLine("Parameters:");
            Console.WriteLine("    [files]    List of docx files to be sanitized, separated by spaces. Wildcard expressions, e.g. *.docx are supported.");
        }
    }
}

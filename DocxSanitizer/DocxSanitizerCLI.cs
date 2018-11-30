using System;
using System.IO;
using System.Collections.Generic;

namespace DocxSanitizer
{
    class DocxSanitizerCLI
    {
        private static Logger log;

        static void Main(string[] args)
        {
            log = new Logger();

            if(args.Length > 0)
            {
                try
                {
                    var files = ParseArgs(args);

                    if(log.IsDebugEnabled()) {
                        log.Debug("Parsed list of files:");
                        foreach(var file in files)
                        {
                            log.Debug("  " + file);
                        }
                    }

                    Console.WriteLine("Sanitizing total " + files.Count + " files.");
                    DocxSanitizer sanitizer = new DocxSanitizer();
                    foreach(var file in files) {
                        try
                        {
                            sanitizer.Sanitize(file);
                            Console.WriteLine("  Processed file '" + file + "'.");
                        }
                        catch(Exception e)
                        {
                            Console.Error.WriteLine("  Error processing file '" + file + "': " + e.Message);
                            log.Debug(e.ToString());
                        }
                    }
                    Console.Write("Finished sanitation. ");
                    Console.WriteLine(sanitizer.Summary.FormatSummary());
                }
                catch(Exception e)
                {
                    Console.Error.WriteLine(e.Message);
                    Console.Error.WriteLine("Halting process.\n");
                    Console.Error.WriteLine("Stacktrace:\n" + e);
                }
            }
            else
            {
                PrintUsage();
            }
        }

        static List<string> ParseArgs(string[] filepaths)
        {
            var list = new List<string>();
            foreach(var filepath in filepaths)
            {
                // Parse wildcard expressions
                if(filepath.Contains("*"))
                {
                    var path = ".";
                    int lastIndex = filepath.LastIndexOfAny(new char[] { Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar });
                    if(lastIndex != -1)
                    {
                        path = filepath.Substring(0, lastIndex);
                    }
                    var wildcard = filepath.Substring(lastIndex + 1);
                    log.Verbose("Split expression '{0}' into '{1}' and '{2}'", filepath, path, wildcard);
                    string[] globbedFiles = Directory.GetFiles(@path, wildcard);
                    list.AddRange(globbedFiles);
                }
                // Use single file
                else
                {
                    if(File.Exists(filepath))
                    {
                        list.Add(filepath);
                    }
                    else
                    {
                        throw new Exception("Unable to parse expression: " + filepath);
                    }
                }
            }
            return list;
        }

        static void PrintUsage()
        {
            Console.WriteLine("DocxSanitizer removes metadata from .docx files.\n");
            Console.WriteLine("Usage:");
            Console.WriteLine("    DocxSanitizer [files]\n");
            Console.WriteLine("Parameters:");
            Console.WriteLine("    [files]    List of docx files to be sanitized, separated by spaces. Wildcard expressions, e.g. *.docx are supported.");
        }
    }
}

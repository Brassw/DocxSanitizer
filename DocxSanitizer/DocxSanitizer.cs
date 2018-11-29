using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace DocxSanitizer
{
    public class DocxSanitizer
    {
        // Overwriting these causes problems 
        static private string[] ignoredKeys = new string[] { "dcterms" };
        private Logger log;
        public Summary Summary;

        public DocxSanitizer()
        {
            log = new Logger();
            Summary = new Summary();
        }

        public void Sanitize(IEnumerable<string> filepaths)
        {
            foreach(var filepath in filepaths)
            {
                Sanitize(filepath);
            }
        }

        public void Sanitize(string filepath)
        {
            try
            {
                using(var document = DocX.Load(filepath))
                {
                    log.Debug("Processing file {0}", filepath);
                    foreach(var keyVal in document.CoreProperties)
                    {
                        bool ignore = false;
                        foreach(var ignoredKey in ignoredKeys)
                        {
                            if(keyVal.Key.Contains(ignoredKey))
                            {
                                ignore = true;
                            }
                        }
                        if(ignore)
                        {
                            log.Verbose("  Ignoring core property {0}={1}.", keyVal.Key, keyVal.Value);
                        }
                        else {
                            log.Verbose("  Overwriting core property {0}={1} with blank value.", keyVal.Key, keyVal.Value);
                            document.AddCoreProperty(keyVal.Key, "");
                        }
                    }
                    foreach(var keyVal in document.CustomProperties)
                    {
                        log.Verbose("  Removing custom property {0}={1}.", keyVal.Key, keyVal.Value);
                        document.CustomProperties.Remove(keyVal.Key);
                    }
                    document.Save();
                    Console.WriteLine("  Processed file '" + filepath + "'.");
                    Summary.FileCount++;
                }
            }
            catch(Exception e)
            {
                Summary.ErrorCount++;
                Console.Error.WriteLine("  Error processing file '" + filepath + "': " + e.Message);
                log.Debug(e.ToString());
            }
        }
    }
}

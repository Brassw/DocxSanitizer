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
        // Overwriting these properties might cause problems 
        private static readonly string[] ignoredProperties = new string[] { "dcterms" };
        private Logger log;
        public Summary Summary;

        public DocxSanitizer()
        {
            log = new Logger();
            Summary = new Summary();
        }

        public void Sanitize(string filepath)
        {
            try
            {
                using(var document = DocX.Load(filepath))
                {
                    log.Debug("Processing file {0}", filepath);
                    foreach(var property in document.CoreProperties)
                    {
                        bool ignore = false;
                        foreach(var ignoredProperty in ignoredProperties)
                        {
                            if(property.Key.Contains(ignoredProperty))
                            {
                                ignore = true;
                            }
                        }
                        if(ignore)
                        {
                            log.Verbose("  Ignoring core property {0}={1}.", property.Key, property.Value);
                        }
                        else {
                            log.Verbose("  Overwriting core property {0}={1} with blank value.", property.Key, property.Value);
                            document.AddCoreProperty(property.Key, "");
                        }
                    }
                    document.Save(); 
                    Summary.FileCount++;
                }
            }
            catch(Exception e)
            {
                Summary.ErrorCount++;
                throw e;
            }
        }
    }
}

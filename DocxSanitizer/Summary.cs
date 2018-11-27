using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxSanitizer
{
    public class Summary
    {
        public long FileCount = 0;
        public long ErrorCount = 0;

        public String formatSummary()
        {
            return "Sanitized successfully " + FileCount + " files, " + ErrorCount + " files with errors.";
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxSanitizer
{
    class Logger
    {
        private bool debugEnabled = false;

        public Logger()
        {
            try
            {
                var debugEnv = System.Environment.GetEnvironmentVariable("DOCX_SANITIZER_DEBUG");
                if("true".Equals(debugEnv))
                {
                    debugEnabled = true;
                }
            }
            catch(Exception e) { }
        }

        public void Debug(string msg, params object[] vars)
        {
            if(debugEnabled)
            {
                Console.WriteLine("DBG: " + string.Format(msg, vars));
            }
        }

        public void Verbose(string msg, params object[] vars)
        {
            Debug(msg, vars);
        }

        public bool IsDebugEnabled()
        {
            return debugEnabled;
        }
    }
}

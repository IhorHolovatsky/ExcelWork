using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;

namespace COFCO.Forms
{
    public class FormsLogger
    {
        public static Logger FormsLoggerInstance = LogManager.GetLogger("MainWindow");
    }
}

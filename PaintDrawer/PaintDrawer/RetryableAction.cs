using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PaintDrawer
{
    class RetryableAction
    {
        private Action action;
        private string actionName;
        private const int DEFAULT_RETRIES = 2;
        public RetryableAction(Action a, string actionName)
        {
            this.action = a;
            this.actionName = actionName;
        }

        public bool Execute(int numRetries = DEFAULT_RETRIES)
        {
            if (numRetries == 0)
            {
                return false;
            }

            try
            {
                Stuff.WriteConsoleMessage("Running " + actionName);
                action.Invoke();
                return true;
            }
            catch(Exception ex)
            {
                Stuff.WriteConsoleError(ex.ToString());
                Stuff.WriteConsoleError("Unable to complete " + actionName + "; trying again...");
                return Execute(numRetries - 1);
            }
        }
    }
}

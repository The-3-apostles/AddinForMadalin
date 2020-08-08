using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace MacroCs
{
    class Program
    {
        static void Main(string[] args)
        {
            SolidEdgeFramework.Application application = null;
            Type type = null;
            try
            {
                // Get the type from the Solid Edge ProgID
                type = Type.GetTypeFromProgID("SolidEdge.Application");
                // Start Solid Edge
                application = (SolidEdgeFramework.Application)
                Activator.CreateInstance(type);
                // Make Solid Edge visible
                application.Visible = true;
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (application != null)
                {
                    Marshal.ReleaseComObject(application);
                    application = null;
                }
            }
        }
    }
}

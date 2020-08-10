using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SolidEdgeFramework;
using SolidEdgePart;
using System.Runtime.InteropServices;

namespace MacroIpad
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SolidEdgeFramework.Application application = null;
            try
            {
                // Connect to a running instance of Solid Edge
                application = (SolidEdgeFramework.Application)
                    Marshal.GetActiveObject("SolidEdge.Application");

                SolidEdgeFramework.SolidEdgeDocument document = null;
                document = (SolidEdgeFramework.SolidEdgeDocument)
                   application.ActiveDocument;

                SolidEdgeFramework.Window window = null;
                window = (SolidEdgeFramework.Window)
                    application.ActiveWindow;

                SolidEdgeFramework.View p = null;
                p = window.View;

                double a, b, c, h, f, g;
                p.GetModelRange(out a, out b, out c, out h, out f, out g);
                Console.WriteLine(a);
                Console.WriteLine(b);
                Console.WriteLine(c);
                Console.WriteLine(h);
                Console.WriteLine(f);
                Console.WriteLine(g);
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

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
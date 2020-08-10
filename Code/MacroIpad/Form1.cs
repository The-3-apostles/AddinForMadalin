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
using SolidEdgeGeometry;
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
            SolidEdgeFramework.SolidEdgeDocument document = null;
            SolidEdgePart.PartDocument partDocument = null;
            SolidEdgePart.Models models = null;
            SolidEdgePart.Model model = null;
            SolidEdgeGeometry.Body body = null;

            SolidEdgeFramework.PropertySets propertySets = null;
            SolidEdgeFramework.Properties propertiesSummary = null;
            SolidEdgeFramework.Property title = null;
            SolidEdgeFramework.Property itemInfo = null;

            SolidEdgeFramework.Properties propertiesProject = null;
            SolidEdgeFramework.Properties documentNumber = null;

            try
            {
                // Attempt to connect to a running instance of Solid Edge.
                application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
                document = application.ActiveDocument as SolidEdgeFramework.SolidEdgeDocument;
                // partDocument = application.ActiveDocument as SolidEdgePart.PartDocument;

                if (document != null)
                {
                    //models = partDocument.Models;
                    //model = models.Item(1);
                    //body = (SolidEdgeGeometry.Body)model.Body;

                    //var MinRangePoint = Array.CreateInstance(typeof(double), 0);
                    //var MaxRangePoint = Array.CreateInstance(typeof(double), 0);

                    //body.GetRange(ref MinRangePoint, ref MaxRangePoint);

                    //foreach (var item in MinRangePoint)
                    //    Console.WriteLine((Double)item * 1000.0);
                    //Console.WriteLine();
                    //foreach (var item in MaxRangePoint)
                    //    Console.WriteLine((Double)item * 1000.0);

                    SolidEdgeFramework.Properties group = null;
                    SolidEdgeFramework.Property field = null;

                    propertySets = (PropertySets)document.Properties;
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\Documents\fisier.txt"))
                    {
                        for (int i = 1; i <= propertySets.Count; ++ i)
                        {
                            group = propertySets.Item(i);
                            file.WriteLine(group.Name);
                            for (int j = 1; j <= group.Count; ++ j)
                            {
                                field = group.Item(j);
                                file.WriteLine(field.Name);
                                Marshal.ReleaseComObject(field);
                            }
                            file.WriteLine();
                            Marshal.ReleaseComObject(group);
                        }
                    }
                    propertiesSummary = propertySets.Item("SummaryInformation");
                    title = propertiesSummary.Item("Title");

                    title.set_Value("TEST");
                    
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                Marshal.ReleaseComObject(propertySets);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
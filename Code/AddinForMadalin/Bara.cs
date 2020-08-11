using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SolidEdgeConstants;
using SolidEdgeFramework;

namespace AddinForMadalin
{
    class Bara
{
        public static void Run()
        {
            SolidEdgeFramework.Application application = null;
            SolidEdgeFramework.SolidEdgeDocument document = null;
            SolidEdgePart.PartDocument partDocument = null;
            SolidEdgePart.Models models = null;
            SolidEdgePart.Model model = null;
            SolidEdgeGeometry.Body body = null;
            application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
            document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
            partDocument = (SolidEdgePart.PartDocument)application.ActiveDocument;

            double[] dims = {6, 8, 10, 12, 13.5, 14, 15, 16, 18, 20, 22, 25, 28, 30, 32, 34, 35, 36, 38, 40, 42, 45, 48, 50, 53, 56, 60, 65, 70, 75, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 200};
            double tolerance=1.05;

            if (document != null)
            {
                PropertySets propertySets = (PropertySets)document.Properties;
                SolidEdgeFramework.Properties customProperties = propertySets.Item("Custom");

                body = (SolidEdgeGeometry.Body)partDocument.Models.Item(1).Body;

                var MinRangePoint = Array.CreateInstance(typeof(double), 3);
                var MaxRangePoint = Array.CreateInstance(typeof(double), 3);

                body.GetExactRange(ref MinRangePoint, ref MaxRangePoint);

                double X = (double)MaxRangePoint.GetValue(0) - (double)MinRangePoint.GetValue(0);
                double Y = (double)MaxRangePoint.GetValue(1) - (double)MinRangePoint.GetValue(1);
                double Z = (double)MaxRangePoint.GetValue(2) - (double)MinRangePoint.GetValue(2);

                X *= 1000;
                Y *= 1000;
                Z *= 1000;

                Console.WriteLine(X);
                Console.WriteLine(Y);
                Console.WriteLine(Z);

                double diam=0;

                double lim = Math.Max(X, Y);

                for (int i=0; i<dims.Length; ++i)
                {
                    if (dims[i] >= lim )
                    {
                        diam = dims[i];
                        break;
                    }
                }

                if (diam == 0)
                {
                    throw new ArgumentException("Diamater greater than 200mm, not supported!");
                }

                string Output;
                Output = $"Rd_{diam}";

                int Consum;
                string ConsumString;

                Consum = (int)Math.Floor(Z * tolerance);
                Consum = Consum - Consum % 10 + 10;
                ConsumString = $"{Consum}";

                Property semifabricat = customProperties.Item("Semifabricat");
                //Property consum = customProperties.Item("Consum");
                Property udm = customProperties.Item("Unitate de masura");

                semifabricat.set_Value(Output);
                //consum.set_Value(Consum);
                udm.set_Value("mm");

                Console.WriteLine("Semifabricat: " + Output);
                Console.WriteLine("Consum: " + ConsumString +"mm");
            }
        }
    }
}

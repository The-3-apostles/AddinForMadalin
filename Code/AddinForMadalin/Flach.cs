using System;
using SolidEdgeFramework;
using SolidEdgeGeometry;
using SolidEdgePart;
using System.Runtime.InteropServices;

namespace AddinForMadalin
{
    static class Flach
    {
        public static void run()
        {
            //TODO: Ce facem daca e din OL 37? 
            const int tolerance = 10;
            try
            {
                SolidEdgeFramework.Application application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
                PartDocument partDocument = (PartDocument)application.ActiveDocument;
                Body body = (Body)partDocument.Models.Item(1).Body;

                var minCoord = Array.CreateInstance(typeof(double), 0);
                var maxCoord = Array.CreateInstance(typeof(double), 0);

                body.GetExactRange(ref minCoord, ref maxCoord);

                int[] dim = new int[3];

                for (int i = 0; i < 3; ++i)
                {
                    Console.WriteLine($"{minCoord.GetValue(i)}, {maxCoord.GetValue(i)}");
                    dim[i] = tolerance * (int)(((double)maxCoord.GetValue(i) - (double)minCoord.GetValue(i)) * 1000 / tolerance + 1);
                }

                Array.Sort(dim);

                SolidEdgeDocument document = document = (SolidEdgeDocument)application.ActiveDocument;
                PropertySets propertySets = (PropertySets)document.Properties;
                SolidEdgeFramework.Properties customProperties = propertySets.Item("Custom");

                Property semifabricat = customProperties.Item("Semifabricat");

                semifabricat.set_Value($"#_{dim[0]}x{dim[1]}x{dim[2]}");

                Property consum = customProperties.Item("Consum");

                consum.set_Value(1);

                Property udm = customProperties.Item("Unitate de masura");

                udm.set_Value("buc");
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                //Marshal.ReleaseComObject(propertySets);
            }
        }
    }
}
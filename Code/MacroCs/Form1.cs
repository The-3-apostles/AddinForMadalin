using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using SolidEdgeFramework;
using SolidEdgePart;
using SolidEdgeCommunity;

namespace MacroCs
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SolidEdgeFramework.Application seApplication = null;
            SolidEdgeFramework.Documents seDocuments = null;
            SolidEdgePart.PartDocument sePartDocument = null;

            Type type = null;
            try
            {
                seApplication = SolidEdgeCommunity.SolidEdgeUtils.Connect(true);
                seDocuments = seApplication.Documents;
                sePartDocument = (PartDocument)seDocuments.Add("SolidEdge.PartDocument");

            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                sePartDocument = null;
                seDocuments = null;
                seApplication = null;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}

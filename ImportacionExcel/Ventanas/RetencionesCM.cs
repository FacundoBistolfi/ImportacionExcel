using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ImportacionExcel
{
    public partial class RetencionesCM : Form
    {

        private Excel.Application app;
        private Workbook libro;
        private Worksheet hoja;

        public RetencionesCM()
        {
            InitializeComponent();
        }


        private void label1_Click(object sender, EventArgs e)
        {
            
        }

        private void RetencionesCM_Load(object sender, EventArgs e)
        {

        }

        private void btn_Abrir_Click(object sender, EventArgs e)
        {
            if (this.openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    app = new Excel.Application();
                    //Muestro el path del archivo elegido en el text box
                    tb_Path.Text = openFileDialog1.FileName;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        private void RetencionesCM_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }


        private void button2_Click(object sender, EventArgs e)
        {
            //Abro el archivo
            libro = app.Workbooks.Open(openFileDialog1.FileName, Type.Missing, true, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);

            //Worksheet sheet = (Worksheet)wb.Sheets["Nombre de la hoja"];

            foreach (Worksheet hojas in libro.Sheets)
                comboBox1.Items.Add(hojas.Name);
            panel2.Show();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<Rectangulo> rectangulos = new List<Rectangulo>();
            bool eof = false;
            object ob, oh;
            string b, h, testo;
            double sumAreas = 0;
            int i = 1;
            do
            {
                ob = hoja.Cells[i, 1].Value;
                oh = hoja.Cells[i, 2].Value;
                if (ob == null || oh == null)
                    eof = true;
                else
                    rectangulos.Add(new Rectangulo(Double.Parse(ob.ToString()), Double.Parse(oh.ToString())));
                i++;
            } while (!eof);

            libro.Close();

            testo = "La puta madre\n";
            testo += "que te pario\n";

            foreach (Rectangulo r in rectangulos)
            {
                testo += ("b=" + r.getBase().ToString() + " h=" + r.getAltura().ToString() + "\n");
            }

            MessageBox.Show(testo, "Hola");
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //hoja = libro.Sheets[comboBox1.SelectedText];
            hoja = libro.Sheets.get_Item(comboBox1.SelectedIndex + 1);
            panel3.Show();
        }


    }
}

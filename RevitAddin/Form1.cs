using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAddin
{
    public partial class Form_Carimbo : Form
    {
        public Form_Carimbo()
        {
            InitializeComponent();
        }

        private void btn_Browser_Click(object sender, EventArgs e)
        {
            //Windows Form para procurar o arquivo excel
            try
            {
                System.Windows.Forms.OpenFileDialog openDialog = new System.Windows.Forms.OpenFileDialog();
                openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer);
                openDialog.Filter = "Excel Files (*.xlsx) | *.xlsx";


                if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    CreateSheet.MinhasVariaveis.path = openDialog.FileName;
                    txtbox_Path.Text = CreateSheet.MinhasVariaveis.path;
                }
                
            }
            catch
            {
                TaskDialog.Show("Erro", "Não foi possivel executar o comando");
            }
        }

        private void btn_Cancel_Click(object sender, EventArgs e)
        {
            CreateSheet.MinhasVariaveis.stats = "Cancelado";
            this.Close();
        }

        private void btn_Mec_Click(object sender, EventArgs e)
        {
            CreateSheet.MinhasVariaveis.stats = "01";
            this.Hide();
        }

        private void btn_Plu_Click(object sender, EventArgs e)
        {
            CreateSheet.MinhasVariaveis.stats = "02";
            this.Hide();
        }

        private void btn_Ele_Click(object sender, EventArgs e)
        {
            CreateSheet.MinhasVariaveis.stats = "03";
            this.Hide();
        }

        private void btn_LVS_Click(object sender, EventArgs e)
        {
            CreateSheet.MinhasVariaveis.stats = "04";
            this.Hide();
        }
    }
}

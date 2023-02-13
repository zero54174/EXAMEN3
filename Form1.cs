using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;
using Microsoft.Office.Interop.Excel;
using objExcel = Microsoft.Office.Interop.Excel;

namespace EXAMEN3
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string ruta = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        public void limpiar()
        {

            txtCed.Text = string.Empty;
            txtCod.Text = "0";
            txtSalario.Text = "0";
            txtNombre.Text = string.Empty;
            txtApellido.Text = string.Empty;
            cboSexo.Text = string.Empty;
            txtCargo.Text = string.Empty;
            cboSal.Text = string.Empty;
            cboActivo.Text = string.Empty;
            txtCed.Focus();

        }
        public void limpiar2()
        {
            txtCed.Text = string.Empty;
            txtCod.Text = "0";
            txtSalario.Text = "0";
            txtNombre.Text = string.Empty;
            txtApellido.Text = string.Empty;
            cboSexo.Text = string.Empty;
            txtCargo.Text = string.Empty;
            cboSal.Text = string.Empty;
            cboActivo.Text = string.Empty;
            txtCed.Focus();

            int f;
            f = dglab.RowCount;
            for (int i = f - 1; i >= 0; i--)
            {
                dglab.Rows.RemoveAt(i);
            }
        }
        public void clean()
        {

            int f = dglab.RowCount;
            for (int i = f - 1; i >= 0; i--)
            {
                dglab.Rows[i].DefaultCellStyle.BackColor = Color.White;
            }
            for (int i = f - 1; i >= 0; i--)
            {
                this.dglab.CurrentCell = null;
                this.dglab.Rows[i].Visible = true;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {

            panel1.Visible = false;
            dglab.AllowUserToAddRows = false;
            dgvlab1.AllowUserToAddRows = false;
            dgvlab2.AllowUserToAddRows = false;
            btnGuardar.Enabled = false;
        }
        int strFila = 0;
        private void btnBuscar_Click(object sender, EventArgs e)
        {
            if (txtBuscar.Text == "")
            {
                dglab.DefaultCellStyle.BackColor = Color.White;
            }
            else
            {
                foreach (DataGridViewRow Row in dglab.Rows)
                {


                    strFila = Convert.ToInt32(Row.Index.ToString());
                    string valor = Convert.ToString(Row.Cells["cedula"].Value);
                    string valor2 = Convert.ToString(Row.Cells["codigo"].Value);
                    string valor3 = Convert.ToString(Row.Cells["intactivoOactivo"].Value);
                    if (valor == this.txtBuscar.Text || valor2 == this.txtBuscar.Text|| valor3 == this.txtBuscar.Text)
                    {
                        this.dglab.CurrentCell = null;

                        int f = dglab.RowCount;
                        for (int i = f - 1; i >= 0; i--)
                        {
                            this.dglab.CurrentCell = null;
                            this.dglab.Rows[i].Visible = false;
                            this.dglab.Rows[strFila].Visible = true;
                        }
                        dglab.Rows[strFila].DefaultCellStyle.BackColor = Color.Green;
                        chBuscar.Checked = false;
                        txtBuscar.Text = String.Empty;
                        txtBuscar.Enabled = false;
                    }
                }
            }
        }

        private void chBuscar_CheckedChanged(object sender, EventArgs e)
        {
            if (chBuscar.Checked == true)
            {
                txtBuscar.Enabled = true;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            clean();
        }

        private void cbSalon_1_CheckedChanged(object sender, EventArgs e)
        {

            btnGuardar.Enabled = true;
        }

        private void cbSalon_2_CheckedChanged(object sender, EventArgs e)
        {

            btnGuardar.Enabled = true;
        }

        private void cbSalon_3_CheckedChanged(object sender, EventArgs e)
        {

            btnGuardar.Enabled = true;
        }

        private void btnAgre_Click(object sender, EventArgs e)
        {
            try
            {
                string nombre, apellido, cargo, sexo, activo, cedula, puesto,desempeño;
                string cuentasAC, valoresPA = "", val = "", vala = "";
                double codigo,salario,h,isa=0;
                cedula = txtCed.Text;
                codigo = double.Parse(txtCod.Text);
                salario = double.Parse(txtSalario.Text);
                h = double.Parse(txtHoras.Text);
                nombre = txtNombre.Text;
                apellido = txtApellido.Text;
                sexo = cboSexo.Text;
                cargo = txtCargo.Text;
                puesto = cboSal.Text;
                activo = cboActivo.Text;

                if (codigo <= 0||salario<=8000|| salario >=500000)
                {

                }
                else
                {

                    double x, y, w, z=0;
                    //cuentas = String.Format("{0}", cuen);
                    valoresPA = String.Format("{0}", codigo);
                    val = String.Format("{0}", salario);
                    x = salario / 30;
                    y = x / 8;
                    w = (y * h) * 2;
                    isa = salario * 0.65;
                    vala = String.Format("{0}", isa);



                }
                //intactivoOactivo
                string[] fila = new string[10];
                fila[0] = cedula;
                fila[1] = valoresPA;
                fila[2] = nombre;
                fila[3] = apellido;
                fila[4] = sexo;
                fila[5] = cargo;
                fila[6] = val;
                fila[7] = vala;
                fila[8] = puesto;
                fila[9] = activo;


                if (cedula == "" || valoresPA == "" || nombre == "" || apellido == "" || sexo == "" || cargo == "" || puesto == "" || activo == ""|| val == "")
                {
                    //MessageBox.Show("caracter no valido");
                }
                else
                {
                    if (cboActivo.Text == "activo") 
                    {
                        dglab.Rows.Add(fila);
                        dgvlab1.Rows.Add(fila);
                        limpiar();
                    }
                    else
                    {
                        if (cboActivo.Text == "inactivo")
                        {
                            dglab.Rows.Add(fila);
                            dgvlab2.Rows.Add(fila);
                            limpiar();
                        }
                        else
                        {
                          
                        }
                    }

                }
            }
            catch
            {
                // MessageBox.Show("los caracter no son validos");
            }
        }

        private void dglab_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {
                DataGridViewSelectedCellCollection cell = dglab.SelectedCells;
                DataGridViewSelectedRowCollection rows = dglab.SelectedRows;
                IEnumerator iter = cell.GetEnumerator(); bool sw = false;
                while (iter.MoveNext() && !sw)
                {
                    DataGridViewTextBoxCell dgvtxt = (DataGridViewTextBoxCell)iter.Current;
                    int columna = dgvtxt.ColumnIndex;
                    int fila = dgvtxt.RowIndex;
                    txtCed.Text = Convert.ToString(dglab[0, fila].Value);
                    txtCod.Text = Convert.ToString(dglab[1, fila].Value);
                    txtNombre.Text = Convert.ToString(dglab[2, fila].Value);
                    txtApellido.Text = Convert.ToString(dglab[3, fila].Value);
                    cboSexo.Text = Convert.ToString(dglab[4, fila].Value);
                    txtCargo.Text = Convert.ToString(dglab[5, fila].Value);
                    txtSalario.Text = Convert.ToString(dglab[6, fila].Value);


                }
                int num;
                num = dglab.CurrentRow.Index;
                dglab.Rows.RemoveAt(num);
            }
            catch (Exception ex)
            {
                MessageBox.Show("no hay elementos");
            }
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            objExcel.Application objAplicacion = new objExcel.Application();
            Workbook objLibro = objAplicacion.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet objHoja = (Worksheet)objAplicacion.ActiveSheet;

            objAplicacion.Visible = false;

            foreach (DataGridViewColumn columna in dglab.Columns)
            {
                objHoja.Cells[1, columna.Index + 1] = columna.HeaderText;
                foreach (DataGridViewRow fila in dglab.Rows)
                {
                    objHoja.Cells[fila.Index + 2, columna.Index + 1] = fila.Cells[columna.Index].Value;
                }
            }

            objLibro.Close();
            limpiar2();
            MessageBox.Show("se guardo correctamente");
        }

        private void btnValidar_Click(object sender, EventArgs e)
        {
            panel1.Visible = true;
        }

        private void btnOcultar_Click(object sender, EventArgs e)
        {
            panel1.Visible=false;
        }
    }
}

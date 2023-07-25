using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        int t;
        int b = 0;
        bool isEmpty = false;

        public Form1()
        {



            InitializeComponent();


            void BindMyComboBox1()
            {
                comboBox1.SelectedIndexChanged -= comboBox1_SelectedIndexChanged;
                string connectionStr = "Data Source = (LocalDB)\\MSSQLLocalDB; Integrated Security = True";
                SqlConnection cn = new SqlConnection(connectionStr);
                SqlDataAdapter da = new SqlDataAdapter();
                SqlCommand cmd = new System.Data.SqlClient.SqlCommand();
                System.Data.DataTable dt = new System.Data.DataTable();
                cmd.CommandText = "Select DISTINCT typee from [TYPE]  ";
                da.SelectCommand = cmd;
                da.SelectCommand.Connection = cn;
                da.Fill(dt);
                comboBox1.DisplayMember = "typee";
                comboBox1.ValueMember = "typee";
                comboBox1.DataSource = dt;

                comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;


            }
            BindMyComboBox1();









        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (!string.IsNullOrEmpty(comboBox1.Text))
            {
                string connectionstr2 = "Data Source = (LocalDB)\\MSSQLLocalDB; Integrated Security = True";
                SqlConnection cn2 = new SqlConnection(connectionstr2);
                SqlDataAdapter da2 = new SqlDataAdapter();
                SqlCommand cmd2 = new System.Data.SqlClient.SqlCommand();
                System.Data.DataTable dt2 = new System.Data.DataTable();
                string a = comboBox1.Text;
                cmd2.CommandText = $"SELECT section FROM Type where typee = '{a}' ";
                da2.SelectCommand = cmd2;
                da2.SelectCommand.Connection = cn2;
                da2.Fill(dt2);
                comboBox2.Enabled = true;
                comboBox2.DisplayMember = "section";
                comboBox2.ValueMember = "section";
                comboBox2.DataSource = dt2;


            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Visible = true;
            dataGridView1.RowHeadersVisible = false;
            label2.Text = comboBox1.Text + comboBox2.Text;
            void connecttosql()
            {
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();

                SqlConnection cn = new SqlConnection();
                SqlCommand cmd = new SqlCommand();
                cn.ConnectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;Integrated Security=True"; 
                cmd.Connection = cn;
                cmd.Connection.Open();
                cmd.CommandText = $"SELECT lot,item,toleranceMin,toleranceNom,toleranceMax,concernepar,donnees{comboBox1.Text + comboBox2.Text},{comboBox1.Text + comboBox2.Text},Tolerances FROM Tolerance where {comboBox1.Text + comboBox2.Text} != ''"; 
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                dt.Clear();






                da.Fill(dt);


                BindingSource source = new BindingSource();
                source.DataSource = dt;
                dataGridView1.DataSource = dt;


                void loaddatagridview2()
                {
                    dataGridView2.Rows.Clear();
                    dataGridView2.Columns.Clear();

                    var etiquette = new DataGridViewTextBoxColumnEx();
                    etiquette.Name = "etiquette";
                    etiquette.HeaderText = "N°Etiquette";
                    dataGridView2.Columns.Add(etiquette);

                    var etiquetteemballagec = new DataGridViewTextBoxColumnEx();
                    etiquetteemballagec.Name = "EtiquetteEmballageC";
                    etiquetteemballagec.HeaderText = "EtiquettageEmballageC";
                    dataGridView2.Columns.Add(etiquetteemballagec);

                    var etiquetteemballagenc = new DataGridViewTextBoxColumnEx();
                    etiquetteemballagenc.Name = "EtiquetteEmballageNC";
                    etiquetteemballagenc.HeaderText = "EtiquettageEmballageNC";
                    dataGridView2.Columns.Add(etiquetteemballagenc);

                    var controlevisuelC = new DataGridViewTextBoxColumnEx();
                    controlevisuelC.Name = "controlevisuelC";
                    controlevisuelC.HeaderText = "Contrôle visuel C";
                    dataGridView2.Columns.Add(controlevisuelC);

                    var controlevisuelNC = new DataGridViewTextBoxColumnEx();
                    controlevisuelNC.Name = "controlevisuelNC";
                    controlevisuelNC.HeaderText = "Contrôle visuel NC";
                    dataGridView2.Columns.Add(controlevisuelNC);

                    var rapportsiic = new DataGridViewTextBoxColumnEx();
                    rapportsiic.Name = "NRapportSIIC";
                    rapportsiic.HeaderText = "N°Rapport SIIC";
                    dataGridView2.Columns.Add(rapportsiic);

                    var etatdistribution = new DataGridViewTextBoxColumnEx();
                    etatdistribution.Name = "etatdistribution";
                    etatdistribution.HeaderText = "Etat de distribution";
                    dataGridView2.Columns.Add(etatdistribution);

                    var OKNOK = new DataGridViewTextBoxColumnEx();
                    OKNOK.Name = "OKNOK";
                    OKNOK.HeaderText = "OK/NOK";
                    dataGridView2.Columns.Add(OKNOK);


                    
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        dataGridView2.Rows.Add();
                    }
                    dataGridView2.Visible = true;
                    dataGridView2.RowHeadersVisible = false;
                }

                loaddatagridview2();

                void donneesdemesures()
                {
                    List<int> valeurs = new List<int>() { };
                    var result = from DataGridViewRow row in dataGridView1.Rows select row;

                    foreach (var row in result)
                    {

                        string a = row.Cells[$"donnees{comboBox1.Text + comboBox2.Text}"].Value?.ToString();
                        valeurs.Add(Convert.ToInt32(a));

                    }

                    b = valeurs.Max();

                    var testt = new DataGridViewTextBoxColumnEx();
                    testt.Name = "Donnee1";
                    dataGridView1.Columns.Add(testt);
               
                    for (int i = 1; i <= 2; i++)
                    {
                        var donnee = new DataGridViewTextBoxColumnEx();
                        donnee.Name = $"Donnee{i+1}";
                        dataGridView1.Columns.Add(donnee);

                    }

                    var jugementok = new DataGridViewTextBoxColumnEx();
                    jugementok.Name = "Jugement Ok";
                    jugementok.HeaderText = "Jugement Ok";
                    dataGridView1.Columns.Add(jugementok);

                    var jugementnok = new DataGridViewTextBoxColumnEx();
                    jugementnok.Name = "Jugement NOK";
                    jugementok.HeaderText = "Jugement NOK";
                    dataGridView1.Columns.Add(jugementnok);

                    var dblcontrval = new DataGridViewTextBoxColumnEx();
                    dblcontrval.Name = "Double contrôle valeur";
                    dblcontrval.HeaderText = "Double contrôle valeur";
                    dataGridView1.Columns.Add(dblcontrval);

                    var dblcontrok = new DataGridViewTextBoxColumnEx();
                    dblcontrok.Name = "Double contrôle OK/NOK";
                    dblcontrok.HeaderText = "Double contrôle OK / NOK";
                    dataGridView1.Columns.Add(dblcontrok);

                    var newResultat = new DataGridViewTextBoxColumn();
                    newResultat.Name = "newResultat";
                    newResultat.HeaderText = "Résultat Min";
                    dataGridView1.Columns.Insert(12, newResultat);


                    

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.Cells["newResultat"] = new DataGridViewTextBoxCellEx();
                    }


                    var newResultatMed = new DataGridViewTextBoxColumn();
                    newResultatMed.Name = "newResultatMed";
                    newResultatMed.HeaderText = "Résultat Med";
                    dataGridView1.Columns.Insert(13, newResultatMed);




                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.Cells["newResultatMed"] = new DataGridViewTextBoxCellEx();
                    }


                    var newResultatMax = new DataGridViewTextBoxColumn();
                    newResultatMax.Name = "newResultatMax";
                    newResultatMax.HeaderText = "Résultat Max";
                    dataGridView1.Columns.Insert(14, newResultatMax);




                   foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.Cells["newResultatMax"] = new DataGridViewTextBoxCellEx();
                    }


                    var newlot = new DataGridViewTextBoxColumn();
                    newlot.Name = "newlot";
                    newlot.HeaderText = "lot";
                    dataGridView1.Columns.Insert(0, newlot);

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.Cells["newlot"] = new DataGridViewTextBoxCellEx();
                        row.Cells["newlot"].Value = row.Cells["lot"].Value;
                    }

                    dataGridView1.Columns.Remove(dataGridView1.Columns["lot"]);

                    var newitems = new DataGridViewTextBoxColumn();
                    newitems.Name = "newitems";
                    newitems.HeaderText = "item";
                    dataGridView1.Columns.Insert(2, newitems);


                    var newToleranceMin = new DataGridViewTextBoxColumn();
                    newToleranceMin.Name = "newToleranceMin";
                    newToleranceMin.HeaderText = "Tolerance Min";
                    dataGridView1.Columns.Insert(3, newToleranceMin);

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.Cells["newitems"] = new DataGridViewTextBoxCellEx();
                        row.Cells["newitems"].Value = row.Cells["item"].Value;

                    }

                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.Cells["newToleranceMin"] = new DataGridViewTextBoxCellEx();
                        row.Cells["newToleranceMin"].Value = row.Cells["toleranceMin"].Value;

                    }

                    var newToleranceNom = new DataGridViewTextBoxColumn();
                    newToleranceNom.Name = "newToleranceNom";
                    newToleranceNom.HeaderText = "Tolerance Nom";
                    dataGridView1.Columns.Insert(4, newToleranceNom);


                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.Cells["newToleranceNom"] = new DataGridViewTextBoxCellEx();
                        row.Cells["newToleranceNom"].Value = row.Cells["toleranceNom"].Value;

                    }
                    var newToleranceMax = new DataGridViewTextBoxColumn();
                    newToleranceMax.Name = "newToleranceMax";
                    newToleranceMax.HeaderText = "Tolerance Max";
                    dataGridView1.Columns.Insert(5, newToleranceMax);


                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        row.Cells["newToleranceMax"] = new DataGridViewTextBoxCellEx();
                        row.Cells["newToleranceMax"].Value = row.Cells["toleranceMax"].Value;

                    }

                    dataGridView1.Columns.RemoveAt(1);
                    dataGridView1.Columns.Remove("toleranceMin");
                    dataGridView1.Columns.Remove("toleranceNom");
                    dataGridView1.Columns.Remove("toleranceMax");



                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        if (dataGridView1.Rows[i].Cells["newitems"].Value.ToString() == dataGridView1.Rows[i + 1].Cells["newitems"].Value.ToString())
                        {

                           
                            var cell1 = (DataGridViewTextBoxCellEx)dataGridView1.Rows[i].Cells["newitems"];
                            cell1.RowSpan = 2;

                            var cell2 = (DataGridViewTextBoxCellEx)dataGridView1.Rows[i].Cells["newToleranceMin"];
                            cell2.RowSpan = 2;

                            var cell3 = (DataGridViewTextBoxCellEx)dataGridView1.Rows[i].Cells["newToleranceNom"];
                            cell3.RowSpan = 2;

                            var cell4 = (DataGridViewTextBoxCellEx)dataGridView1.Rows[i].Cells["newToleranceMax"];
                            cell4.RowSpan = 2;

                            for (int j = 12; j <= 18; j++)
                            {
                                var cell = (DataGridViewTextBoxCellEx)dataGridView1.Rows[i].Cells[j];
                                cell.RowSpan = 2;
                            }

                            string a = dataGridView1.Rows[i].Cells[$"donnees{comboBox1.Text + comboBox2.Text}"].Value.ToString();

                            int o = int.Parse(a);

                            if (o <= 3)
                            {
                                var cell = (DataGridViewTextBoxCellEx)dataGridView1.Rows[i].Cells["Donnee1"];
                                cell.RowSpan = 2;
                                cell = (DataGridViewTextBoxCellEx)dataGridView1.Rows[i].Cells["Donnee2"];
                                cell.RowSpan = 2;
                                cell = (DataGridViewTextBoxCellEx)dataGridView1.Rows[i].Cells["Donnee3"];
                                cell.RowSpan = 2;
                            }


                            foreach (DataGridViewColumn col in dataGridView2.Columns)
                            {
                                var celll = (DataGridViewTextBoxCellEx)dataGridView2.Rows[i].Cells[col.Index];
                                celll.RowSpan = 2;
                            }



                        }

                    }




                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        string a = row.Cells[$"donnees{comboBox1.Text + comboBox2.Text}"].Value.ToString();
                        int o = int.Parse(a);
                        if (o < 3)
                        {
                            for (int i = o + 1; i <= 3; i++)
                            {
                                row.Cells[$"Donnee{i}"].Style.BackColor = System.Drawing.Color.Gray;

                                row.Cells["newResultat"].Style.BackColor = System.Drawing.Color.Gray;
                                row.Cells["newResultatMed"].Style.BackColor = System.Drawing.Color.Gray;
                                row.Cells["newResultatMax"].Style.BackColor = System.Drawing.Color.Gray;

                                row.Cells[$"Donnee{i}"].ReadOnly = true;
                                row.Cells["newResultat"].ReadOnly = true;
                                row.Cells["newResultatMed"].ReadOnly = true;
                                row.Cells["newResultatMax"].ReadOnly = true;
                            }
                            
                        }
                        else if (o > 3)
                        {
                            if (dataGridView1.Rows[row.Index - 1].Cells["newitems"].Value?.ToString() == row.Cells["newitems"].Value?.ToString())
                            {
                                if (o == 5)
                                {
                                    row.Cells["Donnee3"].Style.BackColor = System.Drawing.Color.Gray;
                                    row.Cells["Donnee3"].ReadOnly = true;
                                }
                                if (o == 4)
                                {
                                    row.Cells["Donnee2"].Style.BackColor = System.Drawing.Color.Gray;
                                    row.Cells["Donnee2"].ReadOnly = true;

                                }

                                

                            }

                        }


                    }

                }
                donneesdemesures();


                dataGridView1.Columns[$"newToleranceMin"].DefaultCellStyle.BackColor = System.Drawing.Color.Wheat;
                dataGridView1.Columns[$"newToleranceNom"].DefaultCellStyle.BackColor = System.Drawing.Color.Wheat;
                dataGridView1.Columns[$"newToleranceMax"].DefaultCellStyle.BackColor = System.Drawing.Color.Wheat;

                dataGridView1.Columns[$"newlot"].ReadOnly = true;
                dataGridView1.Columns[$"newitems"].ReadOnly = true;
                dataGridView1.Columns[$"newToleranceMin"].ReadOnly = true;
                dataGridView1.Columns[$"newToleranceNom"].ReadOnly = true;
                dataGridView1.Columns[$"newToleranceMax"].ReadOnly = true;



            }
            connecttosql();

            void mergelots()
            {
                int deb1 = 0;
                int deb2 = 0;
                int deb3 = 0;
                int fin1 = 0;
                int fin2 = 0;
                int fin3 = dataGridView1.Rows.Count - 1;
                if (!(dataGridView1.Rows[0].Cells["Tolerances"].Value?.ToString() is "3"))
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if ((row.Cells["newlot"].Value?.ToString() is "2") && (deb2 == 0))
                        {
                            deb2 = row.Index;
                            fin1 = row.Index - 1;
                        }
                        if ((row.Cells["newlot"].Value?.ToString() is "3") && (deb3 == 0))
                        {
                            deb3 = row.Index;
                            fin2 = row.Index - 1;
                        }
                    }

                    var cell1 = (DataGridViewTextBoxCellEx)dataGridView1.Rows[deb1].Cells["newlot"];
                    cell1.RowSpan = fin1 + 1;
                    var cell2 = (DataGridViewTextBoxCellEx)dataGridView1.Rows[deb2].Cells["newlot"];
                    cell2.RowSpan = fin2 - deb2 + 1;
                    var cell3 = (DataGridViewTextBoxCellEx)dataGridView1.Rows[deb3].Cells["newlot"];
                    cell3.RowSpan = fin3 - deb3 + 1;
                }
                else
                {
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {

                        if ((row.Cells["newlot"].Value?.ToString() is "3") && (deb3 == 0))
                        {
                            deb3 = row.Index;
                            fin1 = row.Index - 1;
                        }
                    }

                    var cell1 = (DataGridViewTextBoxCellEx)dataGridView1.Rows[deb1].Cells["newlot"];
                    cell1.RowSpan = fin1 + 1;
                    var cell3 = (DataGridViewTextBoxCellEx)dataGridView1.Rows[deb3].Cells["newlot"];
                    cell3.RowSpan = fin3 - deb3 + 1;
                }

            }

            mergelots();

            void creationitemsettolerances()
            {

                dataGridView1.Columns[$"{comboBox1.Text + comboBox2.Text}"].Visible = false;
                dataGridView1.Columns["concernepar"].Visible = false;
                dataGridView1.Columns[$"donnees{comboBox1.Text + comboBox2.Text}"].Visible = false;
                dataGridView1.Columns["Tolerances"].Visible = false;

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    string[] splitter = new string[3];
                    t++;
                    string a = row.Cells[$"{comboBox1.Text + comboBox2.Text}"].Value?.ToString();
                    if (a.Contains('-'))
                    {
                        splitter = a.Split('-');
                        row.Cells["newToleranceMin"].Value = splitter[0];
                        row.Cells["newToleranceNom"].Value = splitter[1];
                        try
                        {
                            row.Cells["newToleranceMax"].Value = splitter[2];
                        }
                        catch (Exception ex) { Debug.WriteLine(ex); }


                    }
                    else
                    {
                        row.Cells["newToleranceNom"].Value = a;

                    }
                }

            }

            creationitemsettolerances();

            void mergingtolerances()
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {

                    if ((string.IsNullOrEmpty(row.Cells["newToleranceMin"].Value?.ToString())) && (string.IsNullOrEmpty(row.Cells["newToleranceMax"].Value?.ToString())))
                    {
                        string a = row.Cells["newToleranceNom"].Value.ToString();
                        var cell = (DataGridViewTextBoxCellEx)row.Cells["newToleranceMin"];
                        cell.ColumnSpan = 3;
                        cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                        cell.Value = a;
                    }
                }
            }
            mergingtolerances();

            void test()
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    foreach (DataGridViewColumn col in dataGridView1.Columns)
                    {
                        if (row.Cells[col.Index].Style.BackColor == System.Drawing.Color.Gray)
                        {
                            row.Cells[col.Index].ReadOnly = true;
                        }
                    }
                }
            }

            test();

            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.Height = (dataGridView1.Rows.Count * dataGridView1.RowTemplate.Height) + dataGridView1.ColumnHeadersHeight;
            dataGridView2.Height = (dataGridView1.Rows.Count * dataGridView1.RowTemplate.Height) + dataGridView1.ColumnHeadersHeight;

            tableLayoutPanel10.Location = new System.Drawing.Point(dataGridView1.Location.X+5, dataGridView1.Location.Y+dataGridView1.Height+5);
            tableLayoutPanel10.Visible = true;
            tableLayoutPanel11.Visible = true;




        }

        private void button1_Click(object sender, EventArgs e)
        {
            isEmpty = false;

            string filename = "C:/Users/guizaoui/Desktop/partage/fichier.xlsx";
            var wb = new ClosedXML.Excel.XLWorkbook(filename);
            var ws = wb.Worksheets.Worksheet("Feuil1");


            


            void listeverif()
            {
                string controle = "Contrôle du : ";
                if (radioButton1.Checked == true)
                {
                    controle += radioButton1.Text;

                }
                else if (radioButton2.Checked == true)
                {
                    controle += radioButton2.Text;

                }
                else if (radioButton3.Checked == true)
                {
                    controle += radioButton3.Text;

                }
                else if (radioButton4.Checked == true)
                {
                    controle += radioButton4.Text;

                }

                ws.Cell("T8").Value = textBox1.Text;
                ws.Cell("T10").Value = textBox2.Text;

                ws.Cell("G8").Value = controle;
            }


            listeverif();


            void entete()
            {
                ws.Cell("G4").Value = comboBox1.Text + comboBox2.Text;
                ws.Cell("B4").Value = textBox9.Text;
                ws.Cell("B5").Value = textBox3.Text;
                ws.Cell("B6").Value = textBox4.Text;
                ws.Cell("B7").Value = textBox5.Text;
                ws.Cell("B8").Value = textBox6.Text;
                ws.Cell("B9").Value = textBox7.Text;
                ws.Cell("B10").Value = textBox8.Text;



            }

            entete();


            void itemsettolerances()
            {


                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {

                    for (int j = 0; j < 5; j++)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Visible == true)
                        {
                            ws.Cell(i + 14, j + 2).Value = dataGridView1.Rows[i].Cells[j].Value?.ToString();


                            if ((j + 2 >= 4) && (j + 2 <= 6))
                            {
                                ws.Cell(i + 14, j + 2).Style.Fill.SetBackgroundColor(XLColor.Wheat);
                                

                            }
                        }





                       





                    }
                }



                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 9; j <= 18; j++)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Visible == true)
                        {
                            ws.Cell(i + 14, j - 2).Value = dataGridView1.Rows[i].Cells[j].Value?.ToString();
                        }
                    }

                }
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 9; j <= 18; j++)
                    {
                        if (dataGridView1.Rows[i].Cells[j].Style.BackColor == System.Drawing.Color.Gray)
                        {

                            ws.Cell(i + 14, j - 2).Style.Fill.SetBackgroundColor(XLColor.Gray);

                        }
                    }
                }


            }
            itemsettolerances();


            

            void datagridview2()
            {
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    foreach (DataGridViewColumn col in dataGridView2.Columns)
                    {
                        ws.Cell(row.Index + 14, col.Index + 19).Value = row.Cells[col.Index].Value?.ToString();
                    }
                }
            }
            datagridview2();

            void mergetolerances()
            {
                for (int i = 14; i <= ws.LastRowUsed().RowNumber(); i++)
                {
                    if (ws.Cell(i, 4).Value.ToString() == ws.Cell(i, 5).Value.ToString())
                    {
                        var range = ws.Range($"D{i}:F{i}");
                        range.Value = ws.Cell(i, 5).Value;
                        range.Merge();
                    }
                }
            }
            mergetolerances();



            void mergeitems()
            {

                int lastrow = ws.LastRowUsed().RowNumber();
                for (int i = 15; i <= lastrow; i++)
                {
                    string a = ws.Cell(i, 3).Value.ToString().Trim();
                    string b = ws.Cell(i - 1, 3).Value.ToString().Trim();
                    if (a == b)
                    {
                        var range = ws.Range($"C{i - 1}:C{i}");
                        range.Merge();

                        range = ws.Range($"D{i - 1}:D{i}");
                        range.Merge();

                        range = ws.Range($"E{i - 1}:E{i}");
                        range.Merge();

                        range = ws.Range($"F{i - 1}:F{i}");
                        range.Merge();

                        string k = dataGridView1.Rows[i-15].Cells[$"donnees{comboBox1.Text + comboBox2.Text}"].Value.ToString();
                        int o = int.Parse(k);
                        if (o <= 3)
                        {
                            range = ws.Range($"G{i - 1}:G{i}");
                            range.Merge();

                            range = ws.Range($"H{i - 1}:H{i}");
                            range.Merge();

                            range = ws.Range($"I{i - 1}:I{i}");
                            range.Merge();

                            
                        }
                        range = ws.Range($"J{i - 1}:J{i}");
                        range.Merge();

                        range = ws.Range($"K{i - 1}:K{i}");
                        range.Merge();

                        range = ws.Range($"L{i - 1}:L{i}");
                        range.Merge();

                        range = ws.Range($"M{i - 1}:M{i}");
                        range.Merge();

                        range = ws.Range($"N{i - 1}:N{i}");
                        range.Merge();

                        range = ws.Range($"O{i - 1}:O{i}");
                        range.Merge();

                        range = ws.Range($"P{i - 1}:P{i}");
                        range.Merge();

                        range = ws.Range($"S{i - 1}:S{i}");
                        range.Merge();

                        range = ws.Range($"T{i - 1}:T{i}");
                        range.Merge();

                        range = ws.Range($"U{i - 1}:U{i}");
                        range.Merge();

                        range = ws.Range($"V{i - 1}:V{i}");
                        range.Merge();

                        range = ws.Range($"W{i - 1}:W{i}");
                        range.Merge();

                        range = ws.Range($"X{i - 1}:X{i}");
                        range.Merge();

                        range = ws.Range($"Y{i - 1}:Y{i}");
                        range.Merge();

                        range = ws.Range($"Z{i - 1}:Z{i}");
                        range.Merge();
                    }
                }
            }
            mergeitems();


            void mergelots()
            {

                int deb1 = 14;
                int deb2 = 0;
                int deb3 = 0;


                int fin1 = 0;
                int fin2 = 0;
                int fin3 = ws.LastRowUsed().RowNumber();

                if (dataGridView1.Rows[1].Cells["Tolerances"].Value.ToString() is "3")
                {
                    for (int i = 14; i < ws.LastRowUsed().RowNumber(); i++)
                    {


                        if ((ws.Cell(i, 2).Value.ToString() is "3") && (deb3 == 0))
                        {
                            deb3 = i;
                            fin1 = i - 1;
                        }



                    }
                    var range1 = ws.Range($"B{deb1}:B{fin1}");
                    range1.Style.Alignment.SetTextRotation(90);
                    range1.Value = "Un lot produit même équipe,mêmes conditions(Type,section,machine,couleur)";
                    range1.Merge();

                    var range3 = ws.Range($"B{deb3}:B{fin3}");
                    range3.Style.Alignment.SetTextRotation(90);

                    range3.Value = "Un lot produit même équipe,mêmes conditions(Type,section,machine)";
                    range3.Merge();

                }
                else
                {
                    for (int i = 14; i < ws.LastRowUsed().RowNumber(); i++)
                    {
                        if ((ws.Cell(i, 2).Value.ToString() is "2") && (deb2 == 0))
                        {
                            deb2 = i;
                            fin1 = i - 1;
                        }

                        else if ((ws.Cell(i, 2).Value.ToString() is "3") && (deb3 == 0))
                        {
                            deb3 = i;
                            fin2 = i - 1;
                        }


                    }


                    var range1 = ws.Range($"B{deb1}:B{fin1}");
                    range1.Style.Alignment.SetTextRotation(90);
                    range1.Value = "Un lot produit même équipe,mêmes conditions(Type,section,machine,couleur)";
                    range1.Merge();

                    var range2 = ws.Range($"B{deb2}:B{fin2}");
                    range2.Style.Alignment.SetTextRotation(90);

                    range2.Value = "Un lot produit même équipe,mêmes conditions(Type,section,couleur,toron,machine)";
                    range2.Merge();

                    var range3 = ws.Range($"B{deb3}:B{fin3}");
                    range3.Style.Alignment.SetTextRotation(90);

                    range3.Value = "Un lot produit même équipe,mêmes conditions(Type,section,machine)";
                    range3.Merge();
                }

            }
            mergelots();

            void bordersdatagridview1()
            {
                int lastrow = ws.LastRowUsed().RowNumber();
                int lastcolumn = ws.LastColumnUsed().ColumnNumber();
                for (int i = 14; i <= lastrow; i++)
                {
                    for (int j = 2; j <= 16; j++)
                    {
                        ws.Cell(i, j).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    }
                }
            }

            bordersdatagridview1();

            void bordersdatagridview2()
            {
                int lastrow = ws.LastRowUsed().RowNumber();
                int lastcolumn = ws.LastColumnUsed().ColumnNumber();
                for (int i = 14; i <= lastrow; i++)
                {
                    for (int j = 19; j <= 26; j++)
                    {
                        ws.Cell(i, j).Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                    }
                }
            }

            bordersdatagridview2();

            void good()
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    for (int i = 9; i <= 12; i++)
                    {
                        if ((row.Cells[i].Value?.ToString() is "") && (row.Cells[i].ReadOnly == false))
                        {
                            isEmpty = true;
                        }
                    }
                    bool a = string.IsNullOrWhiteSpace(row.Cells["Jugement Ok"].Value?.ToString());
                    bool b = string.IsNullOrWhiteSpace(row.Cells["Jugement NOK"].Value?.ToString());
                    if (a && b)
                    {
                        isEmpty = true;
                    }
                }
            }
            good();


            void validation()
            {
                var range = ws.Range($"B{ws.LastRowUsed().RowNumber() + 1}:F{ws.LastRowUsed().RowNumber() + 2}");
                range.Merge();
                range.Value = "Validation du Leader QA :";
                range.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);


                range = ws.Range($"G{ws.LastRowUsed().RowNumber()  }:L{ws.LastRowUsed().RowNumber() + 1}");
                range.Merge();
                range.Value = "Validation technicien qualité :";
                range.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);


                range = ws.Range($"M{ws.LastRowUsed().RowNumber() }:P{ws.LastRowUsed().RowNumber() + 1}");
                range.Merge();
                range.Value = "Commentaire :";
                range.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);



            }
            validation();


            void commentaires()
            {
                int first = ws.LastRowUsed().RowNumber() + 2;
                var range = ws.Range($"B{ws.LastRowUsed().RowNumber() + 2}:C{ws.LastRowUsed().RowNumber() + 4}");
                range.Merge();
                range.Value = "Commentaires";
                range.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                range.Style.Fill.SetBackgroundColor(XLColor.LightGray);
                

                range = ws.Range($"D{first}:Z{first}");
                range.Merge();
                range.Value = "Si les valeurs ne correspondent pas aux spécifications du standard , veuillez déclarer sur SIIC le produit non conforme et l'isoler.";
                range.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                range.Style.Fill.SetBackgroundColor(XLColor.LightGray);

                range = ws.Range($"D{first+1}:Z{first+1}");
                range.Merge();
                range.Value = "*Le Leader QA de l'équipe effectue un double contrôle pour Chaque fiche, Choisit la caractériqtique pour effectuer le double controle selon les problèmes rencontrés par l'équipe, note les valeurs trouvées sur la fiche, vérifie le respect du remplissage de toute la fiche et la valide.";
                range.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                range.Style.Fill.SetBackgroundColor(XLColor.LightGray);

                range = ws.Range($"D{first+2}:Z{first+2}");
                range.Merge();
                range.Value = "* Le technicien QA effectue un troisième contrôle ( 1fois / jour / équipe/ Inspecteur QA )";
                range.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                range.Style.Fill.SetBackgroundColor(XLColor.LightGray);


            }
            commentaires();

            if ((isEmpty == true) && (radioButton4.Checked == false))
            {
                MessageBox.Show("Due to one or more cells being empty, the file was not saved.");
            }
            else
            {
                wb.SaveAs($"C:/Users/guizaoui/Downloads/Controle{comboBox1.Text + comboBox2.Text}.xlsx");
            }

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.ReadOnly = false;
            }

            void test()
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    foreach (DataGridViewColumn col in dataGridView1.Columns)
                    {
                        if (row.Cells[col.Index].Style.BackColor == System.Drawing.Color.Gray)
                        {
                            row.Cells[col.Index].ReadOnly = true;
                        }
                    }
                }
            }

            test();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["concernepar"].Value?.ToString() is "toron")
                {
                    row.ReadOnly = true;
                }
            }



        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.ReadOnly = false;
            }

            void test()
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    foreach (DataGridViewColumn col in dataGridView1.Columns)
                    {
                        if (row.Cells[col.Index].Style.BackColor == System.Drawing.Color.Gray)
                        {
                            row.Cells[col.Index].ReadOnly = true;
                        }
                    }
                }
            }

            test();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                CurrencyManager currencyManager1 = (CurrencyManager)BindingContext[dataGridView1.DataSource];

                if (row.Cells["concernepar"].Value?.ToString() is "couleur")
                {
                    currencyManager1.SuspendBinding();
                    row.ReadOnly = true;
                    currencyManager1.ResumeBinding();
                }
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.ReadOnly = false;
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                foreach (DataGridViewColumn col in dataGridView1.Columns)
                {
                    if (row.Cells[col.Index].Style.BackColor == System.Drawing.Color.Gray)
                    {
                        row.Cells[col.Index].ReadOnly = true;
                    }
                }
            }
        }




        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            string[] numerique = new string[] { "stripe numbers", "max thickness" };
            if ((e.ColumnIndex <= 14) && (e.ColumnIndex >= 9))
            {
                try
                {
                    float i;
                    float j;

                    bool a = float.TryParse(Convert.ToString(e.FormattedValue), out i);
                    bool b = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].ReadOnly == false;
                    bool c = (float.TryParse(dataGridView1.Rows[e.RowIndex].Cells["newToleranceMin"].Value.ToString(), out j));

                    if ((!a && b) && c)
                    {
                        e.Cancel = true;
                        MessageBox.Show("Please enter a numeric value");
                    }
                    else if ((a && b) && !c)
                    {
                        if (!numerique.Contains(dataGridView1.Rows[e.RowIndex].Cells["newitems"].Value.ToString()) )
                        {
                            e.Cancel = true;
                            MessageBox.Show("Please enter an alphabetical value");
                        }
                        
                        
                        
                        
                        
                    }
                }
                catch (Exception ex) { Debug.WriteLine(ex); }
            }
            bool m = string.IsNullOrEmpty(Convert.ToString(e.FormattedValue));
            if (((e.ColumnIndex == 15)||(e.ColumnIndex == 16))&&(!m))
            {
                 
                if (e.FormattedValue.ToString().ToUpper() != "X")
                {
                    e.Cancel = true;
                    MessageBox.Show("Please enter a cross(X): ");
                }  
            }
            if (e.ColumnIndex == 16)
            {
                if (!string.IsNullOrEmpty(e.FormattedValue.ToString()))
                {
                    if (!string.IsNullOrEmpty(dataGridView1.Rows[e.RowIndex].Cells[15].Value?.ToString()))
                    {
                        e.Cancel = true;
                        MessageBox.Show("Both values OK and NOK can't be checked at the same time!");
                    }
                }
            }
            if (e.ColumnIndex == 15)
            {
                if (!string.IsNullOrEmpty(e.FormattedValue.ToString()))
                {
                    if (!string.IsNullOrEmpty(dataGridView1.Rows[e.RowIndex].Cells[16].Value?.ToString()))
                    {
                        e.Cancel = true;
                        MessageBox.Show("Both values OK and NOK can't be checked at the same time!");
                    }
                }
            }
            if (e.ColumnIndex == 18)
            {
                bool a = (e.FormattedValue.ToString().ToUpper() == "OK") || (e.FormattedValue.ToString().ToUpper() == "NOK");
                if (a == false)
                {
                    e.Cancel = true;
                    MessageBox.Show("You'll need to either type OK or NOK");
                }
            }
        }

        
    }
}

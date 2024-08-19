using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ExcelDataReader;
using Oracle.ManagedDataAccess.Client;



namespace importarmeta
{
    public partial class Form1 : Form
    {
                        string usuariowinthor;
                        string usuariobanco;
                        string banco;
                        string senhabanco;
                        string numerorotina;
        OracleConnection conwinthor;
        //OracleDataAdapter oda;
        List<meta> tbltemporaria = new List<meta>();


        public Form1(   
            string myusuariowinthor,
            string mysenhabanco,
            string mybanco,
            string myusuariobanco,
            string mynumerorotina
        )
        {
            InitializeComponent();
                usuariowinthor= myusuariowinthor;
                usuariobanco= myusuariobanco;
                banco= mybanco;
                senhabanco= mysenhabanco;
                numerorotina= mynumerorotina;

            try
            {
               // conwinthor = new OracleConnection("Data Source = " + banco +"; User ID = " + usuariobanco +"; Password = " +senhabanco );
             

                string stringConexao = "Data Source=(DESCRIPTION=" +
                        "(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=192.4.1.249)(PORT=1521)))" +
                        "(CONNECT_DATA=(SID=WINT))); " +
                        "User ID="+myusuariobanco+";Password="+ mysenhabanco + ";";
                conwinthor = new OracleConnection(stringConexao);
                conwinthor.Open();

                conwinthor.Close();

                   
            }catch(Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
        string nomearquivo;
        DataSet result;
        int passoatual = 1;
        private OracleParameterCollection ColecaoParametros = new OracleCommand().Parameters;
        string globalano;
        string globalmes;
        string glbalcodfilial;
        string tipometa = "";

        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Text = numerorotina + " - Importar metas via Excel ";
            this.Text =  numerorotina + " - Importar metas via Excel ";
            int anoatual = DateTime.Now.Year;
            ano.Items.Add(anoatual - 1);
            ano.Items.Add(anoatual);
            ano.Items.Add(anoatual + 1);
            // MessageBox.Show("usuario winthor: "+usuariowinthor+ "\nusuario Banco: "+usuariobanco + "\nbanco: "+ banco + "\nsenha banco: "+ senhabanco+ "\nrotina: " + numerorotina);
            label8.Text = usuariowinthor + " (" + banco + "@" + usuariobanco + ")";
            DataSet ods = new DataSet();
            conwinthor.Open();
            OracleCommand cmd1 = new OracleCommand("select codigo from pcfilial where pcfilial.dtexclusao is null and codigo <> 99 order by codigo", conwinthor);
            OracleDataAdapter filias = new OracleDataAdapter(cmd1);
            filias.Fill(ods);
            for( int contagem = 0; contagem <ods.Tables[0].Rows.Count;contagem ++) {
                comboBox3.Items.Add(ods.Tables[0].Rows[contagem][0].ToString());
            }
            conwinthor.Close();

        }

        private void button4_Click(object sender, EventArgs e)
        {
           
          if(  openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                nomearquivo = openFileDialog1.FileName;
                txtArquivo.Text = nomearquivo;
            }
          
        }

        private void button2_Click(object sender, EventArgs e)
        {
            globalano = ano.Text;
            
            switch (comboBox1.Text)
            {
                case "Janeiro":
                    globalmes = "01";
                    break;
                case "Fevereiro":
                    globalmes = "02";
                    break;
                case "Março":
                    globalmes = "03";
                    break;
                case "Abril":
                    globalmes = "04";
                    break;
                case "Maio":
                    globalmes = "05";
                    break;
                case "Junho":
                    globalmes = "06";
                    break;
                case "Julho":
                    globalmes = "07";
                    break;
                case "Agosto":
                    globalmes = "08";
                    break;
                case "Setembro":
                    globalmes = "09";
                    break;
                case "Outubro":
                    globalmes = "10";
                    break;
                case "Novembro":
                    globalmes = "11";
                    break;
                case "Dezembro":
                    globalmes = "12";
                    break;

            }
            glbalcodfilial = comboBox3.Text;
           
            backgroundWorker1.RunWorkerAsync();

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            try {
                if (passoatual == 1) {     
                FileStream stream = File.Open(nomearquivo, FileMode.Open, FileAccess.Read);
               
                //...
                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                

                //3. DataSet - The result of each spreadsheet will be created in the result.Tables
                result = excelReader.AsDataSet(new ExcelDataSetConfiguration()
                {

                    // Gets or sets a value indicating whether to set the DataColumn.DataType 
                    // property in a second pass.
                    UseColumnDataType = true,

                    // Gets or sets a callback to obtain configuration options for a DataTable. 
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                    {

                        // Gets or sets a value indicating the prefix of generated column names.
                        EmptyColumnNamePrefix = "Column",

                        // Gets or sets a value indicating whether to use a row from the 
                        // data as column names.
                        UseHeaderRow = true,

                        // Gets or sets a callback to determine which row is the header row. 
                        // Only called when UseHeaderRow = true.
                       
                    }
                });
                    stream.Close();





                bool codigo =false;
                bool codusur = false;
                bool vlvendaprev = false;
                bool cliposprev = false;
                bool qtdvendaprev = false;
                bool MIX = false;
                    string mensagemdeerro = "Campos obrigatórios não informados no arquivo Excel: \n";

                foreach (object nomecoluna in result.Tables[0].Columns )
                {
                   if (nomecoluna.ToString().ToUpper() == "CODIGO"){codigo = true;  }
                   if (nomecoluna.ToString().ToUpper() == "CODUSUR") { codusur = true;  }
                   if (nomecoluna.ToString().ToUpper() == "VLVENDAPREV") { vlvendaprev = true; }
                   if (nomecoluna.ToString().ToUpper() == "CLIPOSPREV") { cliposprev = true; }
                   if (nomecoluna.ToString().ToUpper() == "QTVENDAPREV") { qtdvendaprev = true; }
                   if (nomecoluna.ToString().ToUpper() == "MIXPREV") { MIX = true; }
                    }
                if (codigo == false){ mensagemdeerro += "CODIGO\n"; }
                if (codusur == false) { mensagemdeerro += "CODUSUR\n"; }
                if (mensagemdeerro != "Campos obrigatórios não informados no arquivo Excel: \n")
                {
                    MessageBox.Show(mensagemdeerro, "Importação Inconsistente",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    return;
                }

                if (radioButton1.Checked == true) {  tipometa = "D";  }
                if (radioButton2.Checked == true) {  tipometa = "S";  }
                if (radioButton3.Checked == true) { tipometa = "M";   }

                tbltemporaria.Clear();


                
             


                for ( int cont = 0; cont< result.Tables[0].Rows.Count; cont++)
                {
                    string descricaogrid="DEPARTAMENTO NAO LOCALIZADO";
                    string nomerca = "VENDEDOR NAO LOCALIZADO";
                         
                   if (executarmanipulacao("SELECT COUNT(*) FROM PCUSUARI WHERE CODUSUR = " + result.Tables[0].Rows[cont]["CODUSUR"].ToString(), "contagem" + cont).ToString() != "0")
                    {
                        //    MessageBox.Show(result.Tables[0].Rows[cont]["CODUSUR"].ToString());
                        nomerca = executarmanipulacao("SELECT NOME FROM PCUSUARI WHERE CODUSUR = " + result.Tables[0].Rows[cont]["CODUSUR"].ToString()   ,"pegarnome"+cont   ).ToString();
                          //  MessageBox.Show(nomerca);
                    } 


                        double VLVENDAPREVGRID = 0;
                        string CLIPOSPREVGRID = "0";
                        string QTVENDAPREVGRID = "0";
                        string MIXGRID = "0";
                    if (tipometa == "D") {
                        if (executarmanipulacao("SELECT count(*) FROM PCDEPTO WHERE codepto =" + result.Tables[0].Rows[cont]["CODIGO"].ToString(),"d").ToString() != "0")
                        {
                           descricaogrid = executarmanipulacao("SELECT DESCRICAO FROM PCDEPTO WHERE codepto =" +result.Tables[0].Rows[cont]["CODIGO"].ToString(),"d").ToString();
                        }
                    }
                    if (tipometa == "S") {

                       if( executarmanipulacao("SELECT count(*) FROM PCSECAO  WHERE codsec = " + result.Tables[0].Rows[cont]["CODIGO"].ToString(),"d").ToString() != "0")
                        { 
                        descricaogrid = executarmanipulacao("SELECT DESCRICAO FROM PCSECAO  WHERE codsec = " + result.Tables[0].Rows[cont]["CODIGO"].ToString(),"d").ToString();
                        }

                    }
                    if (tipometa == "M") {


                            descricaogrid = "TOTAL GERAL";

                    }
                    if (vlvendaprev == true) { VLVENDAPREVGRID= Convert.ToDouble( result.Tables[0].Rows[cont]["VLVENDAPREV"].ToString()); }
                    if (cliposprev == true) { CLIPOSPREVGRID=result.Tables[0].Rows[cont]["CLIPOSPREV"].ToString(); }
                    if (qtdvendaprev == true) { QTVENDAPREVGRID = result.Tables[0].Rows[cont]["QTVENDAPREV"].ToString(); }
                    if (MIX == true) { MIXGRID = result.Tables[0].Rows[cont]["MIXPREV"].ToString(); }

                        // MessageBox.Show(result.Tables[0].Rows[cont]["codigo"].ToString());

                        tbltemporaria.Add(new meta(cont,
                                                
                                                result.Tables[0].Rows[cont]["CODUSUR"].ToString(),
                                                nomerca,
                                                result.Tables[0].Rows[cont]["CODIGO"].ToString(),                                               
                                                descricaogrid,
                                                VLVENDAPREVGRID,
                                                CLIPOSPREVGRID,
                                                QTVENDAPREVGRID,
                                                MIXGRID


                    ));
                }
                    excelReader.Close();
                    passoatual++;
                    Debug.WriteLine("Passo atual: " + tbltemporaria.Count);
                    return;

                }
                if(passoatual == 2)
                {


                    for (int linhaatual=0 ;linhaatual< dataGridView1.Rows.Count;linhaatual++)
                    {
                        string codigo =  dataGridView1.Rows[linhaatual].Cells["CODIGO"].Value.ToString();
                        LimparParametros();
                        AddParametros(":codigo", codigo);
                        AddParametros(":codusur", dataGridView1.Rows[linhaatual].Cells["CODUSUR"].Value.ToString());
                        AddParametros(":data","01/" + globalmes + "/" +globalano + "");
                        AddParametros(":cdfl", glbalcodfilial);
                        AddParametros(":tpmeta","" +tipometa+"");
                       
                       
                       if(
                            executarmanipulacao(
                                "select COUNT(*) "+
                                "from PCMETA "+
                                "where CODIGO =:codigo "+
                                "AND CODUSUR = :codusur "+
                                "AND DATA = TO_DATE(:DATA,'DD/MM/YYYY') "+
                                "AND CODFILIAL = :cdfl "+
                                "AND TIPOMETA = :tpmeta " ,"n"
                             ).ToString() != "0") {
                            //Deletar 
                            executarmanipulacao
                              (
                              "delete " +
                              "from PCMETA " +
                              "where CODIGO =:codigo " +
                              "AND CODUSUR = :codusur " +
                              "AND DATA = TO_DATE(:DATA,'DD/MM/YYYY') " +
                              "AND CODFILIAL = :cdfl " +
                              "AND TIPOMETA = :tpmeta ","n"
                               );
                           
                            
                        }
                        // Inserior novos valores
                        
                        string pcodusur= dataGridView1.Rows[linhaatual].Cells["CODUSUR"].Value.ToString();
                        string pdata =  "01/" + globalmes + "/" + globalano + "";                    
                        string vlvendaprev =dataGridView1.Rows[linhaatual].Cells["VLVENDAPREV"].Value.ToString().Replace(".", "").Replace(",", ".");
                        string qtvendaprev = dataGridView1.Rows[linhaatual].Cells["QTVENDAPREV"].Value.ToString();
                        string cliposprev  = dataGridView1.Rows[linhaatual].Cells["CLIPOSPREV"].Value.ToString();
                        string nomevendedor = dataGridView1.Rows[linhaatual].Cells["NOMEVENDEDOR"].Value.ToString();
                        string nomedepto = dataGridView1.Rows[linhaatual].Cells["DESC"].Value.ToString();
                        string mixprev = dataGridView1.Rows[linhaatual].Cells["MIXPREV"].Value.ToString();

                        if (nomevendedor != "VENDEDOR NAO LOCALIZADO" && nomedepto != "DEPARTAMENTO NAO LOCALIZADO")
                        {
                           
                            executarmanipulacao(
                            "INSERT INTO PCMETA (" +
                            "         CODIGO, CODUSUR, DATA, CODFILIAL, TIPOMETA, VLVENDAPREV, QTVENDAPREV, CLIPOSPREV , rotinalanc, mixprev ) " +
                            "VALUES (" + codigo + "," + pcodusur + ", to_date('" + pdata + "','dd/mm/yyyy'), " + glbalcodfilial + ",'" + tipometa + "'," + vlvendaprev + "," + qtvendaprev + "," + cliposprev + ",3305,"+mixprev+ ") ", "n"
                            );
                        }

                       
                        passoatual = 3;




                    }

                }

            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.ToString());
                return;
            }
            

        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (passoatual == 2) { 
            try { 
            if (result.Tables[0].Rows.Count > 0) {
                dataGridView1.DataSource = tbltemporaria;
                  panel3.Visible = true;
                  button3.Enabled = true;
              }
           else
            {
                MessageBox.Show("Sem informações para exibir");
            }
            }
            catch (Exception EX)
            {
                MessageBox.Show(EX.ToString());
            }
            }
            if (passoatual == 3)
            {
                MessageBox.Show("Processo de importação finalizado!", "Importação", MessageBoxButtons.OK, MessageBoxIcon.Information);
                panel3.Visible = false;
                passoatual=1;
                button3.Enabled = false;
            }
            }

        private object executarmanipulacao(string comando, string parte)
        {
            Debug.WriteLine("Comando: " + comando); // Adiciona esta linha
            OracleCommand cmd = new OracleCommand(comando, conwinthor);
                try {            
                conwinthor.Open();
                cmd.BindByName = true;
                
                foreach (OracleParameter Parametro in ColecaoParametros)
                {
                    cmd.Parameters.Add(new OracleParameter(Parametro.ParameterName, Parametro.Value));
                }
                object retorno = cmd.ExecuteScalar();
                conwinthor.Close();
                return retorno;
                


            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.ToString() +"\n"+ comando+"\n"+parte);
                if(conwinthor.State == ConnectionState.Open)
                {
                    conwinthor.Close();
                }
                return 0;
            }
         

                
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (passoatual == 2) { 
                panel3.Visible = false;
                passoatual--;
                button3.Enabled = false;
            }

        }

        private void dataGridView1_KeyPress(object sender, KeyPressEventArgs e)
        {
          
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Delete)
            {
                try { 
                if( MessageBox.Show("Deseja realmente exlcuir os valores desta linha?", "",MessageBoxButtons.YesNo,MessageBoxIcon.Question)==DialogResult.Yes)
                {
                    int linha = Convert.ToInt16( dataGridView1.CurrentRow.Cells["ID"].Value.ToString());
                    List<META2> meta2 = new List<META2>();
                    dataGridView1.DataSource = meta2;
                                       
                  
                    tbltemporaria.RemoveAt(linha);
                    dataGridView1.DataSource = tbltemporaria;
                }
                }
                catch(Exception EX)
                {
                    MessageBox.Show(EX.ToString());
                }
            }
        }
        public void recaucularMetasAuxiliares()
        {
            calcular_meta_dia();
        }

        public void LimparParametros() { ColecaoParametros.Clear(); }

        public void AddParametros(string parametro, object ValorParametro)
        {
            ColecaoParametros.Add(new OracleParameter(parametro, ValorParametro));
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void calcular_meta_dia()
        {

     
          //  procss_mes();
            DataSet diasUteis = get_dias_uteis();
            int diasuteis_cont = diasUteis.Tables[0].Rows.Count;

            DataSet metas_vendedor_filial = get_metas_totais();
            //Aqui a mágica acontece, ela verificar as metas totais do periodo, e divide pelos dias uteis.
            //Acabou o trabalho de corno papai!!! (Sim, eu convertiaa isso tudo numa planilha tosca que me passavam antes pra gerar o sql)
            for (int linhaatual_rca = 0; linhaatual_rca < metas_vendedor_filial.Tables[0].Rows.Count; linhaatual_rca++)
            {
                string codfilial = metas_vendedor_filial.Tables[0].Rows[linhaatual_rca]["codfilial"].ToString();
                string codusur = metas_vendedor_filial.Tables[0].Rows[linhaatual_rca]["codusur"].ToString();
                double vlvendaprev = Convert.ToDouble(metas_vendedor_filial.Tables[0].Rows[linhaatual_rca]["vlvendaprev"].ToString());
                double meta_dia = vlvendaprev / diasuteis_cont;
                string pdata = globalmes + "/" + globalano + "";
                executarmanipulacao
                             (
                             "delete " +
                             "from pcmetarca " +
                             "where  CODUSUR = " + codusur +
                             " AND  CODFILIAL = " + codfilial +
                             " AND  TO_CHAR(data, 'mm/yyyy') = '" + pdata + @"'", "n"
                              );

          


                for (int linhaatual = 0; linhaatual < diasuteis_cont; linhaatual++)
                {
                    string codigo = diasUteis.Tables[0].Rows[linhaatual]["codfilial"].ToString();
                    //MessageBox.Show(codusur);
                    ColecaoParametros.Clear();

                    DateTime data = DateTime.ParseExact(diasUteis.Tables[0].Rows[linhaatual]["data"].ToString(), "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                    string dataFormatada = data.ToString("dd/MM/yyyy");


                    executarmanipulacao(
                        "INSERT INTO pcmetarca (" +
                        "           CODUSUR    , DATA                                 ,     CODFILIAL    ,     VLVENDAPREV ) " +
                        "VALUES (" + codusur + ", to_date('" + dataFormatada + "','dd/mm/yyyy'), " + codfilial + "," + meta_dia.ToString("0.00", CultureInfo.InvariantCulture) + ") ", "n"
                        );


                }


            }
            MessageBox.Show("Recálculo executado com sucesso", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private DataSet get_metas_totais()
        {
            string pdata = "01/" + globalmes + "/" + globalano + "";
            DataSet dias_consulta = new DataSet();
            string qy = @"
            SELECT 

            codfilial,
            codusur,
            sum(vlvendaprev) vlvendaprev 
            FROM
            pcmeta 
            where
            pcmeta.data  = to_date('" + pdata + @"','dd/mm/yyyy')
            AND TIPOMETA = 'D'
           
            group by
            codfilial,
            codusur";
            dias_consulta = dataset_from_consulta(qy);
            return dias_consulta;
        }
        private DataSet get_dias_uteis()
        {
            string pdata = globalmes + "/" + globalano + "";
            DataSet dias_consulta;
            string qy = @"
            SELECT * 
            FROM pcdiasuteis
            WHERE 
            TO_CHAR(pcdiasuteis.data, 'mm/yyyy') = '" + pdata + @"'
            AND pcdiasuteis.diavendas = 'S'
            AND CODFILIAL = " + glbalcodfilial + @"
            
            ORDER BY DATA"; 
          

            dias_consulta = dataset_from_consulta(qy);
            return dias_consulta;


        }
        private DataSet dataset_from_consulta(string qy)
        {
            Debug.WriteLineIf(true, "Query: " + qy);

            DataSet data = new DataSet();

            try
            {
                conwinthor.Open();

                OracleCommand cmd = new OracleCommand(qy, conwinthor);
                OracleDataAdapter adapter = new OracleDataAdapter(cmd);

                adapter.Fill(data); // Preenche o DataSet com os resultados da consulta
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                conwinthor.Close();
            }

            return data;
        }

        private void procss_mes()
        {
            switch (comboBox1.Text)
            {
                case "Janeiro":
                    globalmes = "01";
                    break;
                case "Fevereiro":
                    globalmes = "02";
                    break;
                case "Março":
                    globalmes = "03";
                    break;
                case "Abril":
                    globalmes = "04";
                    break;
                case "Maio":
                    globalmes = "05";
                    break;
                case "Junho":
                    globalmes = "06";
                    break;
                case "Julho":
                    globalmes = "07";
                    break;
                case "Agosto":
                    globalmes = "08";
                    break;
                case "Setembro":
                    globalmes = "09";
                    break;
                case "Outubro":
                    globalmes = "10";
                    break;
                case "Novembro":
                    globalmes = "11";
                    break;
                case "Dezembro":
                    globalmes = "12";
                    break;

            }

        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            recaucularMetasAuxiliares();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (ano.Text.Length == 0 || comboBox1.Text.Length == 0)
            {
                //MessageBox.Show("Informe o periodo: mês e ano.",,, MessageBoxIcon.Warning);
                MessageBox.Show("Informe o periodo: mês e ano.");

               
                return;

            }
            globalano = ano.Text;
            glbalcodfilial = comboBox3.Text;
            procss_mes();
            backgroundWorker2.RunWorkerAsync();

        }
    }
}

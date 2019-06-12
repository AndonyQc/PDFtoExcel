using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Threading;

namespace PdfExample
{
    public partial class Form1 : Form
    {
        string filtered = "";
        string encabezadoSinComas = "";
        ArrayList listHeaders = new ArrayList();
        ArrayList listPages = new ArrayList();
        bool banderaInicio = false;
        ArrayList itemsToRemove = new ArrayList();
        string ticket = "";
        string date = "";
        DateTime fecha = new DateTime();
        List<string> myList = new List<string>();
        List<string> listaString = new List<string>();
        int contador = 0;
        int contador2 = 0;
        string airport = "";
        int finalAirport = 0;
        int inicioAirport =0;

        //string pagina;

        int contaDelete = 0;

        public Form1()
        {
            InitializeComponent();

        }

        public string encabezado;

        private void RunBtn_Click(object sender, EventArgs e) //Paso uno, resibe PDF => StringBuilder
        {
            if (TxtUrl.Text == "")//En caso que no se ha seleccionado nada
            {
                MessageBox.Show("Please upload a PDF file first !!");
            }
            else
            {
                progressBar1.Visible = true;
                this.timer1.Start();


                PdfReader reader = new PdfReader(TxtUrl.Text.ToString());//Lee la ubicacion del archivo
                StringBuilder builder = new StringBuilder();//El documento se guarada como string builder

                for (int x = 1; x <= reader.NumberOfPages; x++)
                {
                    //PdfDictionary page = reader.GetPageN(x);
                    IRenderListener listener = new SBTextRenderer(builder);
                    PdfContentStreamProcessor processor = new PdfContentStreamProcessor(listener);
                    PdfDictionary pageDic = reader.GetPageN(x);
                    PdfDictionary resourcesDic = pageDic.GetAsDict(PdfName.RESOURCES);
                    processor.ProcessContent(ContentByteUtils.GetContentBytesForPage(reader, x), resourcesDic);

                    if (listPages.Count != 0)
                    {
                        listPages.Add(builder.ToString().Replace(listPages[x - 2].ToString(), ""));
                    }
                    else
                    {
                        listPages.Add(builder.ToString());
                    }


                    if (x == reader.NumberOfPages)
                    {

                        foreach (string pagina in listPages)
                        {
                            for (int p = 0; p < pagina.Length; p++)
                            {
                                if (pagina[p].ToString() == "e" && pagina[p + 1].ToString() == "s" && pagina[p + 2].ToString() == " " && pagina[p + 3].ToString() == "D" && pagina[p + 4].ToString() == "a")
                                {
                                    int index = p + 7;

                                    listHeaders.Add(pagina.Substring(0, index));
                                    break;
                                }
                            }
                        }
                    }


                }//Find de for de conversion de PDF a string

                StringBuilder comenzar = new StringBuilder();
                int Inicio = 0;//Guardara la ubicacion dentro del PDF donde se puede comenzar a copiar caracteres
                int final = 0;

                for (int i = 0; i < builder.ToString().Length; i++)
                {
                    //ATP es la palabra donde comenzara a guardar caracteres
                    char uno = 'A';
                    char dos = 'T';
                    char tres = 'P';

                    encabezado = encabezado + builder.ToString()[i];//Esta variable guarda el encabezado para eliminarlo en las proximas paginas



                    if (builder.ToString()[i] == uno && builder.ToString()[i + 1] == dos && builder.ToString()[i + 2] == tres && banderaInicio == false || builder.ToString()[i] == 'B' && builder.ToString()[i + 1] == 'a' && builder.ToString()[i + 2] == 's' && builder.ToString()[i + 3] == 'e' && builder.ToString()[i + 4] == ':' && banderaInicio == false)
                    {
                        Inicio = i + 2;//Detecta donde inicia a guardar, le suma dos porque son los caracteres de TP de ATP
                                       //  break;//Como se ha encontrado donde comenzar a copiar , se rompe el ciclo
                        banderaInicio = true;
                        //Aqui detecta el nombre del aeropuerto
                        for (int r = i; r < builder.ToString().Length; r++)
                        {
                            string testErase = builder.ToString()[r].ToString();

                            if (builder.ToString()[r]==':')
                            {
                                 inicioAirport = r + 1;
                                
                            }
                            else if(Regex.IsMatch(builder.ToString()[r].ToString(), @"^\d+$"))
                            {
                                 finalAirport = r ;                               
                            }
                             if(inicioAirport !=0 && finalAirport!=0 && airport=="")
                            {
                                airport = builder.ToString().Substring(inicioAirport, (finalAirport - inicioAirport));
                                break;
                            }
                        }

                    }


                    if (i == 2900)
                    {
                        int y = 0;
                    }

                    else if (builder.ToString()[i] == 'U' && builder.ToString()[i + 1] == 'C' && builder.ToString()[i + 2] == 'T' && builder.ToString()[i + 3] == ' ' && builder.ToString()[i + 4] == '2' && builder.ToString()[i + 5] == '2' && builder.ToString()[i + 6] == '1')
                    {
                        final = i;
                        break;
                    }
                    else if (builder.ToString()[i] == 'T' && builder.ToString()[i + 1] == 'T' && builder.ToString()[i + 2] == 'O' && builder.ToString()[i + 3] == ' ' && builder.ToString()[i + 4] == '2' && builder.ToString()[i + 5] == '2' && builder.ToString()[i + 6] == '1')
                    {
                        final = i;
                        break;
                    }




                }//Find de for de deteccion de inicio para copiar

                bool bandera = false;
                // for (int i = Inicio; i < builder.ToString().Length; i++)
                for (int i = Inicio; i < final; i++)
                {
                    if (bandera == false)//Verifica si comienza con cero y con ello lograr comenzar a guardar 
                    {
                        string auxiliar = builder.ToString()[i].ToString();//Se puede borrar, solo se usa de prueba
                        if (builder.ToString()[i].ToString() == "0" || Regex.IsMatch(builder.ToString()[i].ToString(), @"^\d+$"))
                        {
                            bandera = true;
                        }
                    }

                    if (bandera == true)//Como se verifico que si comienza en cero comienza a guardarse
                    {
                        char a = 'T';
                        char b = 'O';
                        char c = 'T';
                        char d = 'A';

                        string auxiliar2 = builder.ToString()[i].ToString();//Solo de prueba, se puede eliminar

                        if (builder.ToString()[i] == a && builder.ToString()[i + 1] == b && builder.ToString()[i + 2] == c && builder.ToString()[i + 3] == d)
                        {
                            int aux = i + 18;//Detecta donde inicia a guardar, le suma dos porque son los caracteres de TP de ATP

                            for (int y = aux; y < builder.ToString().Length; y++)
                            {
                                string auxiliar3 = builder.ToString()[y].ToString();//Se puede borrar, es solo de prueba

                                if (builder.ToString()[y] == 'U' && builder.ToString()[y + 1] == 'S' && builder.ToString()[y + 2] == 'G')
                                {
                                    i = y + 2;
                                    break;
                                }
                            }
                        }
                        else if (builder.ToString()[i] == 'A' && builder.ToString()[i + 1] == 'i' && builder.ToString()[i + 2] == 'r' && builder.ToString()[i + 3] == 'p' && builder.ToString()[i + 4] == 'o' && builder.ToString()[i + 5] == 'r' && builder.ToString()[i + 6] == 't')
                        {
                            break;
                        }
                        else
                        {
                            string cadena = builder.ToString()[i].ToString();
                            comenzar.Append(cadena);
                        }
                    }
                }
                encabezado = encabezado.Replace("DateA", "Date");
                //textBox1.Text = encabezado;
                //textBox2.Text = comenzar.ToString().Replace(encabezado, "");
                string nuevo = comenzar.ToString().Replace(encabezado, "");
                LlenarGrid(nuevo.ToString());
                               
            }
        }

        public void LlenarGrid(string lista)
        {
            listaString = Filter(lista); //Regresa linea filtrada
                                         //StringBuilder remainig = new StringBuilder();


            Boolean banderaFormat = false;

            foreach (string x in listaString)//Aqui en el if se analiza si lleva nueve columnas y de ser asi pasa al else donde se cargan al grid y regresa al if contando desde cero de nuevo
            {
                if (contador < 9 && x != "0")//Cuando alcanza nueve ya lleva una fila de excel y mientras sea distinto a cero agrega a la fila
                {
                    if (contador == 0)
                    {
                        if (x.Trim() != "0" && x.Length >= 14 && Regex.IsMatch(x.Trim(), @"^\d+$"))
                        {
                            banderaFormat = true;

                            string format = FormatDate(x);

                            myList.Add(format.Substring(0, 6));
                            myList.Add(format.Substring(7, 10));
                        }
                        else
                        {
                            continue;
                        }
                    }
                    else
                    {
                        myList.Add(x);
                    }
                    contador++;
                    //remainig.Append(x);
                }
                //else if (x != "0" || contador2 == listaString.Count - 1)
                else if (contador == 9 || contador2 == listaString.Count - 1)
                {
                    dataGridView1.Rows.Add(myList[0], myList[1], myList[2] + myList[3], myList[4], myList[6], myList[7], myList[8].Trim(), myList[9].Trim(), airport);
                    contador = 0;
                    myList.Clear();

                    if (x != "0" && x.Length >= 14)
                    {
                        string format = FormatDate(x);

                        myList.Add(format.Substring(0, 6));
                        myList.Add(format.Substring(7, 10));

                        contador++;
                    }
                 
                }
                contador2++;
            }

            dataGridView1.Update(); //Actualiza el grid
            lblRows.Text = dataGridView1.RowCount.ToString() + " rows";
        }

        public string FormatDate(string ticketDate)
        {
            string aux = ticketDate.Trim();
            if (aux.Length > 14)
            {
                aux = aux.Substring(aux.Length - 14);
                // aux = aux.Substring(1, aux.Length - 1);//Se le quita un cero
            }
            if (aux.Length < 14)
            {
                return "";
            }

            ticket = aux.Substring(0, 6).ToString();
            date = (aux.Substring(6, 2).ToString() + "-" + aux.Substring(8, 2).ToString() + "-" + aux.Substring(10, 4).ToString()).Trim();

            fecha = DateTime.ParseExact(date, "dd-MM-yyyy",
                                                      CultureInfo.InvariantCulture);

            string a = ticket + " " + fecha.ToString("MM/dd/yyyy");
            return a;
        }

        public string FilterString(string Hojas) // Se usa solo para el encabezado
        {
            // string Hojas = ReplaceWords(PDF);
            StringBuilder listaString = new StringBuilder();//Insertar cada uno de numeros apenas encunetre un espacio
            string a = " ";

            for (int x = 0; x < Hojas.Length; x++)
            {

                if (Hojas[x].ToString().Trim() != "")
                {
                    a = a + Hojas[x].ToString();
                }
                else if (a != "")
                {
                    listaString.Append(a);
                    a = "";
                }
            }
            return listaString.ToString();
        }

        public List<string> Filter(string PDF)
        {
            string Hojas = ReplaceWords(PDF);
            List<string> listaString = new List<string>();//Insertar cada uno de numeros apenas encunetre un espacio
            string a = " ";

            for (int x = 0; x < Hojas.Length; x++)
            {
                //Hay unos espacios vacios que no se identifican, este prime if detecta si es 'algo' entonces lo concatena
                //el else es cuando descubre que viene ese espacio raro y verifica si la 'a' no esta vacia(tiene algo concatenado)
                //y en caso de tener algo lo guarda en un campo en la lista como una palabra digamoslo asi
                if (Hojas[x].ToString().Trim() != "")
                {
                    a = a + Hojas[x].ToString();
                }
                else if (a != "")
                {
                    listaString.Add(a);
                    a = "";
                }
            }

            //Este otro for recorrera la lista y remplazara los fees eliminando su espacio mas tres espacios mas

            foreach (string item in listaString)
            {
                if (item == "Volume")
                {
                    itemsToRemove.Add(item);
                    itemsToRemove.Add(listaString.ElementAt(contaDelete + 1));
                    itemsToRemove.Add(listaString.ElementAt(contaDelete + 2));
                    itemsToRemove.Add(listaString.ElementAt(contaDelete + 3));
                }
                contaDelete++;
            }
            foreach (var item in itemsToRemove)
            {
                listaString.Remove(item.ToString());
            }
            return listaString;
        }

        public string ReplaceWords(string Hojas)
        {
            filtered = Hojas;


            foreach (string head in listHeaders)
            {
                filtered = filtered.Replace(head, "");
            }

            encabezadoSinComas = encabezado.Replace(",", ".");
            filtered = filtered.Replace(",", ".");
            filtered = filtered.Replace("USG", "");
            filtered = filtered.Replace(" 221", "");
            filtered = filtered.Replace("Jet", "");
            filtered = filtered.Replace("A1", "");
            filtered = filtered.Replace("0MOD", "");
            filtered = filtered.Replace("ISOIL/APPagina", "");
            filtered = filtered.Replace("1 di2", "");
            filtered = filtered.Replace("Base:", "ATP:");
            filtered = filtered.Replace("LJ 1", "");

            //  filtered = filtered.Replace(" 22", "");
            filtered = filtered.Replace(encabezadoSinComas, "");



            return filtered;
        }

        #region Button that charge URL needed for pdf
        private void AddFile_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Update();
            listPages.Clear();
            itemsToRemove.Clear();
            contaDelete = 0;
            ticket = "";
            date = "";
            fecha = new DateTime();
            myList.Clear();
            listaString.Clear();
            int contador = 0;
            int contador2 = 0;
            listHeaders.Clear();
            this.progressBar1.Value = 0;
            lblRows.Text = "0 rows";
            airport = "";
            inicioAirport = 0;
            finalAirport = 0;
            banderaInicio = false;

            OpenFileDialog dlg = new OpenFileDialog();
            string filePath;
            dlg.Filter = "Pdf Files|*.pdf";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                filePath = dlg.FileName.ToString();
                TxtUrl.Text = filePath.ToString();
            }
        }
        #endregion

        #region Boton que convierte a excell
        private void ExcelBtn_Click(object sender, EventArgs e)
        {
            if (TxtUrl.Text == "")//En caso que no se ha seleccionado nada
            {
                MessageBox.Show("Please upload a PDF file first !!");
            }
            else if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Please run the app to get rows to convert to excel file !!");
            }
            else
            {
                copyAlltoClipboard();
                Microsoft.Office.Interop.Excel.Application xlexcel;
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlexcel = new Microsoft.Office.Interop.Excel.Application();
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlexcel.Cells[1, 2].EntireColumn.NumberFormat = "000000";
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            }
        }

        private void copyAlltoClipboard()
        {
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        #endregion

        private void eNISPAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            lblSupplier.Text = "ENI SPA";
        }

        private void sOCARToolStripMenuItem_Click(object sender, EventArgs e)
        {
            lblSupplier.Text = "SOCAR";
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
           //this.progressBar1.Increment(1);

            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            progressBar1.Step = 15;
            progressBar1.PerformStep();
            if (progressBar1.Value == 100)
            {
                //int milliseconds = 2000;
                //Thread.Sleep(milliseconds);
                progressBar1.Value = 0;
                timer1.Enabled = false;
                progressBar1.Visible = false;
            }

        }

        
    }

    public class SBTextRenderer : IRenderListener
    {

        private StringBuilder _builder;
        public SBTextRenderer(StringBuilder builder)
        {
            _builder = builder;
        }
        #region IRenderListener Members

        public void BeginTextBlock()
        {
        }

        public void EndTextBlock()
        {
        }

        public void RenderImage(ImageRenderInfo renderInfo)
        {
        }

        public void RenderText(TextRenderInfo renderInfo)
        {
            _builder.Append(renderInfo.GetText());
        }

        #endregion
    }

    public class ElementosPdf
    {

        private string a;
        private string b;
        private string c;
        private string d;
        private string e;
        private string f;

        public string A { get => a; set => a = value; }
        public string B { get => b; set => b = value; }
        public string C { get => c; set => c = value; }
        public string D { get => d; set => d = value; }
        public string E { get => e; set => e = value; }
        public string F { get => f; set => f = value; }
    }


}

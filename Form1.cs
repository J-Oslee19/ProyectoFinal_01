using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
//using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;



namespace ProyectoFinal_01
{
    public partial class Form1 : Form
    {
        TextBox txtPrompt;
        TextBox txtRespuesta;
        Label lblEstado;
        string RutaDestino = "C:\\Users\\CompuFire\\Documents\\Documentos generador por IA";
        string RutaPlantilla = @"C:\Plantillas\PlantillaConsulta.docx";
        // Campos de clase para los controles
        private TextBox txtTema;
        private TextBox txtResultado;

        public Form1()
        {
            InitializeComponent();
            CrearInterfaz();


        }

        private void btnGuardarPlantilla_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Aquí se ejecutará el guardado en plantilla Word.");
        }


        private void CrearInterfaz()
        {
            this.Text = "Investigación con IA";
            this.Width = 700;
            this.Height = 500;

            Label lblPrompt = new Label
            {
                Text = "Escribe tu consulta:",
                Top = 20,
                Left = 20,
                Width = 200
            };
            this.Controls.Add(lblPrompt);

            txtPrompt = new TextBox
            {
                Top = 45,
                Left = 20,
                Width = 630,
                Height = 100,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };
            this.Controls.Add(txtPrompt);
            Button btnPlantilla = new Button
            {
                Text = "Guardar Plantilla Word",
                Location = new System.Drawing.Point(250, 380),
                Width = 180
            };
            btnPlantilla.Click += btnGenerarPlantilla_Click; 
            this.Controls.Add(btnPlantilla);

            Button btnConsultar = new Button
            {
                Text = "Consultar IA",
                Top = 155,
                Left = 20,
                Width = 150
            };

            btnConsultar.Click += async (s, e) =>
            {
                string prompt = txtPrompt.Text;

                if (!string.IsNullOrWhiteSpace(prompt))
                {
                    await GuardarDBAsync(); 
                }
                else
                {
                    MessageBox.Show("Por favor, escribe un tema o pregunta para consultar a la IA.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };


            this.Controls.Add(btnConsultar);
            Label lblRespuesta = new Label
            {
                Text = "Respuesta de la IA:",
                Top = 190,
                Left = 20,
                Width = 200
            };
            this.Controls.Add(lblRespuesta);

            txtRespuesta = new TextBox
            {
                Top = 215,
                Left = 20,
                Width = 630,
                Height = 150,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true
            };
            this.Controls.Add(txtRespuesta);
            Button btnGenerarWord = new Button
            {
                Text = "Guardar en Word",
                Top = 380, // Ajusta la posición según tu diseño
                Left = 50,
                Width = 150
            };
            btnGenerarWord.Click += BtnGenerarWord_Click;
            this.Controls.Add(btnGenerarWord);

            Button btnGenerarPTT = new Button
            {
                Text = "Guardar en Presentacion PowerPoint",
                Top = 380, // Ajusta la posición según tu diseño
                Left = 480,
                Width = 190
            };
            btnGenerarPTT.Click += btnGenerarPresentacion_Click;
            this.Controls.Add(btnGenerarPTT);


      

            lblEstado = new Label
            {
                Text = "",
                Top = 415,
                Left = 20,
                Width = 630,
                ForeColor = System.Drawing.Color.DarkGreen
            };
            this.Controls.Add(lblEstado);
        }

        public async Task<string> ConsultarGroqAsync(string prompt)
        {
            try
            {
                string apiKey = "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"; // Clave API
                string endpoint = "https://api.groq.com/openai/v1/chat/completions";

                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization =
                        new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);

                    var requestBody = new
                    {
                        model = "llama3-70b-8192", 
                        messages = new[]
                        {
                    new Dictionary<string, string>
                    {
                        { "role", "user" },
                        { "content", prompt }
                    }
                },
                        temperature = 0.7
                    };

                    var json = JsonConvert.SerializeObject(requestBody);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync(endpoint, content);
                   
                    if ((int)response.StatusCode == 429)
                        return "⚠️ Demasiadas solicitudes. Intenta de nuevo más tarde.";

                    if ((int)response.StatusCode == 401)
                        return "❌ API Key inválida. Verifica tu clave.";

                    if ((int)response.StatusCode == 402)
                        return "❌ Error 402: No tienes créditos o acceso al modelo.";

                    if ((int)response.StatusCode == 400)
                    {
                        string errorContent = await response.Content.ReadAsStringAsync();
                        return $"❌ Error 400: Solicitud inválida. Detalles: {errorContent}";
                    }

                    response.EnsureSuccessStatusCode();

                    string jsonResult = await response.Content.ReadAsStringAsync();
                    dynamic result = JsonConvert.DeserializeObject(jsonResult);

                    
                    return result.choices[0].message.content.ToString().Trim();
                    
                    
                }
            }
            catch (HttpRequestException ex)
            {
                return $"❌ Error HTTP: {ex.Message}";
            }
            catch (Exception ex)
            {
                return $"❌ Error general: {ex.Message}";
            }
        }


        private void GenerarDocumentoDesdePlantilla(string titulo, string contenido)
        {
            try
            {
                // Ruta de la plantilla
                string rutaPlantilla = @"C:\\Users\\CompuFire\\source\\repos\\ProyectoFinal_01\\Docs\\PlantillaConsulta.docx";

                // Ruta de salida
                string rutaSalida = $"C:\\Users\\CompuFire\\source\\repos\\ProyectoFinal_01\\Docs\\Consulta_{DateTime.Now:yyyyMMdd_HHmmss}.docx";

                // Asegura que el directorio exista
                Directory.CreateDirectory(Path.GetDirectoryName(rutaSalida));

                // Inicia Word
                Word.Application wordApp = new Word.Application();
                wordApp.Visible = false; // No mostrar Word

                // Abre la plantilla
                Word.Document doc = wordApp.Documents.Open(rutaPlantilla);

                // Reemplaza los marcadores
                if (doc.Bookmarks.Exists("TituloConsulta"))
                {
                    doc.Bookmarks["TituloConsulta"].Range.Text = titulo;
                }

                if (doc.Bookmarks.Exists("FechaConsulta"))
                {
                    doc.Bookmarks["FechaConsulta"].Range.Text = DateTime.Now.ToString("dd/MM/yyyy");
                }

                if (doc.Bookmarks.Exists("ContenidoConsulta"))
                {
                    doc.Bookmarks["ContenidoConsulta"].Range.Text = contenido;
                }

                // Guarda el documento como nuevo
                doc.SaveAs2(rutaSalida);
                doc.Close();
                wordApp.Quit();

                MessageBox.Show("✅ Documento generado correctamente:\n" + rutaSalida, "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("❌ Error al generar el documento:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnGenerarPlantilla_Click(object sender, EventArgs e)
        {
            

            string titulo = txtPrompt.Text;
            string contenido = txtRespuesta.Text;

            GenerarDocumentoDesdePlantilla(titulo, contenido);
        }

        public void GuardarDocumentoWord(string titulo, string contenido, string piePagina)
        {
            try
            {
                // Crear instancia de Word
                Word.Application wordApp = new Word.Application();
                Word.Document doc = wordApp.Documents.Add();

                // Ocultar Word durante la creación (más rápido y sin interferencia visual)
                wordApp.Visible = false;

                // Agregar título
                if (!string.IsNullOrEmpty(titulo))
                {
                    Word.Paragraph titleParagraph = doc.Content.Paragraphs.Add();
                    titleParagraph.Range.Text = titulo;
                    titleParagraph.Range.Font.Bold = 1;
                    titleParagraph.Range.Font.Size = 16;
                    titleParagraph.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    titleParagraph.Range.InsertParagraphAfter();
                }

                // Agregar contenido
                if (!string.IsNullOrEmpty(contenido))
                {
                    Word.Paragraph contentParagraph = doc.Content.Paragraphs.Add();
                    contentParagraph.Range.Text = contenido;
                    contentParagraph.Range.Font.Size = 12;
                    contentParagraph.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    contentParagraph.Range.InsertParagraphAfter();
                }

                // Agregar pie de página
                if (!string.IsNullOrEmpty(piePagina))
                {
                    // Salto de página antes del pie
                    doc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);

                    Word.Paragraph footerParagraph = doc.Content.Paragraphs.Add();
                    footerParagraph.Range.Text = piePagina;
                    footerParagraph.Range.Font.Italic = 1;
                    footerParagraph.Range.Font.Size = 10;
                    footerParagraph.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    footerParagraph.Range.InsertParagraphAfter();
                }
               
                string nombreArchivo = $"trabajo_{DateTime.Now:yyyyMMdd_HHmmss}.docx";

       

                // Crear la carpeta si no existe
                if (!Directory.Exists(RutaDestino))
                {
                    Directory.CreateDirectory(RutaDestino);
                }

                string filePath = Path.Combine(RutaDestino, nombreArchivo);

                doc.SaveAs2(filePath);

                // Cerrar y liberar recursos
                doc.Close();
                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);

                MessageBox.Show("Documento Word generado con éxito", "Proceso completado", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al generar el documento: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void BtnGenerarWord_Click(object sender, EventArgs e)
        {
            string titulo = txtPrompt.Text; // Cambiado de txtTitulo a txtPrompt
            string contenido = txtRespuesta.Text; // Cambiado de txtContenido a txtRespuesta
            string piePagina = "Generado automáticamente con IA"; // Pie de página predeterminado

            GuardarDocumentoWord(titulo, contenido, piePagina);
        }

        public void GenerarPresentacionPowerPoint(string titulo, string contenido, string piePagina)
        {
            try
            {
                var pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
                var presentations = pptApp.Presentations;
                var presentation = presentations.Add(MsoTriState.msoFalse); // No mostrar PowerPoint
                var slides = presentation.Slides;

                // Diapositiva 1: Título
                var slide1 = slides.Add(1, PpSlideLayout.ppLayoutTitle);
                slide1.Shapes[1].TextFrame.TextRange.Text = titulo;
                slide1.Shapes[1].TextFrame.TextRange.Font.Name = "Calibri";
                slide1.Shapes[1].TextFrame.TextRange.Font.Size = 36;
                slide1.Shapes[2].TextFrame.TextRange.Text = "Presentación generada automáticamente";
                slide1.Shapes[2].TextFrame.TextRange.Font.Name = "Calibri";
                slide1.Shapes[2].TextFrame.TextRange.Font.Size = 20;
                slide1.FollowMasterBackground = MsoTriState.msoFalse;
                slide1.Background.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightSteelBlue);

                // Dividir contenido en partes
                List<string> partesContenido = DividirContenidoEnPartes(contenido, 800);
                int slideIndex = 2;

                foreach (var parte in partesContenido)
                {
                    var slide = slides.Add(slideIndex++, PpSlideLayout.ppLayoutText);
                    slide.Shapes[1].TextFrame.TextRange.Text = "Contenido";
                    slide.Shapes[1].TextFrame.TextRange.Font.Name = "Calibri";
                    slide.Shapes[1].TextFrame.TextRange.Font.Size = 28;

                    slide.Shapes[2].TextFrame.TextRange.Text = parte;
                    slide.Shapes[2].TextFrame.TextRange.Font.Name = "Calibri";
                    slide.Shapes[2].TextFrame.TextRange.Font.Size = 20;

                    slide.FollowMasterBackground = MsoTriState.msoFalse;
                    slide.Background.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.WhiteSmoke);
                }

                // Diapositiva final de cierre
                var slideFinal = slides.Add(slideIndex, PpSlideLayout.ppLayoutTitleOnly);
                slideFinal.Shapes[1].TextFrame.TextRange.Text = "¡Gracias por su atención!";
                slideFinal.Shapes[1].TextFrame.TextRange.Font.Name = "Calibri";
                slideFinal.Shapes[1].TextFrame.TextRange.Font.Size = 32;
                slideFinal.Shapes[1].TextFrame.TextRange.Font.Bold = MsoTriState.msoTrue;
                slideFinal.FollowMasterBackground = MsoTriState.msoFalse;
                slideFinal.Background.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.LightBlue);

                // Agregar pie de página
                foreach (Slide slide in slides)
                {
                    slide.HeadersFooters.Footer.Visible = MsoTriState.msoTrue;
                    slide.HeadersFooters.Footer.Text = piePagina;
                }

                // Guardar la presentación
                string nombreArchivo = $"Presentacion_{DateTime.Now:yyyyMMdd_HHmmss}.pptx";
                string rutaCompleta = Path.Combine(RutaDestino, nombreArchivo);
                presentation.SaveAs(rutaCompleta);

                presentation.Close();
                pptApp.Quit();

                MessageBox.Show($"✅ Presentación guardada exitosamente en:\n{rutaCompleta}", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Error al generar presentación: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<string> DividirContenidoEnPartes(string texto, int caracteresMaximos)
        {
            List<string> partes = new List<string>();
            int inicio = 0;

            while (inicio < texto.Length)
            {
                int longitudRestante = texto.Length - inicio;
                int longitud = Math.Min(caracteresMaximos, longitudRestante);

                // Corregimos el índice para LastIndexOf
                int indiceBusqueda = inicio + longitud - 1;
                if (indiceBusqueda >= texto.Length)
                {
                    indiceBusqueda = texto.Length - 1;
                }

                int fin = texto.LastIndexOf(' ', indiceBusqueda);
                if (fin == -1 || fin <= inicio)
                {
                    fin = inicio + longitud;
                }

                if (fin > texto.Length)
                {
                    fin = texto.Length;
                }

                string parte = texto.Substring(inicio, fin - inicio).Trim();
                partes.Add(parte);
                inicio = fin;
            }

            return partes;
        }


        private void btnGenerarPresentacion_Click(object sender, EventArgs e)
        {
            string titulo = txtPrompt.Text;
            string contenido = txtRespuesta.Text;
            string piePagina = "Generado automáticamente con IA";

            // Llamar a la función que genera la presentación
            GenerarPresentacionPowerPoint(titulo, contenido, piePagina);
        }

        public void GuardarEnBaseDeDatos(string prompt, string respuesta)
        {
            try
            {
                string connectionString = "Server=DESKTOP-AN2IAAL\\SQLEXPRESS;Database=promptdb;Trusted_Connection=True;";
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string query = "INSERT INTO PromptsIA (Prompt, Respuesta) VALUES (@prompt, @respuesta)";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@prompt", prompt);
                        command.Parameters.AddWithValue("@respuesta", respuesta);
                        command.ExecuteNonQuery();
                    }
                }

                MessageBox.Show("✅ Datos guardados en la base de datos correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Error al guardar en la base de datos: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public async Task GuardarDBAsync()
        {
            string prompts = txtPrompt.Text;
            string respuesta = txtRespuesta.Text;

           

            string consultaRespuesta = await ConsultarGroqAsync(prompts);
            txtRespuesta.Text = consultaRespuesta; // Actualiza el campo de texto con la respuesta obtenida
           
            GuardarEnBaseDeDatos(prompts, consultaRespuesta);
        }
        private async void btnConsultar_Click(object sender, EventArgs e)
        {
            string prompt = txtPrompt.Text;

            if (!string.IsNullOrWhiteSpace(prompt))
            {
                await GuardarDBAsync(); // Corrected: Removed assignment to 'string respuesta'
            }
            else
            {
                MessageBox.Show("Por favor, escribe un tema o pregunta para consultar a la IA.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
          
        }
    }
}
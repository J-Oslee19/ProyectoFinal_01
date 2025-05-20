using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;


namespace ProyectoFinal_01
{
    public partial class Form1 : Form
    {
        TextBox txtPrompt;
        TextBox txtRespuesta;
        Label lblEstado;
        string RutaDestino = "C:\\Users\\CompuFire\\Documents\\Documentos generador por IA";

        public Form1()
        {
            InitializeComponent();
            CrearInterfaz();
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
                    await GuardarDBAsync(); // Eliminado el intento de asignar el resultado a una variable
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
                Left = 260,
                Width = 150
            };
            btnGenerarWord.Click += BtnGenerarWord_Click;
            this.Controls.Add(btnGenerarWord);

            Button btnGenerarPTT = new Button
            {
                Text = "Guardar en Presentacion PowerPoint",
                Top = 380, // Ajusta la posición según tu diseño
                Left = 450,
                Width = 200
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
                string apiKey = "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"; // Reemplázala con tu clave válida
                string endpoint = "https://api.groq.com/openai/v1/chat/completions";

                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization =
                        new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);

                    var requestBody = new
                    {
                        model = "llama3-70b-8192", // NUEVO modelo soportado
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
                // Generar nombre único con fecha y hora
                //string nombreArchivo = $"trabajo_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
                //// Guardar el documento en Mis Documentos como "trabajo.docx"
                //string filePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), nombreArchivo);

                string nombreArchivo = $"trabajo_{DateTime.Now:yyyyMMdd_HHmmss}.docx";

                // Ruta completa: Mis Documentos\InvestigacionesIA\mi_investigacion.docx
               // string rutaCarpeta = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), nombreCarpeta);

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
                // Cambiar la referencia explícita para evitar ambigüedad
                Microsoft.Office.Interop.PowerPoint.Application pptApp = new Microsoft.Office.Interop.PowerPoint.Application();
                // Crear instancia de PowerPoint
                //Application pptApp = new Application();
                Presentations presentations = pptApp.Presentations;
                Presentation presentation = presentations.Add(MsoTriState.msoTrue);
                
                Slides slides = presentation.Slides;
                Slide slide1 = slides.Add(1, PpSlideLayout.ppLayoutTitle);
                slide1.Shapes[1].TextFrame.TextRange.Text = titulo;
                slide1.Shapes[2].TextFrame.TextRange.Text = "Investigación generada por IA";

                // Diapositiva 2: Contenido
                Slide slide2 = slides.Add(2, PpSlideLayout.ppLayoutText);
                slide2.Shapes[1].TextFrame.TextRange.Text = "Contenido Principal";
                slide2.Shapes[2].TextFrame.TextRange.Text = contenido;

              
                
                string nombreArchivo = $"Presentacion_{DateTime.Now:yyyyMMdd_HHmmss}.pptx";
                string rutaCompleta = Path.Combine(RutaDestino, nombreArchivo);

                presentation.SaveAs(rutaCompleta);
               // presentation.Close();
                pptApp.Quit();

                MessageBox.Show($"✅ Presentación guardada exitosamente en:\n{rutaCompleta}", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"❌ Error al generar presentación: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
        
    }
}
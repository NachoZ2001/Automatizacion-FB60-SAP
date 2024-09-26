using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace FB60_SAP
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // Establecer el estilo del borde y deshabilitar el cambio de tamaño
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            // Establecer el tamaño mínimo y máximo para evitar el cambio de tamaño
            this.MinimumSize = this.MaximumSize = this.Size;

            // Obtener la ruta del directorio donde está el ejecutable
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

            // Construir la ruta relativa al script dentro del repositorio
            string filepathDiccionario = Path.Combine(baseDirectory, @"..\..\..\..\Script\Data\ACREEDORES.txt");

            // Calcular el diccionario y almacenarlo en la clase estática
            DiccionarioCodigosAcreedores.CalcularDiccionario(filepathDiccionario);
        }

        private void SeleccionarArchivo(TextBox textBox)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Archivos Excel|*.xlsx;*.csv";
                openFileDialog.Title = "Seleccionar el archivo Excel o CSV";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    textBox.Text = openFileDialog.FileName;
                }
            }
        }

        private void buttonSeleccionarExcel_Click(object sender, EventArgs e)
        {
            SeleccionarArchivo(textBoxRutaExcel);
        }

        private void buttonEjecutarScript_Click(object sender, EventArgs e)
        {
            // Obtener la ruta del directorio donde está el ejecutable
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

            MessageBox.Show(baseDirectory);

            // Construir la ruta relativa al script dentro del repositorio
            string scriptRelativePath = Path.Combine(baseDirectory, @"..\..\..\..\Script\SCRIPT_SAP.vbs");

            // Ejecutar el script
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = scriptRelativePath,
                UseShellExecute = true // Esto es importante para permitir la ejecución del script
            };

            try
            {
                Process.Start(startInfo);
                MessageBox.Show("Script ejecutado con éxito.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al ejecutar el script: " + ex.Message);
            }
        }

        private void buttonTransformarExcel_Click(object sender, EventArgs e)
        {
            
            using (var workbook = new XLWorkbook(textBoxRutaExcel.Text))
            {
                var worksheet = workbook.Worksheet(1);

                // Indice de las columnas del Excel
                int indiceColumnaFecha = ObtenerIndiceColumna(worksheet, "Fecha");
                int indiceColumnaTipo = ObtenerIndiceColumna(worksheet, "Tipo");
                int indiceColumnaPuntoVenta = ObtenerIndiceColumna(worksheet, "Punto de Venta");
                int indiceColumnaComprobante = ObtenerIndiceColumna(worksheet, "Número Desde");
                int indiceColumnaComprobanteVentaHasta = ObtenerIndiceColumna(worksheet, "Número Hasta");
                int indiceColumnaCodigoAutorizacion = ObtenerIndiceColumna(worksheet, "Cód. Autorización");
                int indiceColumnaTipoDocEmisor = ObtenerIndiceColumna(worksheet, "Tipo Doc. Emisor");
                int indiceColumnaNroDocEmisor = ObtenerIndiceColumna(worksheet, "Nro. Doc. Emisor");
                int indiceColumnaDenominacionEmisor = ObtenerIndiceColumna(worksheet, "Denominación Emisor");
                int indiceColumnaTipoCambio = ObtenerIndiceColumna(worksheet, "Tipo Cambio");
                int indiceColumnaMoneda = ObtenerIndiceColumna(worksheet, "Moneda");
                int indiceColumnaNetoGravado = ObtenerIndiceColumna(worksheet, "Imp. Neto Gravado");
                int indiceColumnaNetoNoGravado = ObtenerIndiceColumna(worksheet, "Imp. Neto No Gravado");
                int indiceColumnaOperacionesExentas = ObtenerIndiceColumna(worksheet, "Imp. Op. Exentas");
                int indiceColumnaOtrosTributos = ObtenerIndiceColumna(worksheet, "Otros Tributos");
                int indiceColumnaIVA = ObtenerIndiceColumna(worksheet, "IVA");
                int indiceColumnaImporteTotal = ObtenerIndiceColumna(worksheet, "Imp. Total");
                int indiceColumnaTexto = ObtenerIndiceColumna(worksheet, "texto");
                int indiceColumnaIndicador = ObtenerIndiceColumna(worksheet, "indicador");

                int ultimaFila = worksheet.LastRowUsed().RowNumber();

                // Crear un nuevo workbook para el archivo de salida
                var workbookSalida = new XLWorkbook();
                var worksheetSalida = workbookSalida.Worksheets.Add("Datos Procesados");

                // Definir los encabezados de la nueva hoja de cálculo
                worksheetSalida.Cell(1, 1).Value = "Fecha";
                worksheetSalida.Cell(1, 2).Value = "Tipo";
                worksheetSalida.Cell(1, 3).Value = "Punto de Venta";
                worksheetSalida.Cell(1, 4).Value = "Número Desde";
                worksheetSalida.Cell(1, 5).Value = "Número Hasta";
                worksheetSalida.Cell(1, 6).Value = "Cód. Autorización";
                worksheetSalida.Cell(1, 7).Value = "Tipo Doc. Emisor";
                worksheetSalida.Cell(1, 8).Value = "Nro. Doc. Emisor";
                worksheetSalida.Cell(1, 9).Value = "Denominación Emisor";
                worksheetSalida.Cell(1, 10).Value = "Tipo Cambio";
                worksheetSalida.Cell(1, 11).Value = "Moneda";
                worksheetSalida.Cell(1, 12).Value = "Imp. Neto Gravado";
                worksheetSalida.Cell(1, 13).Value = "Imp. Neto No Gravado";
                worksheetSalida.Cell(1, 14).Value = "Imp. Op. Exentas";
                worksheetSalida.Cell(1, 15).Value = "Otros Tributos";
                worksheetSalida.Cell(1, 16).Value = "IVA";
                worksheetSalida.Cell(1, 17).Value = "Imp. Total";
                worksheetSalida.Cell(1, 18).Value = "texto";
                worksheetSalida.Cell(1, 19).Value = "indicador";
                worksheetSalida.Cell(1, 20).Value = "fecha convertida";
                worksheetSalida.Cell(1, 21).Value = "cuenta mayor";
                worksheetSalida.Cell(1, 22).Value = "codigo acreedor";
                worksheetSalida.Cell(1, 23).Value = "referencia";
                worksheetSalida.Cell(1, 24).Value = "centro costo";
                worksheetSalida.Cell(1, 25).Value = "fecha contabilidad";
                worksheetSalida.Cell(1, 26).Value = "neto  + IVA";
                worksheetSalida.Cell(1, 27).Value = "tipo venta";
                worksheetSalida.Cell(1, 28).Value = "categoria venta";
                worksheetSalida.Cell(1, 29).Value = "tipo factura";
                worksheetSalida.Cell(1, 30).Value = "Detalle";


                int filaSalida = 2; // Fila inicial en el nuevo Excel (la 1 es para encabezados)

                // Obtener el directorio base donde está el ejecutable
                string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

                // Construir la ruta de la carpeta donde se guardará el archivo
                string rutaGuardado = Path.Combine(baseDirectory, @"..\..\..\..\Script\Data");

                // Asegurarse de que la ruta completa esté bien formada
                rutaGuardado = Path.GetFullPath(rutaGuardado);

                // Especificar el nombre del archivo (puedes adaptarlo según el formato deseado)
                string nombreArchivo = "ExcelSAP.xlsx"; // Cambiar según necesites

                // Combinar la ruta con el nombre del archivo
                string rutaArchivoFinal = Path.Combine(rutaGuardado, nombreArchivo);

                // Valores necesarios para utilizar
                for (int fila = 2; fila <= ultimaFila; fila++) // Empezamos desde la fila 2, asumiendo que la fila 1 es encabezados
                {
                    // Obtener valor de la columna Fecha y cambiar a puntos
                    string stringFecha = worksheet.Cell(fila, indiceColumnaFecha).GetString();
                    DateTime fecha = DateTime.Parse(stringFecha);
                    stringFecha = fecha.Date.ToString();
                    stringFecha = stringFecha.Replace('/', '.');
                    stringFecha = stringFecha.Split(' ')[0];

                    // Obtener tipo comprobante
                    char tipoMapeado = ExtraerTipoFactura(worksheet.Cell(fila, indiceColumnaTipo).GetString().Trim());
                    if (tipoMapeado == 'O')
                    {
                        MessageBox.Show("Error al mapear el tipo");
                    }

                    // Obtener valor del comprobante
                    string puntoVenta = worksheet.Cell(fila, indiceColumnaPuntoVenta).Value.ToString().PadLeft(5, '0');
                    if (puntoVenta.Count() > 5)
                    {
                        MessageBox.Show($"Error: el punto de  venta tiene {puntoVenta.Count()}");
                    }

                    string numeroComprobante = worksheet.Cell(fila, indiceColumnaComprobante).Value.ToString().PadLeft(8, '0');
                    if (numeroComprobante.Count() > 8)
                    {
                        MessageBox.Show($"Error: el numero de comprobante tiene {numeroComprobante.Count()}");
                    }

                    // Valor de referencia COMPLETO
                    string referencia = $"{puntoVenta}{tipoMapeado}{numeroComprobante}";

                    // Obtener valor de la columna Neto Gravado
                    double netoGravado = 0;
                    string valorCeldaNetoGravado = worksheet.Cell(fila, indiceColumnaNetoGravado).GetString();
                    string valorCeldaNetoGravadoSinComa = valorCeldaNetoGravado.Replace(",", ".");
                    if (valorCeldaNetoGravadoSinComa != "")
                    {
                        netoGravado = Math.Round(double.Parse(valorCeldaNetoGravadoSinComa, CultureInfo.InvariantCulture), 2);
                    }

                    // Obtener valor de la columna Neto No Gravado
                    double netoNoGravado = 0;
                    string valorCeldaNetoNoGravado = worksheet.Cell(fila, indiceColumnaNetoNoGravado).GetString();
                    string valorCeldaNetoNoGravadoSinComa = valorCeldaNetoNoGravado.Replace(",", ".");
                    if (valorCeldaNetoNoGravadoSinComa != "")
                    {
                        netoNoGravado = Math.Round(double.Parse(valorCeldaNetoNoGravadoSinComa, CultureInfo.InvariantCulture), 2);
                    }

                    // Obtener valor de la columna IVA
                    double IVA = 0;
                    string valorCeldaIVA = worksheet.Cell(fila, indiceColumnaIVA).GetString();
                    string valorCeldaIVASinComa = valorCeldaIVA.Replace(",", ".");
                    if (valorCeldaIVASinComa != "")
                    {
                        IVA = Math.Round(double.Parse(valorCeldaIVASinComa, CultureInfo.InvariantCulture), 2);
                    }

                    // Obtener valor de la columna Total
                    double total = 0;
                    string valorCeldaTotal = worksheet.Cell(fila, indiceColumnaImporteTotal).GetString();
                    string valorCeldaTotalSinComa = valorCeldaTotal.Replace(",", ".");
                    if (valorCeldaTotalSinComa != "")
                    {
                        total = Math.Round(double.Parse(valorCeldaTotalSinComa, CultureInfo.InvariantCulture), 2);
                        if (total == 0)
                        {
                            MessageBox.Show($"Error: el total de la fila {fila} es {total}");
                        }
                    }

                    // Obtener valor de la columna Texto
                    string valorCeldaTexto = worksheet.Cell(fila, indiceColumnaTexto).GetString().Trim();

                    // Obtener valor de la columna Indicador
                    string valorCeldaIndicador = worksheet.Cell(fila, indiceColumnaIndicador).GetString().Trim();

                    // Obtener valor de la cuenta mayor
                    string valorCuentaMayor = ExtraerCuentaMayor(valorCeldaTexto.Split('-')[1].ToUpper()).ToString().Trim();
                    if (valorCuentaMayor == "0")
                    {
                        MessageBox.Show("Error al mapear la cuenta mayor");
                    }

                    // Obtener valor codigo acreedor
                    string valorNombreAcreedor = worksheet.Cell(fila, indiceColumnaDenominacionEmisor).GetString().Trim();
                    string codigoAcreedor = ExtraerCodigoAcreedor(valorNombreAcreedor.ToUpper());
                    if (codigoAcreedor == "-1")
                    {
                        MessageBox.Show("Error al mapear el codigo de acreedor");
                        worksheetSalida.Cell(fila, 30).Value = "No esta en el diccionario de ACREEDORES";
                    }

                    // Obtener valor de centro costos
                    string valorCentroCosto = ExtraerCentroCosto(valorCeldaTexto.Split('-')[0].ToUpper()).ToString().Trim();
                    if (valorCentroCosto == "0")
                    {
                        MessageBox.Show("Error al mapear el centro costo");
                    }

                    // Obtener valor de la columna Fecha Contabilidad
                    string stringFechaContabilidad = worksheet.Cell(fila, indiceColumnaFecha).GetString();
                    DateTime fechaContabilidad = DateTime.Parse(stringFechaContabilidad);
                    string fechaContabilidadConvertida = "";
                    if (fechaContabilidad.Month < DateTime.Now.Month)
                    {
                        string añoFechaContabilidad = DateTime.Now.Date.Year.ToString();
                        string mesFechaContabilidad = DateTime.Now.Date.Month.ToString().PadLeft(2, '0');
                        fechaContabilidadConvertida = $"01.{mesFechaContabilidad}.{añoFechaContabilidad}";
                    }
                    else
                    {
                        fechaContabilidadConvertida = stringFechaContabilidad.Replace('/', '.');
                    }

                    // Neto + IVA
                    string netoConIva = Math.Round((IVA + netoGravado), 2).ToString();
                    netoConIva = netoConIva.Replace(",", ".");

                    // Tipo de venta
                    string tipoVenta = valorCeldaTexto.Split('-')[0].ToString();

                    // Categoria de venta
                    string categoriaVenta = valorCeldaTexto.Split('-')[1].ToString();

                    worksheetSalida.Cell(filaSalida, 1).Value = fecha;
                    worksheetSalida.Cell(filaSalida, 2).Value = worksheet.Cell(fila, indiceColumnaTipo).GetString();
                    worksheetSalida.Cell(filaSalida, 3).Value = worksheet.Cell(fila, indiceColumnaPuntoVenta).GetString();
                    worksheetSalida.Cell(filaSalida, 4).Value = worksheet.Cell(fila, indiceColumnaComprobante).GetString();
                    worksheetSalida.Cell(filaSalida, 5).Value = worksheet.Cell(fila, indiceColumnaComprobanteVentaHasta).GetString();
                    worksheetSalida.Cell(filaSalida, 6).Value = worksheet.Cell(fila, indiceColumnaCodigoAutorizacion).GetString();
                    worksheetSalida.Cell(filaSalida, 7).Value = worksheet.Cell(fila, indiceColumnaTipoDocEmisor).GetString();
                    worksheetSalida.Cell(filaSalida, 8).Value = worksheet.Cell(fila, indiceColumnaNroDocEmisor).GetString();
                    worksheetSalida.Cell(filaSalida, 9).Value = worksheet.Cell(fila, indiceColumnaDenominacionEmisor).GetString();
                    worksheetSalida.Cell(filaSalida, 10).Value = worksheet.Cell(fila, indiceColumnaTipoCambio).GetString();
                    worksheetSalida.Cell(filaSalida, 11).Value = worksheet.Cell(fila, indiceColumnaMoneda).GetString();
                    worksheetSalida.Cell(filaSalida, 12).Value = netoGravado.ToString(CultureInfo.InvariantCulture);
                    worksheetSalida.Cell(filaSalida, 13).Value = netoNoGravado.ToString(CultureInfo.InvariantCulture);
                    worksheetSalida.Cell(filaSalida, 14).Value = worksheet.Cell(fila, indiceColumnaOperacionesExentas).GetString();
                    worksheetSalida.Cell(filaSalida, 15).Value = worksheet.Cell(fila, indiceColumnaOtrosTributos).GetString();
                    worksheetSalida.Cell(filaSalida, 16).Value = IVA.ToString(CultureInfo.InvariantCulture);
                    worksheetSalida.Cell(filaSalida, 17).Value = total.ToString(CultureInfo.InvariantCulture);
                    worksheetSalida.Cell(filaSalida, 18).Value = valorCeldaTexto;
                    worksheetSalida.Cell(filaSalida, 19).Value = valorCeldaIndicador;
                    worksheetSalida.Cell(filaSalida, 20).Value = stringFecha;
                    worksheetSalida.Cell(filaSalida, 21).Value = valorCuentaMayor;
                    worksheetSalida.Cell(filaSalida, 22).Value = codigoAcreedor.ToString(CultureInfo.InvariantCulture);
                    worksheetSalida.Cell(filaSalida, 23).Value = referencia;
                    worksheetSalida.Cell(filaSalida, 24).Value = valorCentroCosto;
                    worksheetSalida.Cell(filaSalida, 25).Value = fechaContabilidadConvertida;
                    worksheetSalida.Cell(filaSalida, 26).Value = netoConIva.ToString(CultureInfo.InvariantCulture);
                    worksheetSalida.Cell(filaSalida, 27).Value = tipoVenta;
                    worksheetSalida.Cell(filaSalida, 28).Value = categoriaVenta;
                    worksheetSalida.Cell(filaSalida, 29).Value = tipoMapeado.ToString();

                    filaSalida++;
                }

                // Al final del ciclo, guarda el archivo de salida:
                workbookSalida.SaveAs(rutaArchivoFinal);

                MessageBox.Show("Transformación terminada");
            }
        }

        // Función para obtener el índice de una columna específica
        static int ObtenerIndiceColumna(IXLWorksheet worksheet, string nombreColumna)
        {
            int indiceColumna = -1;

            for (int col = 1; col <= worksheet.LastColumnUsed().ColumnNumber(); col++)
            {
                string valor = worksheet.Cell(1, col).GetString();

                if (valor.Equals(nombreColumna, StringComparison.OrdinalIgnoreCase))
                {
                    indiceColumna = col;
                    break;
                }
            }

            return indiceColumna;
        }
        static long ExtraerCuentaMayor(string nombre)
        {
            return ObtenerValorDeDiccionarioCuentaMayor(DiccionarioCuentasMayor.diccionarioCuentasMayor, nombre);
        }

        static long ExtraerCentroCosto(string tipo)
        {
            return ObtenerValorDeDiccionario(DiccionarioCuentasCentroCosto.diccionarioCuentasCentroCosto, tipo);
        }

        static char ExtraerTipoFactura(string tipo)
        {

            // Normalizar cadenas y extraer tipo y letra para comparación con diccionarioAFIP
            string normalizado = NormalizarAFIP(tipo);

            return ObtenerValorDeDiccionario(DiccionarioTiposFacturas.diccionarioTiposFacturas, normalizado);
        }

        static string ExtraerCodigoAcreedor(string nombre)
        {
            // Encontrar la clave más cercana en el diccionario usando Levenshtein, pero solo si está dentro del umbral
            var resultado = DiccionarioCodigosAcreedores.diccionarioCodigosAcreedores.Keys
                .Select(k => new { Key = k, Distancia = LevenshteinDistance(nombre, k) })
                .OrderBy(x => x.Distancia)
                .FirstOrDefault();

            if (resultado != null)
            {
                string nombreMapeado = resultado.Key;
                string codigo = DiccionarioCodigosAcreedores.diccionarioCodigosAcreedores[nombreMapeado].ToString();
                MessageBox.Show($"El nombre mapeado para {nombre} es {nombreMapeado} con el codigo {codigo}");
                return codigo;
            }
            else
            {
                Console.WriteLine("No se encontró un mapeo adecuado para el nombre.");
                return "-1";
            }
        }

        // Método que calcula la distancia de Levenshtein entre dos cadenas
        public static int LevenshteinDistance(string s, string t)
        {
            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            // Paso 1
            if (n == 0)
                return m;

            if (m == 0)
                return n;

            // Paso 2
            for (int i = 0; i <= n; d[i, 0] = i++) ;
            for (int j = 0; j <= m; d[0, j] = j++) ;

            // Paso 3
            for (int i = 1; i <= n; i++)
            {
                // Paso 4
                for (int j = 1; j <= m; j++)
                {
                    // Paso 5
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;

                    // Paso 6
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            }

            // Paso 7
            return d[n, m];
        }

        // Función para buscar valor en el diccionario
        static char ObtenerValorDeDiccionario(Dictionary<string, char> dictionary, string key)
        {
            if (dictionary.TryGetValue(key.Trim(), out char valor))
            {
                return valor;
            }
            return 'O'; // Valor por defecto o manejo de error según necesidad
        }

        static long ObtenerValorDeDiccionario(Dictionary<string, long> dictionary, string key)
        {
            key = key.RemoveWhiteSpaces();
            // Recorre todas las claves del diccionario
            foreach (var entry in dictionary)
            {
                // Verifica si la clave del diccionario contiene la palabra clave (ignorando mayúsculas)
                if (entry.Key.IndexOf(key.Trim(), StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return entry.Value; // Retorna el valor si encuentra una coincidencia parcial
                }
            }
            return 0; // Valor por defecto si no encuentra una coincidencia/ Valor por defecto o manejo de error según necesidad
        }

        static long ObtenerValorDeDiccionarioCuentaMayor(Dictionary<string, long> dictionary, string key)
        {
            // Recorre todas las claves del diccionario
            foreach (var entry in dictionary)
            {
                // Verifica si la clave del diccionario contiene la palabra clave (ignorando mayúsculas)
                if (entry.Key.IndexOf(key.Trim(), StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return entry.Value; // Retorna el valor si encuentra una coincidencia parcial
                }
            }
            return 0; // Valor por defecto si no encuentra una coincidencia
        }

        // Función para normalizar cadena de AFIP (extraer tipo y letra después del identificador numérico)
        static string NormalizarAFIP(string input)
        {
            int separatorIndex = input.IndexOf('-');
            if (separatorIndex >= 0)
            {
                return input.Substring(separatorIndex + 1).Trim();
            }
            return input.Trim();
        }

        private void buttonEjecutarScriptOP_Click(object sender, EventArgs e)
        {
            // Obtener la ruta del directorio donde está el ejecutable
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

            MessageBox.Show(baseDirectory);

            // Construir la ruta relativa al script dentro del repositorio
            string scriptRelativePath = Path.Combine(baseDirectory, @"..\..\..\..\Script\SCRIPT_SAP_OP.vbs");

            // Ejecutar el script
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = scriptRelativePath,
                UseShellExecute = true // Esto es importante para permitir la ejecución del script
            };

            try
            {
                Process.Start(startInfo);
                MessageBox.Show("Script ejecutado con éxito.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al ejecutar el script: " + ex.Message);
            }
        }
    }

    public static class Extensions
    {
        public static string RemoveWhiteSpaces(this string str)
        {
            return Regex.Replace(str, @"\s+", String.Empty);
        }
    }

    public static class DiccionarioCuentasMayor
    {
        public static Dictionary<string, long> diccionarioCuentasMayor { get; private set; }

        static DiccionarioCuentasMayor()
        {
            diccionarioCuentasMayor = new Dictionary<string, long>
            {
                {"GASTOS COMBUSTIBLES", 8001070200},
                {"GASTOS OBRA SOCIAL", 8001010202},
                {"GASTOS RODADOS", 8001150200},
                {"GASTOS HOTEL", 8001050000},
                {"GASTOS COMIDA", 8001050000},
                {"GASTOS GENERALES", 8001020100},
                {"GASTOS CORREO Y MENSAJERIA", 8001030200},
                {"GASTOS ELECTRICIDAD", 8001090200},
                {"GASTOS LOGISTICA", 8001320000},
                {"GASTOS MANTENIMIENTO INMUEBLES", 8001150200},
                {"GASTOS PUBLICIDAD Y PROPAGANDA - REGALOS EMPRESARIALES", 8001120000},
                {"GASTOS TELEFONIA E INTERNET", 8001030100},
                {"GASTOS CONSUMIBLES OFICINA", 8001040000},
                {"GASTOS ALQUILER DE VEHICHULOS", 8001080400},
                {"GASTOS ALQUILER OFICINAS", 8001080100},
                {"GASTOS ALQUILER DEPOSITO SLA", 8001080300},
                {"GASTOS HONORARIOS PROFESIONALES", 8001130000},
                {"GASTOS HONORARIOS AUDITORIA", 8001110000},
                {"GASTOS DIF DE CAMBIO NEGATIVA REALIZADA", 6603020300},
            };
        }

    }

    public static class DiccionarioCodigosAcreedores
    {
        public static Dictionary<string, long> diccionarioCodigosAcreedores { get; private set; }

        // Método para calcular el diccionario
        public static void CalcularDiccionario(string filePath)
        {
            if (diccionarioCodigosAcreedores == null)
            {
                diccionarioCodigosAcreedores = new Dictionary<string, long>();

                try
                {
                    using (StreamReader sr = new StreamReader(filePath, Encoding.GetEncoding(28591)))
                    {
                        string linea;
                        bool ban = false;
                        long valor = 0;
                        while ((linea = sr.ReadLine()) != null)
                        {
                            if (ban)
                            {
                                // Extraer nombre de la línea
                                ban = false;

                                // Dividir la cadena usando el carácter de tabulación
                                string[] partes = linea.Split('\t', StringSplitOptions.RemoveEmptyEntries);
                                string nombreSociedad = partes[0].Trim();

                                if (nombreSociedad.ToLower() == "empresa")
                                {
                                    ban = true;
                                    continue;
                                }

                                diccionarioCodigosAcreedores.Add(nombreSociedad, valor);

                                Console.WriteLine($"El nombre de la sociedad es {nombreSociedad}");
                            }
                            else if (linea.ToUpper().Contains("ACREEDOR") && linea.ToUpper().Contains("SOCIEDAD"))
                            {
                                // Extraer código acreedor
                                string[] partes = linea.Split('\t', StringSplitOptions.RemoveEmptyEntries);
                                string acreedor = partes[1].Trim();
                                valor = long.Parse(acreedor);
                                Console.WriteLine($"Código ACREEDOR: {acreedor}");
                            }
                            else if (linea.ToUpper().Contains("ACREEDOR") && linea.ToUpper().Contains("SECCIÓN"))
                            {
                                // Colocar bandera en true para que en la próxima línea extraiga el nombre de la sociedad
                                ban = true;
                            }
                        }
                        diccionarioCodigosAcreedores.Add("Fin", 1111);
                        Console.WriteLine("Terminado");
                    }
                }
                catch (Exception e)
                {
                    // Manejar excepciones si el archivo no se puede leer
                    Console.WriteLine("Ocurrió un error al leer el archivo:");
                    Console.WriteLine(e.Message);
                }
            }
        }
    }

    public static class DiccionarioCuentasCentroCosto
    {
        public static Dictionary<string, long> diccionarioCuentasCentroCosto { get; private set; }

        static DiccionarioCuentasCentroCosto()
        {
            diccionarioCuentasCentroCosto = new Dictionary<string, long>
            {
                {"VTAS", 5130010401},
                {"VTA", 5130010401},
                {"MKT", 5130010401},
                {"MRKT", 5130010401},
                {"ADM", 5130010802},
            };
        }

    }

    public static class DiccionarioTiposFacturas
    {
        public static Dictionary<string, char> diccionarioTiposFacturas { get; private set; }

        static DiccionarioTiposFacturas()
        {
            diccionarioTiposFacturas = new Dictionary<string, char>
            {
                {"Factura A", 'A'},
                {"FAC A", 'A'},
                {"FA A", 'A'},
                {"Factura B", 'B'},
                {"FAC B", 'B'},
                {"FA B", 'B'},
                {"Factura C", 'C'},
                {"FAC C", 'C'},
                {"FA C", 'C'},
                {"Nota de Crédito A", 'A'},
                {"NC A", 'A'},
                {"Nota de Débito A", 'A'},
                {"ND A", 'A'},
                {"Recibo A", 'A'},
                {"Recibo B", 'B'},
                {"Recibo C", 'C'}
            };
        }

    }
}

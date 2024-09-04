using System.Diagnostics;
using System.Globalization;
using ClosedXML.Excel;

namespace FB60_SAP
{
    public static class DiccionarioCuentasMayor
    {
        public static Dictionary<string, long> diccionarioCuentasMayor { get; private set; }

        static DiccionarioCuentasMayor()
        {
            diccionarioCuentasMayor = new Dictionary<string, long>
            {
                {"GASTOS COMBUSTIBLES", 8001070200},
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

        static DiccionarioCodigosAcreedores()
        {
            diccionarioCodigosAcreedores = new Dictionary<string, long>
            {
                {"AFIP",  3000030},
                {"ARRESE LUCAS", 2001713},
                {"MALDONADO PETRONA", 2001712},
                {"PISCITELLI ALESSANDRO SEBASTIAN", 2001712},
                {"JOSE LUIS CHIAVARINI",  2001710},
                {"JORGE LUIS CASTRO", 2001709},
                {"CONTI ALBERTO JUAN",  2001708},
                {"ZLAUVINEN HUGO ADOLFO", 2001707},
                {"GARIONE GABREIL HORACIO", 2001706},
                {"GIAVINO JORGE LUIS", 2001705},
                {"DEL RIO ERNESTO LUIS", 2001704},
                {"DE LA ROSA HAZELHOFF CARINA MARCELA", 2001703},
                {"ROSSI GABRIEL SALVADOR",  2001702},
                {"SOSA MARIA PAOLA", 2001701},
                {"ROMANO GUSTAVO ANDRES", 2001700},
                {"SIDICARIO LUIS CARLOS", 2001699},
                {"FERNANDEZ LUIS ESTEBAN",  2001698},
                {"MADDIO JUAN JOSE", 2001697},
                {"ANGEL GILDA ROMINA DEL VALLE", 2001696},
                {"CADIMA GARCIA MAURICIO JAVIER", 2001695},
                {"PORRO SILVIA SUSANA", 2001694},
                {"PEREZ MARIA    JORGELINA", 2001693},
                {"CAZZUCHELLI LUIS BENITO", 2001692},
                {"JUNQUEIRA MARIANO", 2001691},
                {"GARETTO PABLO ALEJANDRO", 2001669},
                {"CUADRADO MORENO FRANCISCO", 2001636},
                {"RIOS STON LUCAS MARTIN", 2001632},
                {"LOSADA FERNANDOANTONIO",  2001631},
                {"GIGON RAMON", 2001630},
                {"STOCHERO PABLO LUIS", 2001619},
                {"山东润丰农科有限公司（香港） ", 2010},
                {"润丰农科香港（阿根廷）有限公司", 2240},
                {"RAINBOW AGROSCIENCES S.A. (RAAR)", 5070},
                {"AGROTERRUM S.A. (Argentina)", 5130},
                {"LPAR LTD", 1003782},
                {"CTAR LTD", 1003788},
                {"YPF. S.A", 1005261},
                {"UNION AGRICOLA AVELLANEDA COOP LTDA", 1005262},
                {"DASER AGRO SA", 1005280},
                {"SCALELLA IGNACIO D", 1005293},
                {"MAERSK LINE", 1005303},
                {"EXOLGAN SA", 1005304},
                {"MIRANDA DIEGO JAVIER", 1005305},
                {"RAYA EDGARDO HORACIO", 1005306},
                {"ASTE GUILLERMO FEDERICO", 1005310},
                {"SENASA", 1005312},
                {"LOGISTICA LA FLORIDA SA", 1005314},
                {"MEDITERRANEAN SHIPPING COMPANY S.A.", 1005315},
                {"TIBONI & CIA SA", 1005317},
                {"FRAVEGA S.A.C.I.E.I", 1005318},
                {"PLASTICOS POL NOR S.R.L.", 1005320},
                {"PLAYAS SUBTERRANEAS S.A.", 1005321},
                {"TRANSPORTE MyC S.R.L.", 1005322},
                {"TELECOM PERSONAL S.A.", 1005323},
                {"IRUSTIA LILIANA CARINA", 1005324},
                {"MERCOCARGA S.A.", 1005328},
                {"TERMINAL 4 S A", 1005329},
                {"TRANSPORTE MOSTTO EDUARDO E", 1005333},
                {"AEROPUERTOS ARGENTINA 2000 S A ", 1005335},
                {"DEHEZA SOCIEDAD ANONIMA INDUSTRIAL COMER", 1005336},
                {"DE SALA ANDRES ANTONIO", 1005338},
                {"TERMINALES RIO DE LA PLATA SOCIEDAD ANON", 1005340},
                {"OPERADORA DE ESTACIONES DE SERVICIO SA.", 1005348},
                {"SEVEN TRADES S.A.", 1005351},
                {"LAKAUT S.A.", 1005352},
                {"EXSEN S.A", 1005357},
                {"OSDE ORG DE SERVICIOS DIRECTOS EMP", 1005359},
                {"PAN AMERICAN ENERGY LLC SUC ARG", 1005360},
                {"TELECOM ARGENTINA SA (Internet,Cel Pers)", 1005362},
                {"CORREO ANREANI SA.", 1005370},
                {"BANK S.A.", 1005374},
                {"INTER AMERICAN CARGO GROUP SA.", 1005381},
                {"INTERBANKING S.A.", 1005397},
                {"ASOCIACION CIVIL CAMPO LIMPIO SGE", 1005437},
                {"MOTOS AIR", 1005456},
                {"EVERGREEN SHIPPING AGENCY ARG.SA.", 1005520},
                {"EL REGRESO SRL", 1005530},
                {"EL MIRADOR SRL", 1005538},
                {"GEMEZ SA.", 1005551},
                {"LUDMAN LEANDRO ARIEL", 1005572},
                {"MAYNAR A G S.A", 1005623},
                {"ARGENTUR INVERSIONES TURISTICAS S.A", 1005626},
                {"NEWTRAL S.A", 1005644},
                {"HOTEL PRESIDENTE S.A", 1005650},
                {"GONZALES CARLOS GUSTAVOY GONZALES LEONA", 1005652},
                {"JOSE HERMANOS SRL", 1005653},
                {"GRUPO LAVALLE SRL", 1005654},
                {"SERVICIOS TARALLI SA", 1005655},
                {"AVENIDA SRL", 1005656},
                {"EL CRUCE SRL", 1005657},
                {"APODO SA", 1005658},
                {"PETRO AVELLANEDA SRL", 1005659},
                {"RESTO, BAR & LOUNGE SRL", 1005660},
                {"INTERTEL SRL", 1005661},
                {"EL FOGON CRIOLLO SRL", 1005663},
                {"UNIDADES EJECUTORAS AUT AP01", 1005665},
                {"BURGENER JAVIER Y ROGGERO HUGO SH", 1005666},
                {"GENIAL SA", 1005667},
                {"CAMAROTTI JAVIER Y RODRIGO SH", 1005668},
                {"FGC FUELS MARKETING SA", 1005669},
                {"ESTANCIA GRANDE SRL", 1005670},
                {"TOSELLI HERMANOS SA", 1005671},
                {"JUSTO SA", 1005672},
                {"MORANDIN MARIANA Y MORANDIN MARCOS", 1005673},
                {"SWISS MEDICAL SA", 1005674},
                {"FADEL SA", 1005675},
                {"BAR Y COMEDOR KM 256", 1005676},
                {"VILLAGAS SA", 1005677},
                {"FURFANTO SA", 1005678},
                {"ESSA SRL", 1005679},
                {"EL TOKIO VIEJO SRL", 1005680},
                {"YAPAI SRL", 1005681},
                {"WESPYT SRL", 1005682},
                {"LEVEAL SA", 1005683},
                {"PETROSERVICE SRL", 1005684},
                {"AZUL COMBUSTIBLES SA", 1005685},
                {"FRANMA SRL", 1005686},
                {"LORETO COMBUSTIBLES SRL", 1005687},
                {"CHEZ CAFE SAS", 1005688},
                {"FIDESUR SA", 1005689},
                {"JPC OIL SRL", 1005690},
                {"MARIA M B DE SOLDANO Y CIA SRL", 1005691},
                {"LA PIOJERA SRL", 1005692},
                {"B Y J COMBUSTIBLES SRL", 1005693},
                {"FOURMAX SRL", 1005695},
                {"SOC HOTELERA VILLA MARIA SA", 1005696},
                {"LUAJUMA SAS", 1005697},
                {"ALT SA", 1005698},
                {"A M GAS SA", 1005699},
                {"EL PUERTO COMBUSTIBLES SA", 1005700},
                {"SERVICENTRO SANTO TOME SRL", 1005701},
                {"ESCOBAR AUTOMOTORES SA", 1005702},
                {"D-COM SRL", 1005703},
                {"SAN JAVIER SERVICIOS SRL", 1005704},
                {"CRESPO SERVICIOS SA", 1005705},
                {"EL CRUCE SA", 1005706},
                {"DON ANTONIO COMBUSTIBLES SRL", 1005707},
                {"LA COSTERITA SRL", 1005708},
                {"AZUL HOTEL SRL", 1005709},
                {"DANMAZ SA", 1005710},
                {"L P LOS GRINGOS SRL", 1005711},
                {"DT3 COMBUSTIBLES SA", 1005712},
                {"BORTOLON Y URQUIZA COMB SRL", 1005713},
                {"CATANGE HOTEL SRL", 1005714},
                {"MONTAGNE OUTDOORS SA", 1005715},
                {"CASINO MELINCUE SA", 1005716},
                {"SALVADOR DI STEFANO SRL", 1005717},
                {"SANCOR COOPERATIVA DE SEGUROS LTDA", 1005718},
                {"EL MARQUES SACI", 1005719},
                {"BIDA SRL", 1005720},
                {"FOOD PATAGONIA SA", 1005721},
                {"EL CHARRUA SRL", 1005723},
                {"DESPEGAR.COM.AR SA", 1005724},
                {"LOMFAKO SA", 1005725},
                {"ALISON LEANDRO DANIEL, ALISON RODRIGO M", 1005726},
                {"DON CALIFA COMBU DE BOLIVAR SA", 1005727},
                {"EMECO SA", 1005728},
                {"INDOSTAN SA", 1005729},
                {"CIRI SAS", 1005730},
                {"LAVIANA COMBUSTIBLES SA", 1005731},
                {"AMX ARGENTINA SA", 1005732},
                {"AGUAS SANTAFESINAS SA", 1005733},
                {"PAITUBI SERVICIOS SRL", 1005734},
                {"LAB DE INV Y DESARROLLO SA", 1005735},
                {"MARAPE SA", 1005736},
                {"COIHUE SRL", 1005737},
                {"PORTAL CERES HOTEL SRL", 1005738},
                {"REAMA SRL", 1005739},
                {"ARCOS DORADOS ARG SA", 1005740},
                {"SPAHN ROBERTO Y SPHAN GISELA SH", 1005741},
                {"LA AGRICOLA REGIONAL COOP", 1005742},
                {"RIO DE JANEIRO SA", 1005743},
                {"PONTEC SA", 1005744},
                {"RIO SIL SA", 1005745},
                {"VASE SAS", 1005746},
                {"AGROSITIO", 1005747},
                {"BELFIORI DANTE Y RUIZ DIAZ MIGUEL S.H.", 1005748},
                {"IRSA INVERSIONES Y REPRESENTACIONES SA", 1005749},
                {"PARKING MALL SA", 1005750},
                {"SORBONA COMBUTIBLES SA", 1005751},
                {"TRANSFLUVIAL SA", 1005752},
                {"CRUZER SRL", 1005753},
                {"LAS URBANAS SRL", 1005754},
                {"XECA SA", 1005755},
                {"CAFE SARMIENTO SRL", 1005756},
                {"MANDOLA HERMANOS SA", 1005757},
                {"YPF SUNCHALES SRL", 1005758},
                {"AGENTES ROMERO SA", 1005759},
                {"HUGO BAUTISTA BALBI SA", 1005760},
                {"RAPIFLET CAROLINA", 1005761},
                {"DI TONDO JOSE E HIJOS", 1005762},
                {"GREAT MEALS SA", 1005763},
                {"ARROYO DUAL SRL", 1005764},
                {"BERRIA ENERGIA SA", 1005765},
                {"QUINTIN RUDECINDO E HIJOS SRL", 1005766},
                {"SUCESORES DE JOSE ERCOLE GRASSO SH", 1005768},
                {"ROMA COMBUSTIBLES SA", 1005769},
                {"COMBUSTIBLES VILLA MARIA SA", 1005770},
                {"TRANSFENOR SRL", 1005771},
                {"PETRORAFAELA SRL", 1005772},
                {"DON LOBO SAS", 1005773},
                {"GESA SA", 1005774},
                {"MAZZON Y CIA SOCIEDAD COLECTIVA", 1005775},
                {"HTL GROUP SA", 1005776},
                {"ALCIDES E PHISALIX SA", 1005777},
                {"HDI SEGUROS SA", 1005778},
                {"ASOC MUTUAL SANCOR SALUD", 1005779},
                {"AUTOPISTAS DE BS AS S.A.", 1005780},
                {"IKIGAI SERVICIOS SRL", 1005781},
                {"PUKEN MEDIA SA", 1005782},
                {"CONAR LATINOAMERICA SRL", 1005783},
                {"PINTURERIAS LAURENTI SA", 1005784},
                {"SANTA BARBARA SA", 1005785},
                {"BIG FISH SA", 1005786},
                {"EL SOLAR DE SIANCAS SA", 1005787},
                {"ERMAYA SRL", 1005788},
                {"ESTACION DE SERVICIO RUTA 34", 1005789},
                {"JUANA RESTO SRL", 1005790},
                {"LOZADA Y NOVILLO SRL", 1005791},
                {"SINDICATO PETROLERO DE CORDOBA", 1005792},
                {"CORREDORES VIALES SA", 1005793},
                {"VALLE MARIA SRL", 1005794},
                {"MARCELO GOTTIG Y CIA SA", 1005795},
                {"COFFEE SAS", 1005796},
                {"SANTA VICTORIA SRL", 1005797},
                {"AROMAS DE ARGENTINA SRL", 1005798},
                {"DEBONA MARCELO FABIAN Y DEBONA VICTOR HU", 1005799},
                {"ASOC MUTUAL EMPL CORR PRIV", 1005800},
                {"PARADOR SAN PEDRO", 1005801},
                {"RIVARA SA", 1005802},
                {"CORP ECONOM PAMPEANA SA", 1005803},
                {"PETROPRINGLES SA", 1005804},
                {"NEW LAURENTS SA", 1005805},
                {"ENERGY CLEAN SA", 1005806},
                {"ANDREA CLAUDIA MOSSIO Y EDITH MARIA MOSS", 1005807},
                {"GROSCAN SRL", 1005808},
                {"EMPRENDIMIENTOS CENTRALES SA", 1005809},
                {"PARKING DEL CENTRO SA ", 1005810},
                {"PORTO RESTO SA", 1005811},
                {"CAO SRL", 1005812},
                {"SANTA ROSA COMBUSTIBLES SRL", 1005813},
                {"DISTRIBUIDORA BULONES COIRO SA", 1005814},
                {"RIAMBER SA", 1005815},
                {"KARMI 13 SA", 1005816},
                {"AUTOPISTAS DEL SOL SA", 1005817},
                {"AUTOPISTRAS URBANAS SA", 1005818},
                {"FUTUROS SRL", 1005819},
                {"AGRODIGITAL SAS", 1005820},
                {"CARIVANA SRL", 1005821},
                {"EST DE SERV SAGITARIO SRL", 1005822},
                {"ARESTE SAS", 1005823},
                {"EL CUARTITO SA", 1005824},
                {"PARRILLA PENA SRL", 1005825},
                {"GRUPO SAN GUILLERMO", 1005826},
                {"TRANSPORTADORA PAMPEANA SRL", 1005827},
                {"BIFUEL SRL", 1005828},
                {"DON EDUARDO SRL", 1005829},
                {"BUENAVISTA SA", 1005830},
                {"CURMONA ALEJANDRO", 1005831},
                {"PELUNCHA SA", 1005832},
                {"BRIMAK SRL", 1005833},
                {"IRONCHAC SA", 1005834},
                {"SOL DEL RIO SRL", 1005835},
                {"LA MARINA LOBOS SRL", 1005836},
                {"FRANCOU H, FRANCOU O GUIFRE M", 1005837},
                {"AL FUEGO SRL", 1005838},
                {"LA RURAL VYP SAS", 1005839},
                {"CHARATA COMBUSTIBLES SACI", 1005840},
                {"EL GUAYACAN SA", 1005841},
                {"EXPRESO SANTA ROSA", 1005842},
                {"BCO DE LA NAC ARGENTINA", 1005843},
                {"CORMORAN SA", 1005844},
                {"EST DE SERV GRAL GUEMES SRL", 1005845},
                {"EMPRENDIMIENTOS GASTRONOMICOS SRL", 1005846},
                {"RUTAS SERRANAS SRL", 1005847},
                {"INTEGRACION GALLO SA", 1005848},
                {"MY BEIYET SRL", 1005849},
                {"LA CUARTA SA", 1005850},
                {"CETROGAR SA", 1005851},
                {"SAGLIK SA", 1005852},
                {"MATOV NEG INMO SRL", 1005853},
                {"ITIN SA", 1005854},
                {"ALTO VERDE SA", 1005855},
                {"MOHA SA", 1005856},
                {"JAUREGUI Y MORALES SRL", 1005857},
                {"BERAZA Y CIA SRL", 1005858},
                {"FLORENCIO B BARBERO SA", 1005859},
                {"MATE JUSTO SRL", 1005860},
                {"MUSIC HOUSE CORP SA ", 1005861},
                {"ELECTRONICA MEGATONE SA", 1005862},
                {"CAMPO MAS SAS", 1005863},
                {"ALGAR SRL", 1005864},
                {"NUESTRO SABORES SRL", 1005865},
                {"PA QUE LE GUSTE SIMPLE ASOC", 1005866},
                {"ARROYITO SERVICIOS SA", 1005867},
                {"CRECIENDO SA", 1005868},
                {"ESTACION VISION SRL", 1005869},
                {"SINERCOM SRL", 1005870},
                {"MAZZON HNOS SRL", 1005871},
                {"LA EXELENCIA SRL ", 1005872},
                {"GRISANTI SERVICIOS SA", 1005873},
                {"STRIANESE MOTORS SA", 1005874},
                {"VENTAS Y SERVICIOS SA", 1005875},
                {"COCEMAR SRL", 1005876},
                {"DIQUESUR SA", 1005877},
                {"GOLDEN ZAR SAS", 1005878},
                {"ESTACION DE SERVICIO FLORENCIA SRL", 1005879},
                {"TECNICAS FERROVIARIAS ARGENTINAS SA", 1005880},
                {"FRANFOOD SA", 1005881},
                {"GRUPO MOVIL SRL", 1005882},
                {"GMRA SA", 1005883},
                {"CAMINO DE LAS SIERRAS SA", 1005884},
                {"COMBUSTIBLES SIMOSA SRL", 1005885},
                {"LAS LEGENDARIAS S.R.L ", 1016908},
                {"VALLE MOTORS S.R.L ", 1016920},
                {"ALTOS DEL BERMEJO S.R.L  ", 1016918},
                {"SALTA SUR S.A. ", 1016911},
                {"ASTILLERO S.R.L ", 1016916},
                {"SEVEN S.A.S", 1016930},
            };
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
 
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // Establecer el estilo del borde y deshabilitar el cambio de tamaño
            this.FormBorderStyle = FormBorderStyle.FixedSingle;

            // Establecer el tamaño mínimo y máximo para evitar el cambio de tamaño
            this.MinimumSize = this.MaximumSize = this.Size;
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

            // Construir la ruta relativa al script dentro del repositorio
            string scriptRelativePath = Path.Combine(baseDirectory, @"Script\Script.vbs");

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
                    char tipoMapeado = ExtraerTipoFactura(worksheet.Cell(fila, indiceColumnaTipo).GetString());
                    if (tipoMapeado == 'O')
                    {
                        MessageBox.Show("Error al mapear el tipo");
                    }

                    // Obtener valor del comprobante
                    string puntoVenta = worksheet.Cell(fila, indiceColumnaPuntoVenta).Value.ToString().PadLeft(5, '0');
                    string numeroComprobante = worksheet.Cell(fila, indiceColumnaComprobante).Value.ToString().PadLeft(8, '0');

                    // Valor de referencia COMPLETO
                    string referencia = $"{puntoVenta}{tipoMapeado}{numeroComprobante}";

                    // Obtener valor de la columna Neto Gravado
                    string valorCeldaNetoGravado = worksheet.Cell(fila, indiceColumnaNetoGravado).GetString();
                    string valorCeldaNetoGravadoSinComa = valorCeldaNetoGravado.Replace(",", ".");
                    double netoGravado = double.Parse(valorCeldaNetoGravadoSinComa, CultureInfo.InvariantCulture);

                    // Obtener valor de la columna Neto No Gravado
                    string valorCeldaNetoNoGravado = worksheet.Cell(fila, indiceColumnaNetoNoGravado).GetString();
                    string valorCeldaNetoNoGravadoSinComa = valorCeldaNetoNoGravado.Replace(",", ".");
                    if (valorCeldaNetoNoGravadoSinComa != "")
                    {
                        double netoNoGravado = double.Parse(valorCeldaNetoNoGravadoSinComa, CultureInfo.InvariantCulture);
                    }             

                    // Obtener valor de la columna IVA
                    string valorCeldaIVA = worksheet.Cell(fila, indiceColumnaIVA).GetString();
                    string valorCeldaIVASinComa = valorCeldaIVA.Replace(",", ".");
                    double IVA = double.Parse(valorCeldaIVASinComa, CultureInfo.InvariantCulture);

                    // Obtener valor de la columna Total
                    string valorCeldaTotal = worksheet.Cell(fila, indiceColumnaImporteTotal).GetString();
                    string valorCeldaTotalSinComa = valorCeldaTotal.Replace(",", ".");
                    double total = double.Parse(valorCeldaTotalSinComa, CultureInfo.InvariantCulture);

                    // Obtener valor de la columna Texto
                    string valorCeldaTexto = worksheet.Cell(fila, indiceColumnaTexto).GetString();

                    // Obtener valor de la columna Indicador
                    string valorCeldaIndicador = worksheet.Cell(fila, indiceColumnaIndicador).GetString();

                    // Obtener valor de la cuenta mayor
                    string valorCuentaMayor = ExtraerCuentaMayor(valorCeldaTexto.Split('-')[1]).ToString();
                    if (valorCuentaMayor == "-1")
                    {
                        MessageBox.Show("Error al mapear la cuenta mayor");
                    }

                    // Obtener valor codigo acreedor
                    string valorNombreAcreedor = worksheet.Cell(fila, indiceColumnaDenominacionEmisor).GetString();
                    long codigoAcreedor = ExtraerCodigoAcreedor(valorNombreAcreedor);
                    if (codigoAcreedor == -1)
                    {
                        MessageBox.Show("Error al mapear el codigo de acreedor");
                    }

                    // Obtener valor de centro costos
                    string valorCentroCosto = ExtraerCentroCosto(valorCeldaTexto.Split('-')[0]).ToString();
                    if (valorCentroCosto == "-1")
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
                    string netoConIva = (IVA + netoGravado).ToString();

                    // Tipo de venta
                    string tipoVenta = valorCeldaTexto.Split('-')[0].ToString();

                    // Categoria de venta
                    string categoriaVenta = valorCeldaTexto.Split('-')[1].ToString();

                    MessageBox.Show($"Fecha convertida:{stringFecha}, Cuenta mayor:{valorCuentaMayor}, Referencia:{referencia}, Centro costo:{valorCentroCosto}, Codigo acreedor: {codigoAcreedor}, Fecha contabilidad:{fechaContabilidadConvertida}, Neto con IVA:{netoConIva.Replace(',', '.')}, Tipo Venta:{tipoVenta}, Categoria Venta: {categoriaVenta}");
                }
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

        static long ExtraerCodigoAcreedor(string nombre)
        {
            // Encontrar la clave más cercana en el diccionario usando Levenshtein, pero solo si está dentro del umbral
            var resultado = DiccionarioCodigosAcreedores.diccionarioCodigosAcreedores.Keys
                .Select(k => new { Key = k, Distancia = LevenshteinDistance(nombre, k) })
                .OrderBy(x => x.Distancia)
                .FirstOrDefault();

            if (resultado != null)
            {
                string nombreMapeado = resultado.Key;
                long codigo = DiccionarioCodigosAcreedores.diccionarioCodigosAcreedores[nombreMapeado];
                MessageBox.Show($"El nombre mapeado para {nombre} es {nombreMapeado} con el codigo {codigo}");
                return codigo;
            }
            else
            {
                Console.WriteLine("No se encontró un mapeo adecuado para el nombre.");
                return -1;
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
            if (dictionary.TryGetValue(key, out char valor))
            {
                return valor;
            }
            return 'O'; // Valor por defecto o manejo de error según necesidad
        }

        static long ObtenerValorDeDiccionario(Dictionary<string, long> dictionary, string key)
        {
            if (dictionary.TryGetValue(key, out long value))
            {
                return value;
            }
            return -1; // Valor por defecto o manejo de error según necesidad
        }

        static long ObtenerValorDeDiccionarioCuentaMayor(Dictionary<string, long> dictionary, string key)
        {
            if (dictionary.TryGetValue(key, out long value))
            {
                return value;
            }
            return 8001020100; // Valor por defecto o manejo de error según necesidad
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
    }
}

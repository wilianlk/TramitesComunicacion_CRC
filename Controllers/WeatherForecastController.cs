using Microsoft.AspNetCore.Mvc;
using System.Collections.Concurrent;
using System.Formats.Asn1;
using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using TramitesComunicacion_CRC.Services;
using System.ComponentModel;
using Newtonsoft.Json;
using OfficeOpenXml;



namespace TramitesComunicacion_CRC.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateOnly.FromDateTime(DateTime.Now.AddDays(index)),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToArray();
        }

        [HttpGet("consultar-por-telefono")]
        public async Task<IActionResult> ConsultarPorTelefono()
        {
            string rutaArchivo = @"C:\Recamier\data.csv";
            try
            {
                var telefonos = await LeerTelefonosDesdeCsvAsync(rutaArchivo);
                if (telefonos == null || telefonos.Length == 0)
                {
                    return BadRequest("No se encontraron teléfonos en el archivo.");
                }

                var webServiceClient = new WebServiceClient();
                string resultByPhone = await webServiceClient.ConsultarRnePorTelefonoAsync(telefonos);
                ExportarJsonAExcel(resultByPhone, @"C:\Recamier\ResultadosConsulta.xlsx");

                return Ok();
            }
            catch (FileNotFoundException ex)
            {
                return NotFound("Archivo no encontrado: " + ex.Message);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Error al procesar el archivo: " + ex.Message);
            }
        }
        public static async Task<string[]> LeerTelefonosDesdeCsvAsync(string rutaArchivo)
        {
            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ",",
                HasHeaderRecord = false
            };

            var telefonos = new ConcurrentBag<string>();

            try
            {
                using (var reader = new StreamReader(rutaArchivo))
                using (var csv = new CsvReader(reader, config))
                {
                    while (await csv.ReadAsync())
                    {
                        var telefono = csv.GetField<string>(0);
                        telefonos.Add(telefono);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al leer el archivo CSV: {ex.Message}");
                throw; 
            }

            return telefonos.ToArray();
        }
        public static void ExportarJsonAExcel(string jsonData, string rutaArchivo)
        {
            // Deserializar los datos JSON en una lista de diccionarios
            var datos = JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(jsonData);

            // Configurar la licencia de EPPlus
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // Crear el archivo de Excel
            FileInfo file = new FileInfo(rutaArchivo);

            // Si el archivo existe, eliminarlo
            if (file.Exists)
            {
                file.Delete();
            }

            // Crear un nuevo paquete de Excel
            using (var package = new ExcelPackage(file))
            {
                // Añadir una nueva hoja de trabajo al paquete
                var worksheet = package.Workbook.Worksheets.Add("Datos");

                // Obtener las claves de los diccionarios para usarlas como encabezados de las columnas
                int columnIndex = 1;
                foreach (var key in datos[0].Keys)
                {
                    if (key == "opcionesContacto")
                    {
                        worksheet.Cells[1, columnIndex].Value = "Sms";
                        columnIndex++;
                        worksheet.Cells[1, columnIndex].Value = "Aplicacion";
                        columnIndex++;
                        worksheet.Cells[1, columnIndex].Value = "Llamada";
                    }
                    else
                    {
                        string header = key == "llave" ? "Telefono" : Capitalize(key);
                        worksheet.Cells[1, columnIndex].Value = header;
                    }
                    columnIndex++;
                }

                // Añadir la columna "Fecha Consultada"
                worksheet.Cells[1, columnIndex].Value = "Fecha Consultada";

                // Obtener la fecha y hora actuales
                string fechaConsultada = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // Escribir los datos en las filas de la hoja de trabajo
                int rowIndex = 2;
                foreach (var item in datos)
                {
                    columnIndex = 1;
                    foreach (var key in item.Keys)
                    {
                        if (key == "opcionesContacto")
                        {
                            var opciones = JsonConvert.DeserializeObject<Dictionary<string, bool>>(item[key].ToString());
                            worksheet.Cells[rowIndex, columnIndex].Value = opciones["sms"] ? "V" : "F";
                            columnIndex++;
                            worksheet.Cells[rowIndex, columnIndex].Value = opciones["aplicacion"] ? "V" : "F";
                            columnIndex++;
                            worksheet.Cells[rowIndex, columnIndex].Value = opciones["llamada"] ? "V" : "F";
                        }
                        else
                        {
                            worksheet.Cells[rowIndex, columnIndex].Value = item[key]?.ToString();
                        }
                        columnIndex++;
                    }
                    // Agregar la fecha consultada a la última columna
                    worksheet.Cells[rowIndex, columnIndex].Value = fechaConsultada;
                    rowIndex++;
                }

                // Guardar el paquete de Excel
                package.Save();
            }
        }
        public static string Capitalize(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            text = text.ToLower();
            string[] words = text.Split(new char[] { ' ', '_', '-' }, System.StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < words.Length; i++)
            {
                words[i] = words[i].Substring(0, 1).ToUpper() + words[i].Substring(1);
            }
            return string.Join(" ", words);
        }

    }
}

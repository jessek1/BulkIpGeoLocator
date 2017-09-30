using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using MaxMind.GeoIP2;
using MaxMind.GeoIP2.Responses;

namespace BulkIpGeoLocator
{
    public class GeoLocator
    {
        private static readonly HttpClient _httpClient = new HttpClient();
        private const int MAX_TRIES = 10;

        public const string DB_CITY_PATH = "E:\\src\\BulkIpGeoLocator\\geolite2-city.mmdb";
        public const string DB_ASN_PATH = "E:\\src\\BulkIpGeoLocator\\geolite2-asn.mmdb";
        private static DatabaseReader cityReader = new DatabaseReader(DB_CITY_PATH);
        private static DatabaseReader asnReader = new DatabaseReader(DB_ASN_PATH);


        public async Task RunGeoLocator(Excel.Worksheet worksheet, int rowNumber, string IpColLetter, int tries = 0)
        {
           
            String value;
            try
            {
                Excel.Range excelCell = (Excel.Range)worksheet.get_Range(IpColLetter + rowNumber, IpColLetter + rowNumber);
                value = (String)excelCell.Value2;
                if (value.IndexOfAny(new char[] { '.' }) == -1)
                    return;
            }
            catch (Exception )
            {
                return;
            }

            if (!String.IsNullOrEmpty(worksheet.Cells[rowNumber, 4 + IP].Value<String>()))
            {
                Console.WriteLine(String.Format("Row #{0} - Skipping, already processed.", rowNumber));
                return;
            }
            //using (var _httpClient = new HttpClient())
            //{
            /*HttpResponseMessage response = null;*/
            CityResponse city = null;
            AsnResponse asn = null;
            try
            {
                /*
                _httpClient.DefaultRequestHeaders.ConnectionClose = false;
                
                response = await _httpClient.GetAsync(new Uri("https://tools.keycdn.com/geo.json?host=" + value)).ConfigureAwait(false);
                response.EnsureSuccessStatusCode();
                */

                city = cityReader.City(value);
                asn = asnReader.Asn(value);

                if (city != null)
                {
                    UpdateExcel2(worksheet, value, city, asn, rowNumber, 4);
                    Console.WriteLine(String.Format("Row #{0} - IP: {1} ... {2}, {3} ", rowNumber, value, city.City.Name, city.Country.Name));
                }
                else
                {
                    Console.WriteLine(String.Format("Row #{0} - IP: {1} ... Null received from MaxMind", rowNumber, value));
                }

            }
            catch (Exception ex)
            {
                
                Console.WriteLine(String.Format("Row #{0} - IP: {1} ... ERR(#{3}): {2}", rowNumber, value, ex.Message, tries));
                /*
                if (tries < MAX_TRIES) {
                    await RunGeoLocator(worksheet, rowNumber, IpColLetter, tries++);
                } else
                {
                    Console.WriteLine(String.Format("Row #{0} - IP: {1} ... FAILED after {2} tries.", rowNumber, value, MAX_TRIES));
                    
                }
                */

                return;
            }

            /*
            if (response.IsSuccessStatusCode)
            {
                var jResult = JObject.Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
                UpdateExcel(worksheet, jResult, rowNumber, 5);

                Console.WriteLine(String.Format("Row #{0} - IP: {1} ... {2}, {3} ", rowNumber, value, jResult["data"]["geo"]["city"].Value<String>(), jResult["data"]["geo"]["country_name"].Value<String>()));
            }
            else
            {
                Console.WriteLine("Failed : {0} - {1}", response.StatusCode, await response.Content.ReadAsStringAsync().ConfigureAwait(false));
            }
            */
            
           
            ///}


            return;
        }

        const int HOST = 1;
        const int IP = 2;
        const int RDNS = 3;
        const int ASN = 4;
        const int ISP = 5;
        const int COUNTRY_NAME = 6;
        const int COUNTRY_CODE = 7;
        const int REGION = 8;
        const int CITY = 9;
        const int POSTAL_CODE = 10;
        const int CONTINENT_CODE = 11;
        const int LATITUDE = 12;
        const int LONGITUDE = 13;
        const int DMA = 14;
        const int AREA_CODE = 15;
        const int TIMEZONE = 16;
        const int DATA = 17;

        public void UpdateExcel(Excel.Worksheet sheet, JObject json, int row, int startingCol)
        {
            var vals = json["data"]["geo"];
            sheet.Cells[row, startingCol + HOST].Value = vals["host"].Value<String>();
            sheet.Cells[row, startingCol + IP].Value = vals["ip"].Value<String>();
            sheet.Cells[row, startingCol + RDNS].Value = vals["rdns"].Value<String>();
            sheet.Cells[row, startingCol + ASN].Value = vals["asn"].Value<String>();
            sheet.Cells[row, startingCol + ISP].Value = vals["isp"].Value<String>();
            sheet.Cells[row, startingCol + COUNTRY_NAME].Value = vals["country_name"].Value<String>();
            sheet.Cells[row, startingCol + COUNTRY_CODE].Value = vals["country_code"].Value<String>();
            sheet.Cells[row, startingCol + REGION].Value = vals["region"].Value<String>();
            sheet.Cells[row, startingCol + CITY].Value = vals["city"].Value<String>();
            sheet.Cells[row, startingCol + POSTAL_CODE].Value = vals["postal_code"].Value<String>();
            sheet.Cells[row, startingCol + CONTINENT_CODE].Value = vals["continent_code"].Value<String>();
            sheet.Cells[row, startingCol + LATITUDE].Value = vals["latitude"].Value<String>();
            sheet.Cells[row, startingCol + LONGITUDE].Value = vals["longitude"].Value<String>();
            sheet.Cells[row, startingCol + DMA].Value = vals["dma_code"].Value<String>();
            sheet.Cells[row, startingCol + AREA_CODE].Value = vals["area_code"].Value<String>();
            sheet.Cells[row, startingCol + TIMEZONE].Value = vals["timezone"].Value<String>();
            //sheet.Cells[row, startingCol + DATA].Value = json.ToString();

        }

        public void UpdateExcel2(Excel.Worksheet sheet, string ip, CityResponse city, AsnResponse asn, int row, int startingCol)
        {
            sheet.Cells[row, startingCol + IP].Value = ip;
            if (asn != null)
            {
                sheet.Cells[row, startingCol + ASN].Value = asn.AutonomousSystemNumber ?? 0;
                sheet.Cells[row, startingCol + ISP].Value = asn.AutonomousSystemOrganization;
            }
            sheet.Cells[row, startingCol + COUNTRY_NAME].Value = city.Country.Name;
            sheet.Cells[row, startingCol + COUNTRY_CODE].Value = city.Country.IsoCode;
            sheet.Cells[row, startingCol + REGION].Value = city.Subdivisions.FirstOrDefault()?.Name;
            sheet.Cells[row, startingCol + CITY].Value = city.City.Name;
            sheet.Cells[row, startingCol + POSTAL_CODE].Value = city.Postal.Code;
            sheet.Cells[row, startingCol + CONTINENT_CODE].Value = city.Continent.Code;
            sheet.Cells[row, startingCol + LATITUDE].Value = city.Location.Latitude ?? 0;
            sheet.Cells[row, startingCol + LONGITUDE].Value = city.Location.Longitude ?? 0;
            sheet.Cells[row, startingCol + DMA].Value = city.Location.MetroCode ?? 0;
            //sheet.Cells[row, startingCol + AREA_CODE].Value = vals["area_code"].Value<String>();
            sheet.Cells[row, startingCol + TIMEZONE].Value = city.Location.TimeZone;



        }
    }
}

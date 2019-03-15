using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using HotSpotUserInfo.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = Microsoft.Office.Interop.Excel;


namespace ConsoleHotSpotApp
{
    public static class Methods
    {
        private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        private const string Hapikey = "demo";
        public static List<Contact> A(DateTime modifiedOnOrAfter)
        {
            var startDateInMilliseconds = modifiedOnOrAfter.ToUniversalTime().Subtract(new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;

            var client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var response = client.GetAsync("https://api.hubapi.com/contacts/v1/lists/recently_updated/contacts/recent?hapikey=" + Hapikey).Result;
            var responseString = response.Content.ReadAsStringAsync().Result;

            var contactsList = new List<Contact>();

            var jContacts = JObject.Parse(responseString);
            var isHasMore = (bool)jContacts["has-more"];
            var vidOffset = (int)jContacts["vid-offset"];
            var timeOffset = (double)jContacts["time-offset"];
            var length = ((JArray)jContacts["contacts"]).Count;
            Console.WriteLine("Getting info about contacts by Ids...");

            for (var i = 0; i < length; i++)
            {
                var contact = GetAndFillContact((int)jContacts.SelectToken("contacts[" + i + "].vid"));
                contactsList.Add(contact);
            }
            if (isHasMore)
            {
                Console.WriteLine("Getting info about contacts by Ids...");

                do
                {
                    response = client.GetAsync("https://api.hubapi.com/contacts/v1/lists/recently_updated/contacts/recent?hapikey=" + Hapikey + "&count=100" + "&vidOffset=" + vidOffset + "&timeOffset=" + timeOffset).Result;
                    responseString = response.Content.ReadAsStringAsync().Result;
                    jContacts = JObject.Parse(responseString);
                    vidOffset = (int)jContacts["vid-offset"];
                    timeOffset = (double)jContacts["time-offset"];
                    length = ((JArray)jContacts["contacts"]).Count;
                    for (var i = 0; i < length; i++)
                    {
                        var contact = GetAndFillContact((int)jContacts.SelectToken("contacts[" + i + "].vid"));
                        contactsList.Add(contact);
                    }
                } while (timeOffset >= startDateInMilliseconds);
            }

            DisplayInExcel(contactsList);
            return contactsList;
        }

        public static Contact GetAndFillContact(int vid)
        {
            var contact = new Contact();
            var client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var contactResponse = client.GetAsync("https://api.hubapi.com/contacts/v1/contact/vid/" + vid + "/profile?hapikey=" + Hapikey).Result;
            var contactString = contactResponse.Content.ReadAsStringAsync().Result;


            var jContact = JObject.Parse(contactString);
            var contactProps = jContact.SelectToken("properties");
            contact.Id = (int)jContact.SelectToken("vid");
            contact.FirstName = contactProps["firstname"] != null ? contactProps.SelectToken("firstname.value").ToString() : "no-firstname-data";
            contact.LastName = contactProps["lastname"] != null ? contactProps.SelectToken("lastname.value").ToString() : "no-lastname-data";
            contact.LifeCycleStage = contactProps["lifecyclestage"] != null ? contactProps.SelectToken("lifecyclestage.value").ToString() : "no-lifecyclestage-data";
            Console.WriteLine("Getting info about company by Id...");
            contact.Company = jContact["associated-company"] != null
                ? GetAndFillCompany((int)jContact.SelectToken("associated-company.company-id"))
                : GetAndFillCompany(null);

            return contact;
        }

        public static Company GetAndFillCompany(int? companyId)
        {

            var company = new Company();
            if (companyId != null)
            {
                var client = new HttpClient();

                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                var companyResponse = client.GetAsync("https://api.hubapi.com/companies/v2/companies/" + companyId + "?hapikey=" + Hapikey).Result;
                var companyString = companyResponse.Content.ReadAsStringAsync().Result;

                var jCompany = JObject.Parse(companyString);
                var companyProps = jCompany.SelectToken("properties");
                company.Id = companyId;
                company.Company_Name = companyProps["name"] != null ? jCompany.SelectToken("properties.name.value").ToString() : "no-data";
                company.Company_WebSite = companyProps["website"] != null ? jCompany.SelectToken("properties.website.value").ToString() : "no-data";
                company.Company_ZipCode = companyProps["zip"] != null ? (int)jCompany.SelectToken("properties.zip.value") : 0;

                var jCitySate = GetCityAndStateFromZip(company.Company_ZipCode);
                dynamic data = JObject.Parse(jCitySate);
                company.Company_City = data.city;
                company.Company_State = data.state;
                company.Company_Phone = companyProps["phone"] != null ? jCompany.SelectToken("properties.phone.value").ToString() : "no-data";

            }
            else
            {
                company.Company_Name = "no-data";
                company.Company_WebSite = "no-data";
                company.Company_ZipCode = 0;
                company.Company_City = "no-data";
                company.Company_State = "no-data";
                company.Company_Phone = "no-data";
            }
            return company;
        }

        public static string GetCityAndStateFromZip(int zip)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            var cityAndState = client.GetAsync("http://ZiptasticAPI.com/" + zip).Result;
            var cityAndStateString = cityAndState.Content.ReadAsStringAsync().Result;

            return cityAndStateString;
        }

        public static void DisplayInExcel(List<Contact> contacts)
        {
            var excelApp = new Excel.Application { Visible = true };
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;
            Console.WriteLine("Transfer to Excell...");

            var type = typeof(Contact);
            var properties = type.GetProperties();
            for (var j = 0; j < properties.Length; j++)
            {
                var curLetter = j;
                if (properties[j].Name == "Company")
                {
                    Type subType = typeof(Company);
                    PropertyInfo[] subProperties = subType.GetProperties();
                    var k = curLetter;
                    foreach (var pInfo in subProperties)
                    {
                        if (pInfo.Name == "Id")
                            continue;
                        workSheet.Cells[1, Alphabet[k].ToString()] = pInfo.Name;
                        var rng = (Excel.Range)workSheet.Cells[1, Alphabet[k].ToString()];
                        rng.Font.Bold = true;
                        rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        rng.Borders.Weight = 3d;
                        k++;
                    }
                }
                else
                {
                    workSheet.Cells[1, Alphabet[j].ToString()] = properties[j].Name;
                    var rng = (Excel.Range)workSheet.Cells[1, Alphabet[j].ToString()];
                    rng.Font.Bold = true;
                    rng.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rng.Borders.Weight = 3d;
                }
            }
            var row = 2;
            var column = 0;
            foreach (var contact in contacts)
            {
                workSheet.Cells[row, column + 1].Value = contact.Id;
                workSheet.Cells[row, column + 2].Value = contact.FirstName;
                workSheet.Cells[row, column + 3].Value = contact.LastName;
                workSheet.Cells[row, column + 4].Value = contact.LifeCycleStage;
                workSheet.Cells[row, column + 5].Value = contact.Company.Company_Name;
                workSheet.Cells[row, column + 6].Value = contact.Company.Company_WebSite;
                workSheet.Cells[row, column + 7].Value = contact.Company.Company_City;
                workSheet.Cells[row, column + 8].Value = contact.Company.Company_State;
                workSheet.Cells[row, column + 9].Value = contact.Company.Company_ZipCode;
                workSheet.Cells[row, column + 10].Value = contact.Company.Company_Phone;

                workSheet.Columns.AutoFit();
                column = 0;
                row++;
            }
        }
    }
}

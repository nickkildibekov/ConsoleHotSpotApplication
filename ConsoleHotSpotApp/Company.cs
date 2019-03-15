using Newtonsoft.Json;

namespace HotSpotUserInfo.Models
{
    public class Company
    {
        public int? Id { get; set; }
        public string Company_Name { get; set; }
        public string Company_WebSite { get; set; }
        public string Company_City { get; set; }
        public string Company_State { get; set; }
        public int Company_ZipCode { get; set; }
        public string Company_Phone { get; set; }
    }
}
using Newtonsoft.Json;

namespace HotSpotUserInfo.Models
{
    public class Contact
    {
        public int Id { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string LifeCycleStage { get; set; }
        public Company Company { get; set; }
        public Contact()
        {
            Company = new Company();
        }
    }
}
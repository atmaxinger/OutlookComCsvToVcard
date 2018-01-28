using System.Collections.Generic;

namespace OutlookComToVcard
{
    public class OutlookComCsvContact
    {
        public string FirstName; //OK
        public string MiddleName; //OK
        public string LastName; //OK
        public string Title; //OK
        public string Suffix; //OK
        public string Nickname; //OK
        public string GivenYomi;
        public string SurnameYomi;
        public string EMailAddress; //OK
        public string EMail2Address; //OK
        public string EMail3Address; //OK
        public string HomePhone; //OK
        public string HomePhone2; //OK
        public string BusinessPhone; //OK
        public string BusinessPhone2; //OK
        public string MobilePhone; //OK
        public string CarPhone; //OK
        public string OtherPhone; //OK
        public string PrimaryPhone; //OK
        public string Pager; //OK
        public string BusinessFax;
        public string HomeFax;
        public string OtherFax;
        public string CompanyMainPhone; //OK
        public string Callback;
        public string RadioPhone;
        public string Telex;
        public string TTYTDDPhone;
        public string IMAddress;
        public string JobTitle;
        public string Department; //OK
        public string Company; //OK
        public string OfficeLocation;
        public string ManagersName;
        public string AssistantsName;
        public string AssistantsPhone;
        public string CompanyYomi;
        public string BusinessStreet;
        public string BusinessCity;
        public string BusinessState;
        public string BusinessPostalCode;
        public string BusinessCountryRegion;
        public string HomeStreet;
        public string HomeCity;
        public string HomeState;
        public string HomePostalCode;
        public string HomeCountryRegion;
        public string OtherStreet;
        public string OtherCity;
        public string OtherState;
        public string OtherPostalCode;
        public string OtherCountryRegion;
        public string PersonalWebPage;
        public string Spouse;
        public string Schools;
        public string Hobby;
        public string Location;
        public string WebPage;
        public string Birthday;
        public string Anniversary;
        public string Notes;
        
        public IEnumerable<string> GetEmails()
        {
            List<string> emails = new List<string>();

            if(EMailAddress != null && EMailAddress != "")
            {
                emails.Add(EMailAddress);
            }

            if (EMail2Address != null && EMail2Address != "")
            {
                emails.Add(EMail2Address);
            }

            if (EMail3Address != null && EMail3Address != "")
            {
                emails.Add(EMail3Address);
            }

            return emails;
        }

        public IEnumerable<string> GetHomePhones()
        {
            List<string> phones = new List<string>();

            if(HomePhone != null && HomePhone != "")
            {
                phones.Add(HomePhone);
            }

            if (HomePhone2 != null && HomePhone2 != "")
            {
                phones.Add(HomePhone2);
            }

            return phones;
        }

        public IEnumerable<string> GetBusinessPhones()
        {
            List<string> phones = new List<string>();

            if (BusinessPhone != null && BusinessPhone != "")
            {
                phones.Add(BusinessPhone);
            }

            if (BusinessPhone2 != null && BusinessPhone2 != "")
            {
                phones.Add(BusinessPhone2);
            }

            if(CompanyMainPhone != null && CompanyMainPhone != "")
            {
                phones.Add(CompanyMainPhone);
            }

            return phones;
        }
    }
}

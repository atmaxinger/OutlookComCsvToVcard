using CsvHelper.Configuration;

namespace OutlookComToVcard
{
    public sealed class OutlookComCsvMap : ClassMap<OutlookComCsvContact>
    {
        public OutlookComCsvMap()
        {
            Map(m => m.FirstName).Name("First Name");
            Map(m => m.MiddleName).Name("Middle Name");
            Map(m => m.LastName).Name("Last Name");
            Map(m => m.Title).Name("Title");
            Map(m => m.Suffix).Name("Suffix");
            Map(m => m.Nickname).Name("Nickname");
            Map(m => m.GivenYomi).Name("Given Yomi");
            Map(m => m.SurnameYomi).Name("Surname Yomi");
            Map(m => m.EMailAddress).Name("E-mail Address");
            Map(m => m.EMail2Address).Name("E-mail 2 Address");
            Map(m => m.EMail3Address).Name("E-mail 3 Address");
            Map(m => m.HomePhone).Name("Home Phone");
            Map(m => m.HomePhone2).Name("Home Phone 2");
            Map(m => m.BusinessPhone).Name("Business Phone");
            Map(m => m.BusinessPhone2).Name("Business Phone 2");
            Map(m => m.MobilePhone).Name("Mobile Phone");
            Map(m => m.CarPhone).Name("Car Phone");
            Map(m => m.OtherPhone).Name("Other Phone");
            Map(m => m.PrimaryPhone).Name("Primary Phone");
            Map(m => m.Pager).Name("Pager");
            Map(m => m.BusinessFax).Name("Business Fax");
            Map(m => m.HomeFax).Name("Home Fax");
            Map(m => m.OtherFax).Name("Other Fax");
            Map(m => m.CompanyMainPhone).Name("Company Main Phone");
            Map(m => m.Callback).Name("Callback");
            Map(m => m.RadioPhone).Name("Radio Phone");
            Map(m => m.Telex).Name("Telex");
            Map(m => m.TTYTDDPhone).Name("TTY/TDD Phone");
            Map(m => m.IMAddress).Name("IMAddress");
            Map(m => m.JobTitle).Name("Job Title");
            Map(m => m.Department).Name("Department");
            Map(m => m.Company).Name("Company");
            Map(m => m.OfficeLocation).Name("Office Location");
            Map(m => m.ManagersName).Name("Manager's Name");
            Map(m => m.AssistantsName).Name("Assistant's Name");
            Map(m => m.AssistantsPhone).Name("Assistant's Phone");
            Map(m => m.CompanyYomi).Name("Company Yomi");
            Map(m => m.BusinessStreet).Name("Business Street");
            Map(m => m.BusinessCity).Name("Business City");
            Map(m => m.BusinessState).Name("Business State");
            Map(m => m.BusinessPostalCode).Name("Business Postal Code");
            Map(m => m.BusinessCountryRegion).Name("Business Country/Region");
            Map(m => m.HomeStreet).Name("Home Street");
            Map(m => m.HomeCity).Name("Home City");
            Map(m => m.HomeState).Name("Home State");
            Map(m => m.HomePostalCode).Name("Home Postal Code");
            Map(m => m.HomeCountryRegion).Name("Home Country/Region");
            Map(m => m.OtherStreet).Name("Other Street");
            Map(m => m.OtherCity).Name("Other City");
            Map(m => m.OtherState).Name("Other State");
            Map(m => m.OtherPostalCode).Name("Other Postal Code");
            Map(m => m.OtherCountryRegion).Name("Other Country/Region");
            Map(m => m.PersonalWebPage).Name("Personal Web Page");
            Map(m => m.Spouse).Name("Spouse");
            Map(m => m.Schools).Name("Schools");
            Map(m => m.Hobby).Name("Hobby");
            Map(m => m.Location).Name("Location");
            Map(m => m.WebPage).Name("Web Page");
            Map(m => m.Birthday).Name("Birthday");
            Map(m => m.Anniversary).Name("Anniversary");
            Map(m => m.Notes).Name("Notes");
        }
    }
}

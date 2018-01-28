using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using MixERP.Net.VCards;
using MixERP.Net.VCards.Models;

namespace OutlookComToVcard
{
    class OutlookComToVcardConverter
    {
        /// <summary>
        /// Converts a list of Outlook.com exported contacts to VCard contacts
        /// </summary>
        /// <param name="csvContacts">The outlook contacts</param>
        /// <param name="correctDates">Add +1 day to birthday an anniversary dates</param>
        /// <returns></returns>
        public IEnumerable<VCard> ConvertOutlookComCsvContacts(IEnumerable<OutlookComCsvContact> csvContacts, bool correctDates)
        {
            List<VCard> vcards = new List<VCard>();

            foreach (var contact in csvContacts)
            {
                vcards.Add(ConvertOutlookComCsvContact(contact, correctDates));
            }

            return vcards;
        }

        /// <summary>
        /// Convert a Outlook.com exported contact to VCard
        /// </summary>
        /// <param name="csvContact">The outlook contact</param>
        /// <param name="correctDates">Add +1 day to birthday an anniversary dates</param>
        /// <returns></returns>
        public VCard ConvertOutlookComCsvContact(OutlookComCsvContact csvContact, bool correctDates)
        {
            VCard vcard = new VCard
            {
                Version = MixERP.Net.VCards.Types.VCardVersion.V4,

                FirstName = csvContact.FirstName,
                MiddleName = csvContact.MiddleName,
                LastName = csvContact.LastName,
                Title = csvContact.Title,
                Suffix = csvContact.Suffix,
                NickName = csvContact.Nickname,

                Organization = csvContact.Company,
                OrganizationalUnit = csvContact.Department,

                Emails = getEmails(csvContact),
                Telephones = ConvertTelephones(csvContact),
                Addresses = ConvertAddresses(csvContact)
            };

            if (csvContact.Birthday != "")
            {
                vcard.BirthDay = ConvertDate(csvContact.Birthday, correctDates);
            }

            if (csvContact.Anniversary != "")
            {
                vcard.Anniversary = ConvertDate(csvContact.Anniversary, correctDates);
            }

            return vcard;
        }

        private IEnumerable<Email> getEmails(OutlookComCsvContact csvContact)
        {
            List<Email> emails = new List<Email>();

            foreach (string e in csvContact.GetEmails())
            {
                Email email = new Email();
                email.EmailAddress = e;
                emails.Add(email);
            }

            return emails;
        }

        private IEnumerable<Telephone> ConvertTelephones(OutlookComCsvContact csvContact)
        {
            List<Telephone> telephones = new List<Telephone>();

            foreach (string p in csvContact.GetHomePhones())
            {
                Telephone telephone = new Telephone();
                telephone.Number = p;
                telephone.Type = MixERP.Net.VCards.Types.TelephoneType.Home;
                telephones.Add(telephone);
            }

            foreach (string p in csvContact.GetBusinessPhones())
            {
                Telephone telephone = new Telephone();
                telephone.Number = p;
                telephone.Type = MixERP.Net.VCards.Types.TelephoneType.Work;
                telephones.Add(telephone);
            }

            if (csvContact.MobilePhone != null && csvContact.MobilePhone != "")
            {
                Telephone mobile = new Telephone();
                mobile.Number = csvContact.MobilePhone;
                mobile.Type = MixERP.Net.VCards.Types.TelephoneType.Cell;
                telephones.Add(mobile);
            }

            if (csvContact.CarPhone != null && csvContact.CarPhone != "")
            {
                Telephone car = new Telephone();
                car.Number = csvContact.CarPhone;
                car.Type = MixERP.Net.VCards.Types.TelephoneType.Car;
                telephones.Add(car);
            }

            if (csvContact.OtherPhone != null && csvContact.OtherPhone != "")
            {
                Telephone other = new Telephone();
                other.Number = csvContact.OtherPhone;
                // TODO: Find a better type for "other"
                other.Type = MixERP.Net.VCards.Types.TelephoneType.Personal;
                telephones.Add(other);
            }

            if (csvContact.PrimaryPhone != null && csvContact.PrimaryPhone != "")
            {
                Telephone primary = new Telephone();
                primary.Number = csvContact.OtherPhone;
                primary.Type = MixERP.Net.VCards.Types.TelephoneType.Preferred;
                telephones.Add(primary);
            }

            if (csvContact.Pager != null && csvContact.Pager != "")
            {
                Telephone pager = new Telephone();
                pager.Number = csvContact.OtherPhone;
                pager.Type = MixERP.Net.VCards.Types.TelephoneType.Pager;
                telephones.Add(pager);
            }


            if (csvContact.HomeFax != null && csvContact.HomeFax != "")
            {
                Telephone fax = new Telephone();
                fax.Number = csvContact.HomeFax;
                fax.Type = MixERP.Net.VCards.Types.TelephoneType.Fax;
                telephones.Add(fax);
            }
            if (csvContact.BusinessFax != null && csvContact.BusinessFax != "")
            {
                Telephone fax = new Telephone();
                fax.Number = csvContact.BusinessFax;
                fax.Type = MixERP.Net.VCards.Types.TelephoneType.Fax;
                telephones.Add(fax);
            }
            if (csvContact.OtherFax != null && csvContact.OtherFax != "")
            {
                Telephone fax = new Telephone();
                fax.Number = csvContact.OtherFax;
                fax.Type = MixERP.Net.VCards.Types.TelephoneType.Fax;
                telephones.Add(fax);
            }

            return telephones;
        }

        private Address ToAddress(String Country, String PostalCode, String Locality, String Street)
        {
            Address address = new Address();
            address.Country = Country;
            address.PostalCode = PostalCode;
            address.Locality = Locality;
            address.Street = Street;
            return address;
        }

        private IEnumerable<Address> ConvertAddresses(OutlookComCsvContact csvContact)
        {
            List<Address> addresses = new List<Address>();

            addresses.Add(ToAddress(csvContact.HomeCountryRegion, csvContact.HomePostalCode, csvContact.HomeCity, csvContact.HomeStreet));
            addresses.Add(ToAddress(csvContact.BusinessCountryRegion, csvContact.BusinessPostalCode, csvContact.BusinessCity, csvContact.BusinessStreet));
            addresses.Add(ToAddress(csvContact.OtherCountryRegion, csvContact.OtherPostalCode, csvContact.OtherCity, csvContact.OtherStreet));

            return addresses;
        }

        private DateTime ConvertDate(String date, bool correctDates)
        {
            // date is stored as "DD.MM.YYYY HH:mm:ss"

            string[] splitDateAndTime = date.Split(' ');
            string[] splitDate = splitDateAndTime[0].Split('.');

            int day = Int16.Parse(splitDate[0]);
            int month = Int16.Parse(splitDate[1]);
            int year = Int32.Parse(splitDate[2]);

            DateTime dt = new DateTime(year, month, day);

            if (correctDates)
            {
                dt = dt.AddDays(1);
            }

            return dt;
        }
    }
}

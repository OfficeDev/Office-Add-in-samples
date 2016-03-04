using DXDemos.Office365.Models;
using DXDemos.Office365.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace DXDemos.Office365.Controllers
{
    public class ContactController : ApiController
    {
        // GET: api/Contact
        public IEnumerable<ContactModel> Get()
        {
            return new List<ContactModel>();
        }

        // GET: api/Contact?id=ridize@rzna.onmicrosoft.com
        public ContactModel Get(string id)
        {
            //get the contact from DocumentDB
            var contact = DocumentDBRepository<ContactModel>.GetItem("Contacts", i => i.Id == id);
            if (contact != null)
            {
                //populate associates
                contact.Associates = DocumentDBRepository<ContactModel>.GetItems("Contacts", i => i.Domain == contact.Domain && i.Id != id).ToList();
            }
            return contact;
        }

        // POST: api/Contact
        public void Post([FromBody]ContactModel value)
        {
            //get details off header
            var callerName = this.Request.Headers.GetValues("callerName").FirstOrDefault();
            var callerEmail = this.Request.Headers.GetValues("callerEmail").FirstOrDefault();

            value.Domain = value.Id.Split('@')[1];

            //HACK: add some random notes for this new contact
            value.Notes = new List<NoteModel>();
            value.Notes.Add(new NoteModel() { Id = Guid.NewGuid(), Note = "Contact created " + DateTime.Now.ToLongDateString(), AuthorName = callerName, AuthorEmail = callerEmail, PostDate = DateTime.Now.ToLongDateString(), Email = value.Id });

            //HACK: add some random invoices for this new contact
            value.Invoices = new List<InvoiceModel>();
            value.Invoices.Add(new InvoiceModel() { Id = Guid.NewGuid(), InvoiceNumber = "220704", InvoiceDate = "March 18, 2015", Amount = 1727.40M, Status = "Invoiced" });
            value.Invoices.Add(new InvoiceModel() { Id = Guid.NewGuid(), InvoiceNumber = "220703", InvoiceDate = "February 9, 2015", Amount = 866.92M, Status = "Overdue" });
            value.Invoices.Add(new InvoiceModel() { Id = Guid.NewGuid(), InvoiceNumber = "220702", InvoiceDate = "January 16, 2015", Amount = 2348.07M, Status = "Paid" });
            value.Invoices.Add(new InvoiceModel() { Id = Guid.NewGuid(), InvoiceNumber = "220701", InvoiceDate = "December 3, 2014", Amount = 1986.13M, Status = "Paid" });

            //initialize associates
            value.Associates = new List<ContactModel>();

            //save to DocumentDB
            DocumentDBRepository<ContactModel>.CreateItemAsync("Contacts", value);
        }
    }
}

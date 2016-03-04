using DXDemos.Office365.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace DXDemos.Office365.Controllers
{
    public class NoteController : ApiController
    {
        // POST: api/Note
        public void Post([FromBody]NoteModel value)
        {
            //get the contact based on the email
            var contact = DocumentDBRepository<ContactModel>.GetItem("Contacts", i => i.Id == value.Email);
            if (contact != null)
            {
                value.Id = Guid.NewGuid();
                contact.Notes.Insert(0, value);
                DocumentDBRepository<ContactModel>.UpdateItemAsync("Contacts", contact.Id, contact);
            }
        }
    }
}

using System;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Http;
using System.Xml;

namespace HelloWorldWeb.Controllers
{
    public class PhotoController : ApiController
    {
        public string Get()
        {
            string url = "http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1";

            // Create the request.
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            WebResponse response = request.GetResponse();

            using (Stream responseStream = response.GetResponseStream())
            {
                // Process the result.
                StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                string result = reader.ReadToEnd();

                // Parse the XML response and get the URL.
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(result);
                string photoURL = "http://bing.com" + doc.SelectSingleNode("/images/image/url").InnerText;

                // Fetch the photo and return it as a Base64-encoded string.
                return getPhotoFromURL(photoURL);
            }
        }

        private string getPhotoFromURL(string imageURL)
        {
            var webClient = new WebClient();
            byte[] imageBytes = webClient.DownloadData(imageURL);
            return Convert.ToBase64String(imageBytes);
        }
    }
}
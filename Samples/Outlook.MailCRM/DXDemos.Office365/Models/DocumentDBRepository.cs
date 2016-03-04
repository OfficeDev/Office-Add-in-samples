using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.Documents;
using Microsoft.Azure.Documents.Client;
using Microsoft.Azure.Documents.Linq;
using System.Configuration;
using System.Linq.Expressions;

namespace DXDemos.Office365.Models
{
    public static class DocumentDBRepository<T>
    {
       private static string databaseId;
       private static String DatabaseId
       {
            get
            {
                if (string.IsNullOrEmpty(databaseId))
                {
                    databaseId = ConfigurationManager.AppSettings["ddb:database"];
                }

                return databaseId;
            }
        }

        private static string collectionId;
        private static Database database;
        private static Database Database
        {
            get
            {
                if (database == null)
                {
                    database = ReadOrCreateDatabase();
                }

                return database;
            }
        }

        private static DocumentCollection collection;
        private static DocumentCollection Collection
        {
            get
            {
                if (collection == null)
                {
                    collection = ReadOrCreateCollection(Database.SelfLink);
                }

                return collection;
            }
        }

        private static DocumentClient client;
        private static DocumentClient Client
        {
            get
            {
                if (client == null)
                {
                    string endpoint = ConfigurationManager.AppSettings["ddb:endpoint"];
                    string authKey = ConfigurationManager.AppSettings["ddb:authKey"];
                    Uri endpointUri = new Uri(endpoint);
                    client = new DocumentClient(endpointUri, authKey);
                }

                return client;
            }
        }

        private static DocumentCollection ReadOrCreateCollection(string databaseLink)
        {
            var col = Client.CreateDocumentCollectionQuery(databaseLink)
                              .Where(c => c.Id == collectionId)
                              .AsEnumerable()
                              .FirstOrDefault();

            if (col == null)
            {
                col = Client.CreateDocumentCollectionAsync(databaseLink, new DocumentCollection { Id = collectionId }).Result;
            }

            return col;
        }

        private static Database ReadOrCreateDatabase()
        {
            var db = Client.CreateDatabaseQuery()
                            .Where(d => d.Id == DatabaseId)
                            .AsEnumerable()
                            .FirstOrDefault();

            if (db == null)
            {
                db = Client.CreateDatabaseAsync(new Database { Id = DatabaseId }).Result;
            }

            return db;
        }

        public static IEnumerable<T> GetItems(string collectionToQuery, Expression<Func<T, bool>> predicate)
        {
            collectionId = collectionToQuery;
            return Client.CreateDocumentQuery<T>(Collection.DocumentsLink)
                .Where(predicate)
                .AsEnumerable();
        }

        public static T GetItem(string collectionToQuery, Expression<Func<T, bool>> predicate)
        {
            collectionId = collectionToQuery;
            return Client.CreateDocumentQuery<T>(Collection.DocumentsLink)
                        .Where(predicate)
                        .AsEnumerable()
                        .FirstOrDefault();
        }

        public static Document GetDocument(string collectionToQuery, string id)
        {
            collectionId = collectionToQuery;
            return Client.CreateDocumentQuery(Collection.DocumentsLink)
                .Where(d => d.Id == id)
                .AsEnumerable()
                .FirstOrDefault();
        }

        public static async Task<Document> UpdateItemAsync(string collectionToQuery, string id, T item)
        {
            Document doc = GetDocument(collectionToQuery, id);
            return await Client.ReplaceDocumentAsync(doc.SelfLink, item);
        }

        public static async Task<Document> CreateItemAsync(string collectionToQuery, T item)
        {
            collectionId = collectionToQuery;
            return await Client.CreateDocumentAsync(Collection.SelfLink, item);
        }
    }
}

// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace OfficeAddinMicrosoftGraphASPNET.Helpers
{
    /// <summary>
    /// Provides methods for accessing OData endpoints.
    /// </summary>
    internal static class ODataHelper
    {
        /// <summary>
        /// Gets any JSON array from any OData endpoint that requires an access token.
        /// </summary>
        /// <typeparam name="T">The .NET type to which the members of the array will be converted.</typeparam>
        /// <param name="itemsUrl">The URL of the OData endpoint.</param>
        /// <param name="accessToken">An OAuth access token.</param>
        /// <returns>Collection of T items that the caller can cast to any IEnumerable type.</returns>
        internal static async Task<IEnumerable<T>> GetItems<T>(string itemsUrl, string accessToken)
        {
            dynamic jsonData = await SendRequestWithAccessToken(itemsUrl, accessToken);

            // Convert to .NET class and populate the properties of the model objects,
            // and then populate the IEnumerable object and return it.
            JArray jsonArray = jsonData.value;
            return JsonConvert.DeserializeObject<IEnumerable<T>>(jsonArray.ToString());
        }

        /// <summary>
        /// Sends a request to the specified OData URL with the specified access token.
        /// </summary>
        /// <param name="itemsUrl">The OData endpoint URL.</param>
        /// <param name="accessToken">The access token for the endpoint resource.</param>
        /// <returns></returns>
        internal static async Task<dynamic> SendRequestWithAccessToken(string itemsUrl, string accessToken)
        {
            dynamic jsonData = null;

            using (var client = new HttpClient())
            {
                // Create and send the HTTP Request
                using (var request = new HttpRequestMessage(HttpMethod.Get, itemsUrl))
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                    using (HttpResponseMessage response = await client.SendAsync(request))
                    {
                        if (response.IsSuccessStatusCode)
                        {
                            HttpContent content = response.Content;
                            string responseContent = await content.ReadAsStringAsync();

                            jsonData = JsonConvert.DeserializeObject(responseContent);
                        }
                    }
                }
            }
            return jsonData;
        }
    }
}


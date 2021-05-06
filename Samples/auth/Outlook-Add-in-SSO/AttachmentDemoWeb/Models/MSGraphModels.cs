// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the root of the project.

/* 
    This file provides .NET model classes for Microsoft Graph data. 
*/

using Newtonsoft.Json;

namespace AttachmentDemoWeb.Models
{
    /// <summary>
    /// Models Microsoft Graph entities.
    /// </summary>
    /// <remark>
    /// There are many properties that almost all Microsoft Graph objects have. 
    /// To avoid redundancy in the class definitions, use this abstract class.
    ///</remark>
    public abstract class MSGraphObject
    {
        public string Name { get; set; }
    }

    /// <summary>
    /// Objects that have an Etag.
    /// </summary>
    public interface IEtagable
    {
        // When the JSON property name begins with a character that cannot
        // begin a .NET property name, a JsonProperty attribute maps the names.
        [JsonProperty("@odata.etag")]
        string Etag { get; set; }
    }

    /// <summary>
    /// A OneDriveItem can be a file or folder.
    /// </summary>
    public class OneDriveItem : MSGraphObject, IEtagable
    {
        public string Etag { get; set; }
    }
}
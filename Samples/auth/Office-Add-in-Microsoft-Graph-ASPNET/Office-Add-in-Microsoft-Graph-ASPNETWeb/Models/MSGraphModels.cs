// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OfficeAddinMicrosoftGraphASPNET.Models
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
        public string Id { get; set; }
    }

    public class ExcelWorkbook : MSGraphObject
    {
        // This class uses only inherited properties.
    }
}

// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

namespace OfficeAddinSSOWeb.Models
{
    public class ErrorViewModel
    {
        public string? RequestId { get; set; }

        public bool ShowRequestId => !string.IsNullOrEmpty(RequestId);
    }
}
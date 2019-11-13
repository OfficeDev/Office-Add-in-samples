// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharedModule
{
	class HtmlHelpHelper
	{

		[DllImport("hhctrl.ocx", CharSet = CharSet.Auto, SetLastError=true)]
		public static extern int HtmlHelp(IntPtr hwndCaller, [MarshalAs(UnmanagedType.LPTStr)]string pszFile, int uCommand, int dwData);
	}
}

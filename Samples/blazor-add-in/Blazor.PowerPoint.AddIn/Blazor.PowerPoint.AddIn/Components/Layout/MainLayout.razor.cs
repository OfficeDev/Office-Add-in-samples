/* Copyright(c) Maarten van Stam. All rights reserved. Licensed under the MIT License. */
namespace Blazor.PowerPoint.AddIn.Components.Layout
{
    public partial class MainLayout
    {
        private string TrademarkMessage1 { get; set; } = "Copyright © " + @DateTime.Now.Year + " Maarten van Stam.";
        private string TrademarkMessage2 { get; set; } = "All rights reserved.";

        private string FrameworkDescription { get; set; } = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription;
    }
}
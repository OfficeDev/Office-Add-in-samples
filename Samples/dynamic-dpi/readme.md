<p><b>MFC ActiveX Not DPI Aware </b> - ActiveX Control created from the MFC template that does use the DPI of the window to determine how to scale its contents.</p>
<p><b>MFCApplicationDPIAware</b> – MFC App that host 2 ActiveX controls and it’s DDPI Aware</p>
<p><b>VSTO ExcelAddin1</b> - VSTO Add-in for Excel.  Open using VSTOSharedAddin solution.</p>
<p><b>VSTO WordAddin1</b> - VSTO Add-in for Word.  Open using VSTOSharedAddin solution.</p>
<p><b>VSTOSharedAddin</b> - Shared Winform and Windows API calls for VSTO Addins.</p>
<p><b>Window Based ActiveX</b> – MFC Window based ActiveX that supports DDPI on WM_SIZE event at this moment (tested with App and Office - works fine)</p>
<p><b>Windowless ActiveX</b> – MFC Windowless ActiveX that supports DDPI on WM_PAINT (it gets HWND of host window from HDC and after it cat get DPI form it, works well in Office and our test App – just requiring that control should be invalidated while WM_DPICHANGE event occurs)</p>

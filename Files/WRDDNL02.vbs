Set objNetwork = CreateObject("WScript.Network")
objNetwork.AddWindowsPrinterConnection "\\opv-prtral01\WRIS_CLJ", "WRIS_CLJ."
objNetwork.AddWindowsPrinterConnection "\\opv-prtral01\WRIS_KM1", "WRIS_KM1."
objNetwork.SetDefaultPrinter "\\opv-prtral01\WRIS_KM1"

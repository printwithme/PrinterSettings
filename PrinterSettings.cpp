#include <Windows.h>
#include <iostream>
#include <stdlib.h>
#include <string>

#include "atlbase.h"
#include "atlstr.h"
#include "comutil.h"


enum class PrinterResult
{
    OK = 0,
    PRINTER_NOT_FOUND = 10,
    INVALID_ARGUMENT,
    DUPLEX_NOT_SUPPORTED,
    COLOR_NOT_SUPPORTED,
    ORIENTATION_NOT_SUPPORTED,
    OTHER_ERROR,
};


PrinterResult ChangePrinterSettings(LPTSTR pPrinterName, short orientation, short color, short duplex)
{
    HANDLE hPrinter = NULL;
    DWORD dwNeeded = 0;
    PRINTER_INFO_2* pi2 = NULL;
    DEVMODE* pDevMode = NULL;
    PRINTER_DEFAULTS pd;
    BOOL bFlag;
    LONG lFlag;

    // Open printer handle (on Windows NT, you need full-access because you
    // will eventually use SetPrinter)
    ZeroMemory(&pd, sizeof(pd));
    pd.DesiredAccess = PRINTER_ALL_ACCESS;
    bFlag = OpenPrinter(pPrinterName, &hPrinter, &pd);
    if (!bFlag || (hPrinter == NULL)) {
        return PrinterResult::PRINTER_NOT_FOUND;
    }

    // The first GetPrinter tells you how big the buffer should be in
    // order to hold all of PRINTER_INFO_2. Note that this should fail with
    // ERROR_INSUFFICIENT_BUFFER.  If GetPrinter fails for any other reason
    // or dwNeeded isn't set for some reason, then there is a problem
    SetLastError(0);
    bFlag = GetPrinter(hPrinter, 2, 0, 0, &dwNeeded);
    if ((!bFlag) && (GetLastError() != ERROR_INSUFFICIENT_BUFFER) || (dwNeeded == 0))
    {
        ClosePrinter(hPrinter);
        return PrinterResult::OTHER_ERROR;
    }

    // Allocate enough space for PRINTER_INFO_2
    pi2 = (PRINTER_INFO_2*)GlobalAlloc(GPTR, dwNeeded);
    if (pi2 == NULL)
    {
        ClosePrinter(hPrinter);
        return PrinterResult::OTHER_ERROR;
    }

    // The second GetPrinter fills in all the current settings, so all you
    // need to do is modify what you're interested in
    bFlag = GetPrinter(hPrinter, 2, (LPBYTE)pi2, dwNeeded, &dwNeeded);
    if (!bFlag)
    {
        GlobalFree(pi2);
        ClosePrinter(hPrinter);
        return PrinterResult::OTHER_ERROR;
    }

    // If GetPrinter didn't fill in the DEVMODE, try to get it by calling
    // DocumentProperties
    if (pi2->pDevMode == NULL)
    {
        dwNeeded = DocumentProperties(NULL, hPrinter,
            pPrinterName,
            NULL, NULL, 0);
        if (dwNeeded <= 0)
        {
            GlobalFree(pi2);
            ClosePrinter(hPrinter);
            return PrinterResult::OTHER_ERROR;
        }

        pDevMode = (DEVMODE*)GlobalAlloc(GPTR, dwNeeded);
        if (pDevMode == NULL)
        {
            GlobalFree(pi2);
            ClosePrinter(hPrinter);
            return PrinterResult::OTHER_ERROR;
        }

        lFlag = DocumentProperties(NULL, hPrinter,
            pPrinterName,
            pDevMode, NULL,
            DM_OUT_BUFFER);
        if (lFlag != IDOK || pDevMode == NULL)
        {
            GlobalFree(pDevMode);
            GlobalFree(pi2);
            ClosePrinter(hPrinter);
            return PrinterResult::OTHER_ERROR;
        }
        pi2->pDevMode = pDevMode;
    }

    // Driver is reporting that it doesn't support this change
    if (!(pi2->pDevMode->dmFields & DM_ORIENTATION))
    {
        GlobalFree(pi2);
        ClosePrinter(hPrinter);
        if (pDevMode)
            GlobalFree(pDevMode);
        return PrinterResult::ORIENTATION_NOT_SUPPORTED;
    }

    // Driver is reporting that it doesn't support this change
    if (!(pi2->pDevMode->dmFields & DM_COLOR))
    {
        GlobalFree(pi2);
        ClosePrinter(hPrinter);
        if (pDevMode)
            GlobalFree(pDevMode);
        return PrinterResult::COLOR_NOT_SUPPORTED;
    }

    // Driver is reporting that it doesn't support this change
    if (!(pi2->pDevMode->dmFields & DM_DUPLEX))
    {
        GlobalFree(pi2);
        ClosePrinter(hPrinter);
        if (pDevMode)
            GlobalFree(pDevMode);
        return PrinterResult::DUPLEX_NOT_SUPPORTED;
    }

    // Specify exactly what we are attempting to change
    pi2->pDevMode->dmFields = DM_ORIENTATION | DM_COLOR | DM_DUPLEX;
    pi2->pDevMode->dmOrientation = orientation;
    pi2->pDevMode->dmColor = color;
    pi2->pDevMode->dmDuplex = duplex;

    // Do not attempt to set security descriptor
    pi2->pSecurityDescriptor = NULL;

    // Make sure the driver-dependent part of devmode is updated
    lFlag = DocumentProperties(NULL, hPrinter,
        pPrinterName,
        pi2->pDevMode, pi2->pDevMode,
        DM_IN_BUFFER | DM_OUT_BUFFER);
    if (lFlag != IDOK)
    {
        GlobalFree(pi2);
        ClosePrinter(hPrinter);
        if (pDevMode)
            GlobalFree(pDevMode);
        return PrinterResult::OTHER_ERROR;
    }

    // Update printer information
    bFlag = SetPrinter(hPrinter, 2, (LPBYTE)pi2, 0);
    if (!bFlag)
        // The driver doesn't support, or it is unable to make the change
    {
        GlobalFree(pi2);
        ClosePrinter(hPrinter);
        if (pDevMode)
            GlobalFree(pDevMode);
        return PrinterResult::OTHER_ERROR;
    }

    // Tell other apps that there was a change
    SendMessageTimeout(HWND_BROADCAST, WM_DEVMODECHANGE, 0L,
        (LPARAM)(LPCSTR)pPrinterName,
        SMTO_NORMAL, 1000, NULL);

    // Clean up
    if (pi2)
        GlobalFree(pi2);
    if (hPrinter)
        ClosePrinter(hPrinter);
    if (pDevMode)
        GlobalFree(pDevMode);
    return PrinterResult::OK;
}

wchar_t* Convert(char* s)
{
    // newsize describes the length of the
    // wchar_t string called wcstring in terms of the number
    // of wide characters, not the number of bytes.
    size_t newsize = strlen(s) + 1;

    // The following creates a buffer large enough to contain
    // the exact number of characters in the original string
    // in the new format. If you want to add more characters
    // to the end of the string, increase the value of newsize
    // to increase the size of the buffer.
    wchar_t* wcstring = new wchar_t[newsize];

    // Convert char* string to a wchar_t* string.
    size_t convertedChars = 0;
    mbstowcs_s(&convertedChars, wcstring, newsize, s, _TRUNCATE);

    return wcstring;

}

void usage(char* argv[])
{
    std::cout << argv[0] << " printer orientation color duplex copies collate" << std::endl;
}


short OrientationOption(std::string arg)
{
    if (arg == "portrait")
    {
        return DMORIENT_PORTRAIT;
    }
    else if (arg == "landscape")
    {
        return DMORIENT_LANDSCAPE;
    }
    else
    {
        throw new std::invalid_argument("Invalid orientation option " + arg);
    }
}

short ColorOption(std::string arg)
{
    if (arg == "color")
    {
        return DMCOLOR_COLOR;
    }
    else if (arg == "monochrome")
    {
        return DMCOLOR_MONOCHROME;
    }
    else
    {
        throw new std::invalid_argument("Invalid color option " + arg);
    }
}

short DuplexOption(std::string arg)
{
    if (arg == "vertical")
    {
        return DMDUP_VERTICAL;
    }
    else if (arg == "horizontal")
    {
        return DMDUP_HORIZONTAL;
    }
    else if (arg == "simplex")
    {
        return DMDUP_SIMPLEX;
    }
    else
    {
        throw new std::invalid_argument("Invalid duplex option " + arg);
    }
}

int main(int argc, char* argv[])
{
    if (argc != 7)
    {
        usage(argv);
        return EXIT_FAILURE;
    }

    char* printerName = argv[1];
    short orientation = OrientationOption(argv[2]);
    short color = ColorOption(argv[3]);
    short duplex = DuplexOption(argv[4]);

    TCHAR *printerNameT;
    printerNameT = Convert(printerName);

    std::cout << "Setting " << printerName << " " << color << " " << duplex  << std::endl;

    PrinterResult result = ChangePrinterSettings(printerNameT, orientation, color, duplex);

    if (result != PrinterResult::OK)
    {
        int code = static_cast<int>(result);
        std::cerr << printerName << " - " << "failed to change printer settings " << "(" << code << ")" << std::endl;
        return code;
    }
    else
    {
        std::cout << printerName << " - " << "settings changed" << std::endl;
        return EXIT_SUCCESS;
    }
}

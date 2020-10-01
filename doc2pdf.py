import os, re, time, datetime, win32com.client

def print_to_Bullzip(file):
    util = win32com.client.Dispatch("Bullzip.PDFUtil")
    settings = win32com.client.Dispatch("Bullzip.PDFSettings")
    settings.PrinterName = util.DefaultPrinterName      # make sure we're controlling the right PDF printer

    outputFile = re.sub("\.[^.]+$", ".pdf", file)
    statusFile = re.sub("\.[^.]+$", ".status", file)

    settings.SetValue("Output", outputFile)
    settings.SetValue("ConfirmOverwrite", "no")
    settings.SetValue("ShowSaveAS", "never")
    settings.SetValue("ShowSettings", "never")
    settings.SetValue("ShowPDF", "no")
    settings.SetValue("ShowProgress", "no")
    settings.SetValue("ShowProgressFinished", "no")     # disable balloon tip
    settings.SetValue("StatusFile", statusFile)         # created after print job
    settings.WriteSettings(True)                        # write settings to the runonce.ini
    util.PrintFile(file, util.DefaultPrinterName)       # send to Bullzip virtual printer
  # wait until print job completes before continuing
    # otherwise settings for the next job may not be used
    timestamp = datetime.datetime.now()
    while( (datetime.datetime.now() - timestamp).seconds < 10):
        if os.path.exists(statusFile) and os.path.isfile(statusFile):
            error = util.ReadIniString(statusFile, "Status", "Errors", '')
            if error != "0":
                raise IOError("PDF was created with errors")
            os.remove(statusFile)
            return
        time.sleep(0.1)
    raise IOError("PDF creation timed out")   
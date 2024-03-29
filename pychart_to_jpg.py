from win32com.client import Dispatch
import os
import pythoncom
class Pyxlchart(object):
    """
    This class exports charts in an Excel Spreadsheet to the FileSystem
    win32com libraries are required.
    """
    def __init__(self):
        pythoncom.CoInitialize()
        self.WorkbookDirectory = ''
        self.WorkbookFilename = ''
        self.GetAllWorkbooks = False
        self.SheetName = ''
        self.ChartName = ''
        self.GetAllWorkbookCharts = False
        self.GetAllWorksheetCharts = False
        self.ExportPath = ''
        self.ImageFilename = ''
        self.ReplaceWhiteSpaceChar = '_'
        self.ImageType = 'jpg'

    def __del__(self):
        pass

    def start_export(self):
        if self.WorkbookDirectory == '':
            return "WorkbookDirectory not set"
        else:
            self._export()

    def _export(self):
        """
        Exports Charts as determined by the settings in class variabels.
        """
        excel = Dispatch("excel.application")
        excel.Visible = False
        wb = excel.Workbooks.Open(os.path.join(self.WorkbookDirectory ,self.WorkbookFilename))
        self._get_Charts_In_Worksheet(wb,self.SheetName,self.ChartName)
        wb.Close(False)
        excel.Quit()

    def _get_Charts_In_Worksheet(self,wb,worksheet = "", chartname = ""):
        if worksheet != "" and chartname != "":
            sht = self._change_sheet(wb,worksheet)
            cht = sht.ChartObjects(chartname)
            self._save_chart(cht, sht.Name)
            return
        if worksheet == "":
            for sht in wb.Worksheets:
                for cht in sht.ChartObjects():
                    if chartname == "":
                        self._save_chart(cht, sht.Name)
                    else:
                        if chartname == cht.Name:
                            self._save_chart(cht, sht.Name)
        else:
            sht = wb.Worksheets(worksheet)
            for cht in sht.ChartObjects():
                if chartname == "":
                    self._save_chart(cht, sht.Name)
                else:
                    if chartname == cht.Name:
                        self._save_chart(cht, sht.Name)

    def _change_sheet(self,wb,worksheet):
        try:
            return wb.Worksheets(worksheet)
        except:
            raise NameError('Unable to Select Sheet: ' + worksheet + ' in Workbook: ' + wb.Name)

    def _save_chart(self,chartObject, sheetnamme):
        imagename = self._get_filename(chartObject.Name, sheetnamme)
        savepath = os.path.join(self.ExportPath,imagename)
        print(savepath)
        chartObject.Chart.Export(savepath,self.ImageType)

    def _get_filename(self, chartname, sheetnamme):
        """
        Replaces white space in self.WorkbookFileName with the value given in self.ReplaceWhiteSpaceChar
        If self.ReplaceWhiteSpaceChar is an empty string then self.WorkBookFileName is left as is
        """
        if self.ImageFilename == '':
            self.ImageFilename == chartname
        if self.ReplaceWhiteSpaceChar != '':
            chartname.replace(' ',self.ReplaceWhiteSpaceChar)
        if self.ImageFilename != "":
            return self.ImageFilename + "_" + sheetnamme + "_" + chartname + "." + self.ImageType
        else:
            return sheetnamme + "_" + chartname + '.' + self.ImageType

def main():
    xl = Pyxlchart()
    xl.WorkbookDirectory = "G:\\"
    xl.WorkbookFilename = "chart.xlsx"
    xl.SheetName = ""
    xl.ImageFilename = "Movies"
    xl.ExportPath = "G:\\"
    xl.ChartName = ""
    xl.start_export()

if __name__ == "__main__":
    main()

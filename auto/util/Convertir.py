import  jpype     
import  asposecells     
jpype.startJVM()

from asposecells.api import Workbook

def ConvToXLSX(dir, name):

    workbook = Workbook(dir)
    workbook.save(name)

def apagarJVM():
    
    jpype.shutdownJVM()
    
import xlwings as xw
import numpy as np
from numpy.ma.core import transpose

from model.Grafica_vogel import grafica_vogel
from model.Vogel_saturado import Vogel_saturado


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]

    Pwf = np.array(sheet['B8:B13'].value)                          #np.array([2400,2000,1500,1000,500,0])
    Pr = sheet['C4'].value
    Qmax = sheet['C5'].value
    sheet["B7"].value= 'Pwf (psi)'
    sheet["B8"].options(transpose = True).value = Pwf
    sheet["B4"].value = 'Pr'
    sheet["C4"].value = Pr
    sheet["B5"].value = 'Qmax'
    sheet["C5"].value = Qmax
    sheet["C7"].value = 'Q (bbl/d)'
    Qo = Vogel_saturado(Pr, Qmax, Pwf)
    Qo = np.round(Qo, 3)
    sheet["C8"].options(transpose=True).value = Qo
    grafica = grafica_vogel(Pwf, Qo)
    sheet.pictures.add(grafica, name = 'Grafica IPR', update = True, left = sheet.range('G3').left, top = sheet.range('G3').top)

@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("tarea.xlsm").set_mock_caller()
    main()

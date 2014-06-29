#!C:\Python27
# -*- coding: utf-8 -*-
import sys, os, datetime

from beniShortFunc import get_unicode
from beniExcel import dict_list_to_excel
from openpyxl import load_workbook
import re

ABBY_FILE_NAME = "Tabla ABBY.xlsx"


def scrape_abby1File(wbName=None):

    wbName = wbName or ABBY_FILE_NAME

    # carga el archivo
    wb = load_workbook(filename=wbName, use_iterators=True)

    # creo un objeto abby
    abby = AbbyFile(wb)

    # leo el archivo
    abby.read()

    # parseo el contenido a una lista de diccionarios
    abby.parse()

    # vuelco el contenido en un excel
    abby.create_excel()

    # cierro el objeto
    # abby.close()

    return abby


def find_nth(s, x, n):
    i = -1
    for _ in range(n):
        i = s.find(x, i + len(x))
        if i == -1:
            break
    return i

def convert_to_float(strValue):
    strValue = strValue.strip().replace(".","").replace(",",".")
    floatValue = float(strValue)
    return floatValue

# USER CLASSES
class AbbyFile():

    def __init__(self, wb):

        self.wb = wb
        self.content = []
        self.records = []
        self.errors = []

    def read(self):

        # toma la hoja activa
        ws = self.wb.get_active_sheet()

        # lee optimizadamente todo el archivo
        for row in ws.iter_rows():

            # crea lista vacia que tendra los valuees de la fila
            valuesList = []

            for cell in row:

                # agrega value de la celda a la lista de la row
                valuesList.append(get_unicode(cell.internal_value))

            # si la lista no esta vacia procede a agregarla al contenido
            if not self._empty(valuesList):

                # elimina todas las celdas vacias del final de la lista
                valuesList = self._remove_lasts_none(valuesList)

                # agrega lista de valuees al contenido
                self.content.append(valuesList)

    def parse(self):

        # crea un objeto AbbyParser
        ap = AbbyParser()

        # itera entre las rows del content
        for row in self.content:

            # le da una linea para parsear al abby
            ap.parse_row(row)

            # toma los nuevos records que haya podido crear
            self.records.extend(ap.get_records())

            # toma los errores que se hayan detectado
            self.errors.extend(ap.get_errors())

    def create_excel(self):

        dict_list_to_excel(self.records, "ABBY parsed.xlsx")

    def close(self):
        pass

    def print_content(self, limit=10):

        i = 0
        for row in self.content:

            print row

            i += 1
            if i > limit:
                break

    def print_records(self, limit=100):
        pass


    # METODOS PRIVADOS
    def _empty(self, valuesList):
        """indica si una lista esta vacia"""

        RV = True

        for value in valuesList:
            if value != None:
                RV = False

        return RV

    def _remove_lasts_none(self, valuesList):
        """elimina todos los None values del final de una lista"""

        RV = valuesList

        while RV[-1] == None:
            del RV[-1]

        return RV


# INTERNAL CLASSES
class AbbyParser():

    # DATA
    STR_TITLE_TYPE = u"T\xcdTULO"
    STR_AGVALUE_TYPE = u"Valor total:"
    PATTERN_SUBT1 = "^[a-z][)][A-Z\s]{1,}"
    STR_HEAD1_TYPE = "Tarifa"
    PATTERN_SUBT2 = "^[0-9][\.]"
    PATTERN_HEAD1 = "^[0-9].{1,}Tarifa.{1,}:"
    PATTERN_HEAD1_INI_PART = "^[0-9].{1,}Tarifa.{1,}"
    PATTERN_HEAD1_FINAL_PART = ".{1,}:"
    STR_NONE_iMPORT_TYPE = u"Sin importaci\xf3n"

    def __init__(self):

        # Titulo
        self.nroTitulo = None
        self.descTitulo = None

        # Primer Subtitulo
        self.letraSubTitulo1 = None
        self.descSubTitulo1 = None

        # Segundo Subtitulo
        self.nroSubTitulo2 = None
        self.descSubTitulo2 = None

        # Tabla
        self.nroProducto = None
        self.nroTarifa = None
        self.descProducto = None
        self.unidadProducto = None

        # datos dentro de la Tabla
        self.idPais = None
        self.descPais = None
        self.year = []
        self.cantidad = []
        self.value = []

        # cursor
        self.typeRow = ""
        self.typeLastRow = ""
        self.lastRow = ""

        # resultados
        self.records = []
        self.errors = []

    def parse_row(self, row):

        # identifica el tipo de la row
        rowType = self._get_row_type(row)

        # extrae los datos de la row segun el tipo
        if rowType == "title":
            self._parse_title(row)

        elif rowType == "subt1":
            self._parse_subt1(row)

        elif rowType == "subt2":
            self._parse_subt2(row)

        elif rowType == "agValues":
            self._parse_agValues(row)

        elif rowType == "tblHead1":
            self._parse_tblHead1(row)

        elif rowType == "tblHead2":
            self._parse_tblHead2(row)

        elif rowType == "tblRow":
            self._parse_tblRow(row)

        elif rowType == "tblHead1IniPart":
            self._parse_tblHead1IniPart(row)

        elif rowType == "tblHead1FinalPart":
            self._parse_tblHead1FinalPart(row)

        elif rowType == "noneImport":
            self._parse_noneImport(row)

        elif rowType == "ignore":
            pass

        else:
            print "No se reconoce el tipo!!", row

        # construye los registros que se puedan construir
        self._build_records()

    def get_records(self):
        """devuelve la lista de records y la deja vacia"""

        RV = list(self.records)
        self.records = []

        return RV

    def get_errors(self):
        """devuelve la lista de errors y la deja vacia"""

        RV = list(self.errors)
        self.errors = []

        return RV

    # METODOS PRIVADOS
    def _get_row_type(self, row):

        if len(row) == 1 and self.STR_NONE_iMPORT_TYPE in row[0]:
            RV = "noneImport"

        elif len(row) == 1 and self.typeLastRow == "tblHead1IniPart" and \
            self._re_match(self.PATTERN_HEAD1_FINAL_PART, row[0]):

            RV = "tblHead1FinalPart"

        elif len(row) > 1:
            RV = "tblRow"

        elif len(row) == 1 and self._re_match(self.PATTERN_HEAD1, row[0]):
            RV = "tblHead1"

        elif len(row) == 1 and self._re_match(self.PATTERN_HEAD1_INI_PART, 
                                              row[0]):
            RV = "tblHead1IniPart"

        # elif len(row) == 1 and self._re_match(self.PATTERN_HEAD2,
                                              # row[0], re.U):
            # RV = "tblHead2"

        elif len(row) == 1 and self.STR_TITLE_TYPE in row[0]:
            RV = "title"

        elif len(row) == 1 and self.STR_AGVALUE_TYPE in row[0]:
            RV = "agValues"

        elif len(row) == 1 and self._re_match(self.PATTERN_SUBT1, row[0]):
            RV = "subt1"

        elif len(row) == 1 and self._re_match(self.PATTERN_SUBT2, row[0]):
            RV = "subt2"

        elif self._ignore_row(row):
            RV = "ignore"

        else:
            RV = None

        self.typeRow = RV

        return RV

    def _re_match(self, pattern, string):

        RV = False

        matchObj = re.match(pattern, string, re.U)

        if matchObj:
            RV = True

        return RV

    def _build_records(self):

        newRecords = []

        if self.typeRow == "agValues" or self.typeRow == "tblRow":

            i = 0
            for value in self.value:

                newRecord = dict()

                # Titulo
                newRecord["nroTitulo"] = self.nroTitulo
                newRecord["descTitulo"] = self.descTitulo

                # Primer Subtitulo
                newRecord["letraSubTitulo1"] = self.letraSubTitulo1
                newRecord["descSubTitulo1"] = self.descSubTitulo1

                # Segundo Subtitulo
                newRecord["nroSubTitulo2"] = self.nroSubTitulo2
                newRecord["descSubTitulo2"] = self.descSubTitulo2

                # Tabla
                newRecord["nroProducto"] = self.nroProducto
                newRecord["nroTarifa"] = self.nroTarifa
                newRecord["descProducto"] = self.descProducto
                newRecord["unidadProducto"] = self.unidadProducto

                # datos dentro de la Tabla
                newRecord["idPais"] = self.idPais
                newRecord["descPais"] = self.descPais
                newRecord["year"] = self.year[i]
                newRecord["cantidad"] = self.cantidad[i]
                newRecord["value"] = value

                # agrega el nuevo record
                newRecords.append(newRecord)
                i += 1

        self.records.extend(newRecords)

        self.typeLastRow = str(self.typeRow)

    def _parse_title(self, row):

        # crea el parser
        parser = TitleParser(row)

        # extrae los datos que devuelve el parser
        self.nroTitulo = parser.nroTitulo
        self.descTitulo = parser.descTitulo

        # guarda el tipo de row como el ultimo visto
        self.typeRow = "title"

        # toma los errores que se hayan detectado
        self.errors.extend(parser.errors)

    def _parse_subt1(self, row):

        # crea el parser
        parser = Subt1Parser(row)

        # extrae los datos que devuelve el parser
        self.letraSubTitulo1 = parser.letraSubTitulo1
        self.descSubTitulo1 = parser.descSubTitulo1

        # toma los errores que se hayan detectado
        self.errors.extend(parser.errors)

    def _parse_subt2(self, row):

        # crea el parser
        parser = Subt2Parser(row)

        # extrae los datos que devuelve el parser
        self.nroSubTitulo2 = parser.nroSubTitulo2
        self.descSubTitulo2 = parser.descSubTitulo2

        # toma los errores que se hayan detectado
        self.errors.extend(parser.errors)

    def _parse_agValues(self, row):

        # crea el parser
        parser = AgValuesParser(row)

        # si la ultima row fue un titulo, pone en None los subtitulos
        if self.typeLastRow == "title":

            # Primer Subtitulo
            self.letraSubTitulo1 = 0
            self.descSubTitulo1 = "Todos"

            # Segundo Subtitulo
            self.nroSubTitulo2 = 0
            self.descSubTitulo2 = "Todos"

        # si la ultima row fue un titulo, pone en None el subt2
        if self.typeLastRow == "subt1":

            # Segundo Subtitulo
            self.nroSubTitulo2 = 0
            self.descSubTitulo2 = "Todos"

        # Tabla
        self.nroProducto = 0
        self.nroTarifa = u"NA"
        self.descProducto = u"Todos"
        self.unidadProducto = u"NA"

        # datos dentro de la Tabla que no se tienen
        self.idPais = 0
        self.descPais = "Todos"
        self.cantidad = [None, None]

        # datos de la tabla que se tienen
        self.year = parser.year
        self.value = parser.value

        # toma los errores que se hayan detectado
        self.errors.extend(parser.errors)

    def _parse_tblHead1(self, row):

        # crea el parser
        parser = Head1Parser(row)

        # extrae los datos que devuelve el parser
        self.nroProducto = parser.nroProducto
        self.nroTarifa = parser.nroTarifa
        self.descProducto = parser.descProducto
        self.unidadProducto = parser.unidadProducto

        # toma los errores que se hayan detectado
        self.errors.extend(parser.errors)

    def _parse_tblHead2(self, row):

        # crea el parser
        parser = Head2Parser(row)

        # extrae los datos que devuelve el parser
        self.nroProducto = parser.nroProducto
        self.nroTarifa = parser.nroTarifa
        self.descProducto = parser.descProducto
        self.unidadProducto = parser.unidadProducto

        # toma los errores que se hayan detectado
        self.errors.extend(parser.errors)

    def _parse_tblRow(self, row):

        # crea el parser
        parser = TblRowParser(row)

        # datos dentro de la Tabla
        self.idPais = None
        self.descPais = parser.descPais
        self.year = parser.year
        self.cantidad = parser.cantidad
        self.value = parser.value

        # toma los errores que se hayan detectado
        self.errors.extend(parser.errors)

    def _parse_tblHead1IniPart(self, row):
        """toma la parte inicial de un head1 y la guarda hasta que
        aparezca la parte final"""

        self.lastRow = row[0]

    def _parse_tblHead1FinalPart(self, row):
        """toma la parte final de un head1 y la concatena con la parte
        inicial previamente guardada, luego llama al metodo de head1"""

        tblHead1Row = [get_unicode(self.lastRow + u" " + row[0])]

        # print tblHead1Row
        self._parse_tblHead1(tblHead1Row)

    def _parse_noneImport(self, row):
        """crea un registro similar a tlbRow pero sin datos"""
        
        # datos dentro de la Tabla
        self.idPais = 0
        self.descPais = u"Todos"
        self.year = [1945, 1946]
        self.cantidad = [u"NA", u"NA"]
        self.value = [u"NA", u"NA"]

    def _ignore_row(self, row):
        RV = False

        if (u"(Conclusi\xf3n)" in row[0]) or (u"" == row[0]):
            RV = True

        return RV


class TitleParser():

    def __init__(self, row):
        self.nroTitulo = self._get_nroTitulo(row)
        self.descTitulo = self._get_descTitulo(row)
        self.errors = [None]

    def _get_nroTitulo(self, row):

        i = row[0].find(".")

        RV = row[0][:i].replace(u'T\xcdTULO', "").strip()

        return RV

    def _get_descTitulo(self, row):

        i = row[0].find(".")
        pattern = "[A-Z][A-Z\s]{1,}"

        RV = re.search(pattern, row[0][i+1:], re.U).group().strip()

        return RV


class Subt1Parser():

    def __init__(self, row):
        self.letraSubTitulo1 = self._get_letraSubTitulo1(row)
        self.descSubTitulo1 = self._get_descSubTitulo1(row)
        self.errors = [None]

    def _get_letraSubTitulo1(self, row):

        i = row[0].find(")")

        RV = row[0][:i].strip()

        return RV

    def _get_descSubTitulo1(self, row):

        i = row[0].find(")")
        pattern = "[A-Z][A-Z\s]{1,}"

        RV = re.search(pattern, row[0][i+1:], re.U).group().strip()

        return RV


class Subt2Parser():

    def __init__(self, row):
        self.nroSubTitulo2 = self._get_nroSubTitulo2(row)
        self.descSubTitulo2 = self._get_descSubTitulo2(row)
        self.errors = [None]

    def _get_nroSubTitulo2(self, row):

        # print row
        i = row[0].strip().find(".")

        RV = row[0].strip()[:i]
        # print RV
        return RV

    def _get_descSubTitulo2(self, row):

        i = row[0].find(".")
        pattern = "[A-Z].{1,}"

        RV = re.search(pattern, row[0][i+1:], re.U).group().strip()

        return RV


class AgValuesParser():

    def __init__(self, row):
        self.year = self._get_year(row)
        self.value = self._get_value(row)
        self.errors = [None]

    def _get_year(self, row):
        RV = []

        # agrega el primer año
        i = row[0].find(":")
        j = row[0].find("m$n")

        RV.append(int(row[0][i+1:j].strip()))

        # agrega el segundo año
        i = row[0].find(";")
        j = row[0].rfind("m$n")

        RV.append(int(row[0][i+1:j].strip()))

        return RV

    def _get_value(self, row):
        RV = []

        # agrega el valor del primer año
        i = row[0].find("m$n")
        j = row[0].find(";")

        strValue = row[0][i+4:j]
        floatValue = convert_to_float(strValue)

        RV.append(floatValue)

        # agrega el valor del segundo año
        i = row[0].rfind("m$n")
        j = row[0].find(")")

        strValue = row[0][i+4:j]
        floatValue = convert_to_float(strValue)

        RV.append(floatValue)

        return RV


class Head1Parser():

    def __init__(self, row):
        self.nroProducto = self._get_nroProducto(row)
        self.nroTarifa = self._get_nroTarifa(row)
        self.descProducto = self._get_descProducto(row)
        self.unidadProducto = self._get_unidadProducto(row)
        self.errors = [None]

    def _get_nroProducto(self, row):

        indexList = [row[0].strip().find("."),
                     row[0].strip().find("-"),
                     row[0].strip().find("(")]

        # remueve el -1, si existe
        try:
            indexList.remove(-1)
        except:
            pass

        # calcula el indice minimo
        minIndex = min(indexList)

        # si es mayor que uno, lo utiliza
        if minIndex > 1:
            i = minIndex
        # si no, utiliza 1 como minimo absoluto
        else:
            i = 1

        RV = row[0].strip()[:i]

        return RV

    def _get_nroTarifa(self, row):

        if u"varios y no tarifados" in row[0]:
            RV = "varios y no tarifados"

        else:

            iStart = row[0].find(u"Tarifa") + 6
            # iEnd = find_nth(row[0], ".", 2) - 1
            iEnd = row[0].find(u")")

            RV = row[0][iStart:iEnd].strip()

        return RV

    def _get_descProducto(self, row):

        ### chequea que haya un "primer punto" antes de los parentesis
        # busca el primer punto que encuentra
        iDot = row[0].find(".")

        # busca los indices de los parentesis
        iParenthesis = [row[0].find("("), row[0].find(")")]

        # remueve el -1, si existe
        try:
            iParenthesis.remove(-1)
        except:
            pass
        
        # calcula el indice minimo
        minIndex = min(iParenthesis)

        # si el indice del primer punto es menor, esta antes de los parentesis
        if iDot < minIndex:
            iStart = find_nth(row[0], ".", 2) - 1

        # si no, hay que tomar entonces el primer punto (ya que falta el otro)
        else:
            iStart = find_nth(row[0], ".", 1) - 1        
        
        # busca el indice final de la descripcion a partir de la ultima coma
        iEnd = row[0].rfind(",")

        # regex pattern
        pattern = "[A-Z].{1,}"

        # intenta matchear el patron, sino devuelve error
        try:
            RV = re.search(pattern, row[0][iStart:iEnd], re.U).group().strip()
        except:
            RV = "Parsing error"

        return RV

    def _get_unidadProducto(self, row):

        iStart = row[0].rfind(",") + 1
        iEnd = row[0].rfind(":")

        RV = row[0][iStart:iEnd].strip()
        RV = self._homogeneizar_unidadProducto(RV)

        return RV

    def _homogeneizar_unidadProducto(self, unidadProducto):
        
        RV = unidadProducto

        if get_unicode(unidadProducto.strip()) == u"Kg.":
            RV = u"kilogramos"

        return RV


class Head2Parser():
    pass


class TblRowParser():

    def __init__(self, row):

        self.idPais = None
        self.descPais = self._get_descPais(row)
        self.year = self._get_year(row)
        self.cantidad = self._get_cantidad(row)
        self.value = self._get_value(row)
        self.errors = [None]

    def _get_descPais(self, row):

        if not row[0]:
            RV = "Missing error"

        elif u"Total" in row[0] or u"total" in row[0]:
            RV = u"Todos"

        else:
            RV = row[0].replace(".", "").strip()

        return RV

    def _get_year(self, row):

        RV = [u"1945", u"1946"]

        return RV

    def _get_cantidad(self, row):

        cant_1945 = None
        cant_1946 = None

        try:
            cant_1945 = convert_to_float(row[1])
        except:
            pass

        try:
            cant_1946 = convert_to_float(row[2])
        except:
            pass

        RV = [cant_1945, cant_1946]

        return RV

    def _get_value(self, row):

        value_1945 = None
        value_1946 = None

        try:
            value_1945 = convert_to_float(row[3])
        except:
            pass

        try:
            value_1946 = convert_to_float(row[4])
        except:
            pass

        RV = [value_1945, value_1946]

        return RV








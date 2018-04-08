#!/usr/bin/env
#-*- coding: utf-8 -*-

from xlrd import open_workbook 
from xlwt import easyxf 
from xlutils.copy import copy
    
col_index = 3

for j in range(1, 10):

    nazwa = "ZATR070" + str(j) + ".xls"
    unazwa = "ZATR070" + str(j) + "u.xls"
    rb = open_workbook(nazwa, formatting_info=True)
    r_sheet = rb.sheet_by_index(0)
    wb = copy(rb)
    w_sheet = wb.get_sheet(0)

    for row_index in range(1, r_sheet.nrows):
        komorka = r_sheet.cell(row_index, col_index).value
        if "KSIEGOWA" in komorka:
            w_sheet.write(row_index, col_index, "KSIEG.")
        elif "ST.ASYS" in komorka:
            w_sheet.write(row_index, col_index, "ST.ASYS.")
        elif "ST.T.ANA" in komorka:
            w_sheet.write(row_index, col_index, "ST.T.ANA.")
        elif "POM.APTECZNA" in komorka:
            w_sheet.write(row_index, col_index, "POM.APT.")    
        elif "ASYSTENT" in komorka:
            w_sheet.write(row_index, col_index, "ASYS.")
        elif "K.PR.BAK" in komorka:
            w_sheet.write(row_index, col_index, "KIER.PR.BAK.")    
        elif "REJ.MED" in komorka:
            w_sheet.write(row_index, col_index, "REJ.MED.")    
        elif "ML.ASYS" in komorka:
            w_sheet.write(row_index, col_index, "ML.ASYS.")    
        elif "KIER.CENT.UTRZ" in komorka:
            w_sheet.write(row_index, col_index, "KIER.CENT.UTRZ.")    
        elif "Mt'.ASYSTENT" in komorka:
            w_sheet.write(row_index, col_index, "ML.ASYS.")    
        elif "ST.T.FIZ" in komorka:
            w_sheet.write(row_index, col_index, "ST.T.FIZ.")    
        elif "ST.ASYS" in komorka:
            w_sheet.write(row_index, col_index, "ST.ASYS.")    
        elif "KIER.POR.CH.ST" in komorka:
            w_sheet.write(row_index, col_index, "KIER.POR.CH.")    
        elif "SPEC.D/S PRAC" in komorka:
            w_sheet.write(row_index, col_index, "SPEC. D/S PRAC.")    
        elif "S.STA.M" in komorka:
            w_sheet.write(row_index, col_index, "ST.STA.M.")    
        elif "NACZ.P." in komorka:
            w_sheet.write(row_index, col_index, "NACZ.PIEL.")    
        elif "KIER.S.RU.CHOR" in komorka:
            w_sheet.write(row_index, col_index, "KIER.S.RU.CHOR.")    
        elif "KIER.SEK" in komorka:
            w_sheet.write(row_index, col_index, "KIER.SEK.")
        elif "STAT.MED" in komorka:
            w_sheet.write(row_index, col_index, "STAT.MED.")
        elif "INSP.DS.ADM-GO" in komorka:
            w_sheet.write(row_index, col_index, "INSP. D/S ADM-GO")
        elif "SPECJAL." in komorka:
            w_sheet.write(row_index, col_index, "SPEC.")
        elif "ST.KSIEG" in komorka:
            w_sheet.write(row_index, col_index, "ST.KSIEG.")
        elif "SPEC D/S SOCJ" in komorka:
            w_sheet.write(row_index, col_index, "SPEC. D/S SOCJ.")
        elif "ROB.GOSP." in komorka:
            w_sheet.write(row_index, col_index, "PR.GOSP.")
        elif "P,POM.DENT" in komorka:
            w_sheet.write(row_index, col_index, "POM.DENT.")
        elif "KIE.S.TE" in komorka:
            w_sheet.write(row_index, col_index, "KIE.S.TE.")
        elif "KIER.ZES.T.MED" in komorka:
            w_sheet.write(row_index, col_index, "KIER.ZES.T.MED.")
        elif "Mt'.ASYSTENT" in komorka:
            w_sheet.write(row_index, col_index, "ML.ASYSTENT")
        elif "K.K.ORG-Gť.ENE" in komorka:
            w_sheet.write(row_index, col_index, "KIER.K.ORG-GL.ENE.")
        elif "ST.TECH.FARMAC" in komorka:
            w_sheet.write(row_index, col_index, "ST.T.FARMAC")
        elif "ST.TECH.ORT." in komorka:
            w_sheet.write(row_index, col_index, "ST.T.ORT.")
        elif "TECH.FIZ" in komorka:
            w_sheet.write(row_index, col_index, "T.FIZ.")
        elif "Mt'.ASYST." in komorka:
            w_sheet.write(row_index, col_index, "ML.ASYS.")
        elif "Z-CA DYR.D/S L" in komorka:
            w_sheet.write(row_index, col_index, "Z-CA DYR. D/S L.")
        elif "ST.ASYS" in komorka:
            w_sheet.write(row_index, col_index, "ST.ASYS.")
        elif "BRYGARDZISTA" in komorka:
            w_sheet.write(row_index, col_index, "BRYGADZISTA")
        elif "TECH.ANA" in komorka:
            w_sheet.write(row_index, col_index, "T.ANA.")
        elif "TECH.RTG" in komorka:
            w_sheet.write(row_index, col_index, "T.RTG")
        elif "OPER.KOTťúW" in komorka:
            w_sheet.write(row_index, col_index, "OPER.KOTLOW.")
        elif "PIEL.ODDZIAť." in komorka:
            w_sheet.write(row_index, col_index, "PIEL.ODDZ.")
        elif "KIER.DZ.íYWIEN" in komorka:
            w_sheet.write(row_index, col_index, "KIER.DZ.ZYWIEN.")
        elif "ST.DIETYCZKA" in komorka:
            w_sheet.write(row_index, col_index, "ST.DIETETYCZKA")
        elif "ELEKTROMECHANI" in komorka:
            w_sheet.write(row_index, col_index, "ELEKTROMECHANIK")
        elif "Mť.ASYSTENT" in komorka:
            w_sheet.write(row_index, col_index, "ML.ASYS.")
        elif "ST.TECH.FARMAC" in komorka:
            w_sheet.write(row_index, col_index, "ST.T.FARMAC")
        elif "KIER.ZESPOťU" in komorka:
            w_sheet.write(row_index, col_index, "KIER.ZES.")
        elif "RATOWMIK MED" in komorka:
            w_sheet.write(row_index, col_index, "RATOWNIK MED.")
        elif "Mť.ASYST.-REZ." in komorka:
            w_sheet.write(row_index, col_index, "ML.ASYS.-REZ.") 
        elif "PIEL.ODDZIAť." in komorka:
            w_sheet.write(row_index, col_index, "PIEL.ODDZ.") 
        elif "Mť.AS.PIEL.EPI" in komorka:
            w_sheet.write(row_index, col_index, "ML.ASYS.PIEL.EPI.") 
        elif "KIER.SEK.KSIÉG" in komorka:
            w_sheet.write(row_index, col_index, "KIER.SEK.KSIEG.") 
        elif "INSP D/S P.POí" in komorka:
            w_sheet.write(row_index, col_index, "INSP D/S P.PO.") 
        elif "SPRZĆT." in komorka:
            w_sheet.write(row_index, col_index, "SPRZAT.") 
        elif "TECHN.INFORMAT" in komorka:
            w_sheet.write(row_index, col_index, "T.INFORMAT.") 
        elif "TECH.NARZ.WZR." in komorka:
            w_sheet.write(row_index, col_index, "T.NARZ.WZR.") 
        elif "Gť.KSIÉG" in komorka:
            w_sheet.write(row_index, col_index, "GL.KSIEG.")  
        elif "REF.D/S KSIÉG." in komorka:
            w_sheet.write(row_index, col_index, "REF. D/S KSIEG.")
        elif "TECH.FIZJOTER." in komorka:
            w_sheet.write(row_index, col_index, "T.FIZ.")
        elif "INFORMATYK" in komorka:
            w_sheet.write(row_index, col_index, "INFORMAT.")
    
    wb.save(unazwa)

#------------------------------------------
rb = open_workbook("ZATR0710.xls", formatting_info=True)
r_sheet = rb.sheet_by_index(0)
wb = copy(rb)
w_sheet = wb.get_sheet(0)

for row_index in range(1, r_sheet.nrows):
    komorka = r_sheet.cell(row_index, col_index).value
    if "KSIEGOWA" in komorka:
        w_sheet.write(row_index, col_index, "KSIEG.")
    elif "ST.ASYS" in komorka:
        w_sheet.write(row_index, col_index, "ST.ASYS.")
    elif "ST.T.ANA" in komorka:
        w_sheet.write(row_index, col_index, "ST.T.ANA.")
    elif "POM.APTECZNA" in komorka:
        w_sheet.write(row_index, col_index, "POM.APT.")    
    elif "ASYSTENT" in komorka:
        w_sheet.write(row_index, col_index, "ASYS.")
    elif "K.PR.BAK" in komorka:
        w_sheet.write(row_index, col_index, "KIER.PR.BAK.")    
    elif "REJ.MED" in komorka:
        w_sheet.write(row_index, col_index, "REJ.MED.")    
    elif "ML.ASYS" in komorka:
        w_sheet.write(row_index, col_index, "ML.ASYS.")    
    elif "KIER.CENT.UTRZ" in komorka:
        w_sheet.write(row_index, col_index, "KIER.CENT.UTRZ.")    
    elif "Mt'.ASYSTENT" in komorka:
        w_sheet.write(row_index, col_index, "ML.ASYS.")    
    elif "ST.T.FIZ" in komorka:
        w_sheet.write(row_index, col_index, "ST.T.FIZ.")    
    elif "ST.ASYS" in komorka:
        w_sheet.write(row_index, col_index, "ST.ASYS.")    
    elif "KIER.POR.CH.ST" in komorka:
        w_sheet.write(row_index, col_index, "KIER.POR.CH.")    
    elif "SPEC.D/S PRAC" in komorka:
        w_sheet.write(row_index, col_index, "SPEC. D/S PRAC.")    
    elif "S.STA.M" in komorka:
        w_sheet.write(row_index, col_index, "ST.STA.M.")    
    elif "NACZ.P." in komorka:
        w_sheet.write(row_index, col_index, "NACZ.PIEL.")    
    elif "KIER.S.RU.CHOR" in komorka:
        w_sheet.write(row_index, col_index, "KIER.S.RU.CHOR.")    
    elif "KIER.SEK" in komorka:
        w_sheet.write(row_index, col_index, "KIER.SEK.")
    elif "STAT.MED" in komorka:
        w_sheet.write(row_index, col_index, "STAT.MED.")
    elif "INSP.DS.ADM-GO" in komorka:
        w_sheet.write(row_index, col_index, "INSP. D/S ADM-GO")
    elif "SPECJAL." in komorka:
        w_sheet.write(row_index, col_index, "SPEC.")
    elif "ST.KSIEG" in komorka:
        w_sheet.write(row_index, col_index, "ST.KSIEG.")
    elif "SPEC D/S SOCJ" in komorka:
        w_sheet.write(row_index, col_index, "SPEC. D/S SOCJ.")
    elif "ROB.GOSP." in komorka:
        w_sheet.write(row_index, col_index, "PR.GOSP.")
    elif "P,POM.DENT" in komorka:
        w_sheet.write(row_index, col_index, "POM.DENT.")
    elif "KIE.S.TE" in komorka:
        w_sheet.write(row_index, col_index, "KIE.S.TE.")
    elif "KIER.ZES.T.MED" in komorka:
        w_sheet.write(row_index, col_index, "KIER.ZES.T.MED.")
    elif "Mt'.ASYSTENT" in komorka:
        w_sheet.write(row_index, col_index, "ML.ASYSTENT")
    elif "K.K.ORG-Gť.ENE" in komorka:
        w_sheet.write(row_index, col_index, "KIER.K.ORG-GL.ENE.")
    elif "ST.TECH.FARMAC" in komorka:
        w_sheet.write(row_index, col_index, "ST.T.FARMAC")
    elif "ST.TECH.ORT." in komorka:
        w_sheet.write(row_index, col_index, "ST.T.ORT.")
    elif "TECH.FIZ" in komorka:
        w_sheet.write(row_index, col_index, "T.FIZ.")
    elif "Mt'.ASYST." in komorka:
        w_sheet.write(row_index, col_index, "ML.ASYS.")
    elif "Z-CA DYR.D/S L" in komorka:
        w_sheet.write(row_index, col_index, "Z-CA DYR. D/S L.")
    elif "ST.ASYS" in komorka:
        w_sheet.write(row_index, col_index, "ST.ASYS.")
    elif "BRYGARDZISTA" in komorka:
        w_sheet.write(row_index, col_index, "BRYGADZISTA")
    elif "TECH.ANA" in komorka:
        w_sheet.write(row_index, col_index, "T.ANA.")
    elif "TECH.RTG" in komorka:
        w_sheet.write(row_index, col_index, "T.RTG")
    elif "OPER.KOTťúW" in komorka:
        w_sheet.write(row_index, col_index, "OPER.KOTLOW.")
    elif "PIEL.ODDZIAť." in komorka:
        w_sheet.write(row_index, col_index, "PIEL.ODDZ.")
    elif "KIER.DZ.íYWIEN" in komorka:
        w_sheet.write(row_index, col_index, "KIER.DZ.ZYWIEN.")
    elif "ST.DIETYCZKA" in komorka:
        w_sheet.write(row_index, col_index, "ST.DIETETYCZKA")
    elif "ELEKTROMECHANI" in komorka:
        w_sheet.write(row_index, col_index, "ELEKTROMECHANIK")
    elif "Mť.ASYSTENT" in komorka:
        w_sheet.write(row_index, col_index, "ML.ASYS.")
    elif "ST.TECH.FARMAC" in komorka:
        w_sheet.write(row_index, col_index, "ST.T.FARMAC")
    elif "KIER.ZESPOťU" in komorka:
        w_sheet.write(row_index, col_index, "KIER.ZES.")
    elif "RATOWMIK MED" in komorka:
        w_sheet.write(row_index, col_index, "RATOWNIK MED.")
    elif "Mť.ASYST.-REZ." in komorka:
        w_sheet.write(row_index, col_index, "ML.ASYS.-REZ.") 
    elif "PIEL.ODDZIAť." in komorka:
        w_sheet.write(row_index, col_index, "PIEL.ODDZ.") 
    elif "Mť.AS.PIEL.EPI" in komorka:
        w_sheet.write(row_index, col_index, "ML.ASYS.PIEL.EPI.") 
    elif "KIER.SEK.KSIÉG" in komorka:
        w_sheet.write(row_index, col_index, "KIER.SEK.KSIEG.") 
    elif "INSP D/S P.POí" in komorka:
        w_sheet.write(row_index, col_index, "INSP D/S P.PO.") 
    elif "SPRZĆT." in komorka:
        w_sheet.write(row_index, col_index, "SPRZAT.") 
    elif "TECHN.INFORMAT" in komorka:
        w_sheet.write(row_index, col_index, "T.INFORMAT.") 
    elif "TECH.NARZ.WZR." in komorka:
        w_sheet.write(row_index, col_index, "T.NARZ.WZR.") 
    elif "Gť.KSIÉG" in komorka:
        w_sheet.write(row_index, col_index, "GL.KSIEG.")  
    elif "REF.D/S KSIÉG." in komorka:
        w_sheet.write(row_index, col_index, "REF. D/S KSIEG.")
    elif "TECH.FIZJOTER." in komorka:
        w_sheet.write(row_index, col_index, "T.FIZ.")
    elif "INFORMATYK" in komorka:
        w_sheet.write(row_index, col_index, "INFORMAT.")
        
wb.save("ZATR0710u.xls")
from openpyxl import load_workbook
import qrcode

#Load the excel file
wb = load_workbook("01.xlsx")

#Add the sheet name
sheet_info = wb['qry_Beneficiaries']

#qr_gen function is responsible for generating QR codes with MEMCard formate
def qr_gen(name = None, card = None, phone = None, nid = None, union = None, ward = None) :
    qname   = name
    qcard   = "CARD: " + str(card)
    qphone  = phone
    qnid    = "NID: " + str(nid)
    qunion  = "UNION: " + str(int(union)) if union != None else union
    qward   = "WARD: " + str(int(ward)) if ward != None else ward

    img = qrcode.make(f"MECARD:N:{qname};ADR:{qcard};TEL:{qphone};EMAIL:{qnid};ADR:{qunion};ADR:{qward};;")
    type(img)
    img.save(f"qrfiles/{card}.png")

#Bulky processing
for row in range(2, 1697) :
    str_row = str(row)
    qr_gen(
        sheet_info["I" + str_row].value,
        sheet_info["A" + str_row].value,
        sheet_info["G" + str_row].value,
        sheet_info["H" + str_row].value,
        sheet_info["K" + str_row].value,
        sheet_info["F" + str_row].value,
    )

"""Blatt „Start“: Rechtsform-Auswahl, Hero, Kacheln, Ablauf."""
from openpyxl.styles import Alignment, Border, Font, Protection
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink

from design_pro import INPUT_BG, INPUT_FG, INPUT_LINE, SANS, TEAL, TEAL_L, fill, side

def purchase_selector(wb):
    """Direktauswahl Privat / Kapitalgesellschaft auf der Startseite (steuert den Namen „Rechtsform“)."""
    ws, s09 = wb["Start"], wb["S09 Steuern"]
    current = s09["D12"].value
    ws.row_dimensions[23].height = 34
    ws.row_dimensions[22].height = 3.75
    lab = ws["C23"]
    lab.value = "KAUF ALS"
    lab.font = Font(name=SANS, sz=9, b=True, color=TEAL)
    lab.alignment = Alignment(horizontal="right", vertical="center", indent=1)
    ws.merge_cells("D23:G23")
    sel = ws["D23"]
    sel.value = current
    sel.font = Font(name=SANS, sz=11, b=True, color=INPUT_FG)
    sel.fill = fill(INPUT_BG)
    sel.alignment = Alignment(horizontal="left", vertical="center", indent=1, shrink_to_fit=True)
    sel.protection = Protection(locked=False)
    ln = side("medium", INPUT_LINE)
    for col in "DEFG":
        ws[f"{col}23"].fill = fill(INPUT_BG)
        ws[f"{col}23"].border = Border(top=ln, bottom=ln, left=ln if col == "D" else None, right=ln if col == "G" else None)
    dv = DataValidation(type="list", formula1="=L_Rechtsform", allow_blank=False, showDropDown=False)
    dv.promptTitle = "Kaufstruktur"
    dv.prompt = "Privatperson oder Kapitalgesellschaft (vermögensverwaltende bzw. gewerbliche GmbH) wählen."
    dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add("D23")
    wb.defined_names["Rechtsform"] = DefinedName("Rechtsform", attr_text="Start!$D$23")
    # Schritt 9 zeigt die Auswahl nur noch an
    for dvs in list(s09.data_validations.dataValidation):
        if "D12" in str(dvs.sqref):
            s09.data_validations.dataValidation.remove(dvs)
    s09["D12"].value = "=Rechtsform"
    s09["D12"].protection = Protection(locked=True)
    s09["D12"].fill = fill(TEAL_L)
    s09["D12"].font = Font(name=SANS, sz=10, b=True, color=TEAL)
    s09["D12"].border = Border()
    s09["D12"].hyperlink = Hyperlink(ref="D12", location="'Start'!D23", display="Auswahl auf der Startseite")
    s09["F12"].value = "Auswahl direkt auf der Startseite („Kauf als“) – Klick auf das Feld öffnet sie. " + str(s09["F12"].value or "")
    s09["F12"].data_type = "s"


def apply(wb):
    purchase_selector(wb)

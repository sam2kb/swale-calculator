#!/usr/bin/env python3
from __future__ import annotations

import argparse
from copy import copy
from pathlib import Path
from typing import Union

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.datavalidation import DataValidation


def make(out: Union[str, Path] = "swale_calculator.xlsx") -> Path:
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True  # force Excel recalculation on open

    normal_style = wb._named_styles[0]
    normal_font = copy(normal_style.font)
    normal_font.name = "Roboto"
    normal_font.size = 10
    normal_style.font = normal_font

    ws = wb.active
    ws.title = "Swale Calculator"
    ws.page_margins = PageMargins(left=0.25, right=0.25, top=0.25, bottom=0.25, header=0.25, footer=0.25)

    # -----------------------------
    # Styles (simple, inline)
    # -----------------------------
    title_font = Font(bold=True, size=14)
    header_font = Font(bold=True, size=11)
    bold_font = Font(bold=True)

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    # Use opaque ARGB fills (FF......) so Excel shows them
    fill_header = PatternFill("solid", fgColor=Color(theme=9, tint=0.8))
    fill_sub = PatternFill("solid", fgColor="FFFAFAFA")

    thin = Side(style="thin", color="FF888888")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def set_col_width(widths: dict[str, float]) -> None:
        for col, width in widths.items():
            ws.column_dimensions[col].width = width

    def apply_border(cell_range: str) -> None:
        for row in ws[cell_range]:
            for cell in row:
                cell.border = border

    # -----------------------------
    # Layout
    # -----------------------------
    set_col_width(
        {
            "A": 20,
            "B": 11,
            "C": 10,
            "D": 10,
            "E": 10,
            "F": 11,
            "G": 11,
            "H": 12,
            "I": 11,
        }
    )

    # -----------------------------
    # SITE STATISTICS
    # -----------------------------
    ws["A1"] = "SITE STATISTICS"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.row_dimensions[1].height = 22
    ws.merge_cells("A1:C1")

    ws["A2"] = "Description"
    ws["B2"] = "Sq. Feet"
    ws["C2"] = "Acres"
    ws["D2"] = "Percent"
    for addr in ("A2", "B2", "C2", "D2"):
        ws[addr].font = header_font
        ws[addr].alignment = center
        ws[addr].fill = fill_header

    site_rows = [
        ("Lot size", 3),
        ("Building (incl. patios)", 4),
        ("Driveway", 5),
        ("Equipment pads", 6),
        ("Other impervious", 7),
    ]

    for label, r in site_rows:
        ws[f"A{r}"] = label
        ws[f"A{r}"].alignment = left

        ws[f"B{r}"].number_format = "#,##0"
        ws[f"B{r}"].alignment = right

        ws[f"C{r}"] = f'=IF(OR(B{r}="",B{r}=0),"",B{r}/43560)'
        ws[f"C{r}"].number_format = "0.000"
        ws[f"C{r}"].alignment = right

        # Percent: blank for Lot size row, else fraction of lot
        ws[f"D{r}"] = "" if r == 3 else f'=IF(OR(B{r}="",B{r}=0,$B$3="",$B$3=0),"",B{r}/$B$3)'
        ws[f"D{r}"].number_format = "0.0%"
        ws[f"D{r}"].alignment = right

    # Seed example values (edit as needed)
    ws["B3"] = 6741
    ws["B4"] = 3224
    ws["B5"] = 640
    ws["B6"] = 9
    ws["B7"] = ""

    ws["A9"] = "Impervious area"
    ws["B9"] = "=SUM(B4:B7)"
    ws["C9"] = '=IF(OR(B9="",B9=0),"",B9/43560)'
    ws["D9"] = '=IF(OR(B9="",B9=0,$B$3="",$B$3=0),"",B9/$B$3)'
    for addr in ("A9", "B9", "C9", "D9"):
        ws[addr].font = bold_font
        ws[addr].alignment = right if addr != "A9" else left
    ws["B9"].number_format = "#,##0"
    ws["C9"].number_format = "0.000"
    ws["D9"].number_format = "0.0%"

    ws["A10"] = "Open space"
    ws["B10"] = "=$B$3-$B$9"
    ws["C10"] = '=IF(OR(B10="",B10=0),"",B10/43560)'
    ws["D10"] = '=IF(OR(B10="",B10=0,$B$3="",$B$3=0),"",B10/$B$3)'
    for addr in ("A10", "B10", "C10", "D10"):
        ws[addr].font = bold_font
        ws[addr].alignment = right if addr != "A10" else left
    ws["B10"].number_format = "#,##0"
    ws["C10"].number_format = "0.000"
    ws["D10"].number_format = "0.0%"

    apply_border("A2:D10")

    # -----------------------------
    # STORMWATER RETENTION
    # -----------------------------
    ws["A13"] = "STORMWATER RETENTION"
    ws["A13"].font = title_font
    ws["A13"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    ws.row_dimensions[13].height = 22
    ws.merge_cells("A13:C13")

    ws["A24"] = "Retention basis"
    ws["C24"] = "MAX"
    ws["C24"].alignment = center
    ws.merge_cells("C24:D24")

    dv_basis = DataValidation(type="list", formula1='"MAX,LOT ONLY,IMP ONLY"', allow_blank=False)
    ws.add_data_validation(dv_basis)
    dv_basis.add(ws["C24"])

    ws["A14"] = "Required (cf)"
    ws["A14"].font = bold_font
    ws["A14"].alignment = left
    ws["B14"] = (
        '=IF($C$24="LOT ONLY",(0.5/12)*$B$3,'
        'IF($C$24="IMP ONLY",(1/12)*$B$9,'
        'MAX((0.5/12)*$B$3,(1/12)*$B$9)))'
    )
    ws["B14"].number_format = "#,##0.0"
    ws["B14"].alignment = right

    ws["A15"] = "Provided (cf)"
    ws["A15"].font = bold_font
    ws["A15"].alignment = left
    # Provided = sum volumes where Select=YES (H19:H22, I19:I22)
    ws["B15"] = '=SUMPRODUCT(--($I$19:$I$22="YES"),$H$19:$H$22)'
    ws["B15"].number_format = "#,##0.0"
    ws["B15"].alignment = right

    ws["A16"] = CellRichText(
        TextBlock(InlineFont(b=True), "FHA Type B:"),
        TextBlock(InlineFont(b=False), " Drainage to front and rear"),
    )
    ws["A16"].alignment = left
    ws.merge_cells("A16:C16")

    apply_border("A14:B15")

    ws["E14"] = (
        '="Required retention uses the larger of two methods:"&CHAR(10)&'
        '"       (a) 1/2"" over total lot area "&TEXT($B$3,"#,##0")&" sf"&CHAR(10)&'
        '"       (b) 1"" over impervious area "&TEXT($B$9,"#,##0")&" sf"'
    )
    ws.merge_cells("E14:I16")
    ws["E14"].alignment = Alignment(horizontal="left", vertical="bottom", wrap_text=True)

    # -----------------------------
    # SWALES TABLE
    # -----------------------------
    header_row = 18
    headers = [
        ("Swale", "A"),
        ("Type", "B"),
        ("Bot W (ft)", "C"),
        ("Bot L (ft)", "D"),
        ("Width (ft)", "E"),
        ("Length (ft)", "F"),
        ("Depth (ft)", "G"),
        ("Volume (cf)", "H"),
        ("Select", "I"),
    ]

    for text, col in headers:
        cell = f"{col}{header_row}"
        ws[cell] = text
        ws[cell].font = header_font
        ws[cell].alignment = center
        ws[cell].fill = fill_header

    swale_defaults = [
        # Trapezoid/Frustum: bottom WxL in C/D, top WxL in E/F, depth in G (ft)
        dict(name="Swale A", type="Trapezoid", bw=6, bl=35, tw=9, tl=38, depth_ft=0.75),
        dict(name="Swale B", type="Trapezoid", bw=2, bl=16, tw=8, tl=22, depth_ft=1.00),
        # V-shape: E=top width, F=length, G=depth (ft)
        dict(name="Swale C", type="V-Shape", bw=None, bl=None, tw=8, tl=48, depth_ft=1.00),
        dict(name="Swale D", type="V-Shape", bw=None, bl=None, tw=8, tl=24, depth_ft=1.00),
    ]

    for i, s in enumerate(swale_defaults):
        r = header_row + 1 + i  # 19..22
        ws[f"A{r}"] = s["name"]
        ws[f"A{r}"].alignment = left

        ws[f"B{r}"] = s["type"]
        ws[f"B{r}"].fill = fill_sub
        ws[f"B{r}"].alignment = center

        # inputs
        for col in ("C", "D", "E", "F", "G"):
            ws[f"{col}{r}"].alignment = right

        if s["bw"] is not None:
            ws[f"C{r}"] = s["bw"]
        if s["bl"] is not None:
            ws[f"D{r}"] = s["bl"]
        ws[f"E{r}"] = s["tw"]
        ws[f"F{r}"] = s["tl"]
        ws[f"G{r}"] = s["depth_ft"]

        # Volume formula:
        # - V-Shape: 0.5*TopWidth*Depth*Length
        # - Trapezoid/Frustum: h/3*(A1+A2+sqrt(A1*A2)), A1=C*D, A2=E*F, h=G
        ws[f"H{r}"] = (
            f'=IF(UPPER($B{r})="V-SHAPE",'
            f'IF(OR(E{r}="",F{r}="",G{r}=""),"",0.5*E{r}*G{r}*F{r}),'
            f'IF(OR(C{r}="",D{r}="",E{r}="",F{r}="",G{r}=""),"",'
            f'G{r}/3*((C{r}*D{r})+(E{r}*F{r})+SQRT((C{r}*D{r})*(E{r}*F{r}))))'
            f')'
        )
        ws[f"H{r}"].number_format = "#,##0.0"
        ws[f"H{r}"].alignment = right

        ws[f"I{r}"] = "YES" if i < 2 else "NO"
        ws[f"I{r}"].alignment = center

    # YES/NO dropdown
    dv_use = DataValidation(type="list", formula1='"YES,NO"', allow_blank=False)
    ws.add_data_validation(dv_use)
    dv_use.add("I19:I22")

    apply_border("A18:I22")

    out_path = Path(out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))
    return out_path


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Generate Swale Calculator workbook.")
    p.add_argument("--out", default="swale_calculator.xlsx", help="Output .xlsx path.")
    return p.parse_args()


if __name__ == "__main__":
    args = parse_args()
    out_path = make(args.out)
    print(f"Wrote: {out_path}")
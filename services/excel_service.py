from __future__ import annotations

from datetime import datetime
from io import BytesIO
from typing import Iterable
from uuid import uuid4

from openpyxl import load_workbook

FALTANTES_SHEET = "_APP_FALTANTES"
CORTES_SHEET = "_APP_CORTES"
OFICIOS_SHEET = "_APP_OFICIOS"

FALTANTES_HEADERS = [
    "ID", "Requerimiento", "Contrato", "Procedimiento", "Codigo_documento",
    "Documento", "Fecha_deteccion", "Auditor", "Corte"
]
CORTES_HEADERS = ["ID", "Requerimiento", "Corte", "Fecha", "Creado_por", "Documentos"]
OFICIOS_HEADERS = ["ID", "Requerimiento", "Fecha", "Cortes", "Referencia", "Creado_por"]


def _load(data: bytes):
    return load_workbook(BytesIO(data))


def _save(wb) -> bytes:
    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def _ensure_technical_sheets(wb) -> None:
    specs = [
        (FALTANTES_SHEET, FALTANTES_HEADERS),
        (CORTES_SHEET, CORTES_HEADERS),
        (OFICIOS_SHEET, OFICIOS_HEADERS),
    ]
    for name, headers in specs:
        if name not in wb.sheetnames:
            ws = wb.create_sheet(name)
            ws.append(headers)
        ws = wb[name]
        ws.sheet_state = "hidden"


def list_requirements(data: bytes) -> list[str]:
    wb = _load(data)
    return [s for s in wb.sheetnames if s.startswith("Req_")]


def list_contracts(data: bytes, requirement: str) -> list[dict]:
    wb = _load(data)
    ws = wb[requirement]
    rows = []
    for row in range(8, ws.max_row + 1):
        number = ws.cell(row, 1).value
        contract = ws.cell(row, 2).value
        if number is None or not contract:
            continue
        rows.append({
            "row": row,
            "numero": number,
            "contrato": str(contract).strip(),
            "obra": ws.cell(row, 3).value or "",
            "contratista": ws.cell(row, 4).value or "",
            "faltantes_visible": ws.cell(row, 5).value or "",
            "auditor": ws.cell(row, 6).value or "",
            "ulop": ws.cell(row, 7).value or "",
        })
    return rows


def load_catalog(data: bytes, procedure: str) -> list[dict]:
    wb = _load(data)
    ws = wb[procedure]
    items = []
    for row in range(3, ws.max_row + 1):
        code = ws.cell(row, 1).value
        concept = ws.cell(row, 2).value
        if not code or not concept:
            continue
        prefix = str(code).split("_")[0]
        items.append({"codigo": str(code).strip(), "documento": str(concept).strip(), "etapa": prefix})
    return items


def get_faltantes(data: bytes, requirement: str, contract: str) -> list[dict]:
    wb = _load(data)
    _ensure_technical_sheets(wb)
    ws = wb[FALTANTES_SHEET]
    headers = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
    result = []
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, headers["Requerimiento"]).value == requirement and ws.cell(r, headers["Contrato"]).value == contract:
            result.append({h: ws.cell(r, c).value for h, c in headers.items()})
    return result


def document_history(data: bytes, requirement: str, contract: str) -> dict[str, dict]:
    """Resumen por código: cortes históricos y si ya está pendiente del siguiente corte."""
    history: dict[str, dict] = {}
    for row in get_faltantes(data, requirement, contract):
        code = str(row.get("Codigo_documento") or "")
        if not code:
            continue
        entry = history.setdefault(code, {"cuts": [], "pending": False, "documento": row.get("Documento")})
        cut = row.get("Corte")
        if str(cut) == "PENDIENTE":
            entry["pending"] = True
        else:
            try:
                cut_num = int(cut)
                if cut_num not in entry["cuts"]:
                    entry["cuts"].append(cut_num)
            except (TypeError, ValueError):
                pass
    for entry in history.values():
        entry["cuts"].sort()
    return history


def add_faltantes(data: bytes, requirement: str, contract: str, procedure: str, auditor: str, selected: Iterable[dict]) -> bytes:
    """Agrega una nueva solicitud. Un documento histórico puede repetirse; uno ya PENDIENTE no."""
    wb = _load(data)
    _ensure_technical_sheets(wb)
    tech = wb[FALTANTES_SHEET]

    pending_codes = set()
    for r in range(2, tech.max_row + 1):
        if tech.cell(r, 2).value == requirement and tech.cell(r, 3).value == contract:
            if str(tech.cell(r, 9).value) == "PENDIENTE":
                pending_codes.add(str(tech.cell(r, 5).value))

    now = datetime.now().replace(microsecond=0)
    for item in selected:
        if item["codigo"] in pending_codes:
            continue
        tech.append([
            uuid4().hex, requirement, contract, procedure, item["codigo"], item["documento"],
            now, auditor, "PENDIENTE"
        ])
        pending_codes.add(item["codigo"])

    visible_ws = wb[requirement]
    target_row = None
    for r in range(8, visible_ws.max_row + 1):
        if str(visible_ws.cell(r, 2).value or "").strip() == contract:
            target_row = r
            break
    if target_row is None:
        raise ValueError("No se encontró el contrato en la hoja seleccionada.")

    # La celda visible conserva una lista única; el historial técnico sí conserva reiteraciones.
    docs = []
    for r in range(2, tech.max_row + 1):
        if tech.cell(r, 2).value == requirement and tech.cell(r, 3).value == contract:
            doc = tech.cell(r, 6).value
            if doc and str(doc) not in docs:
                docs.append(str(doc))
    visible_ws.cell(target_row, 5).value = "\n".join(docs)
    return _save(wb)


def pending_summary(data: bytes, requirement: str) -> list[dict]:
    wb = _load(data)
    _ensure_technical_sheets(wb)
    ws = wb[FALTANTES_SHEET]
    result = []
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 2).value == requirement and str(ws.cell(r, 9).value) == "PENDIENTE":
            result.append({
                "contrato": ws.cell(r, 3).value,
                "codigo": ws.cell(r, 5).value,
                "documento": ws.cell(r, 6).value,
                "fecha": ws.cell(r, 7).value,
                "auditor": ws.cell(r, 8).value,
            })
    return result


def list_cuts(data: bytes, requirement: str) -> list[dict]:
    wb = _load(data)
    _ensure_technical_sheets(wb)
    ws = wb[CORTES_SHEET]
    result = []
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 2).value == requirement:
            result.append({
                "id": ws.cell(r, 1).value,
                "corte": int(ws.cell(r, 3).value),
                "fecha": ws.cell(r, 4).value,
                "creado_por": ws.cell(r, 5).value,
                "documentos": int(ws.cell(r, 6).value or 0),
            })
    return sorted(result, key=lambda x: x["corte"])


def create_cut(data: bytes, requirement: str, user: str, cut_date) -> bytes:
    wb = _load(data)
    _ensure_technical_sheets(wb)
    falt = wb[FALTANTES_SHEET]
    cuts = wb[CORTES_SHEET]

    existing = [int(cuts.cell(r, 3).value) for r in range(2, cuts.max_row + 1) if cuts.cell(r, 2).value == requirement and cuts.cell(r, 3).value]
    next_cut = max(existing, default=0) + 1
    affected = 0
    for r in range(2, falt.max_row + 1):
        if falt.cell(r, 2).value == requirement and str(falt.cell(r, 9).value) == "PENDIENTE":
            falt.cell(r, 9).value = next_cut
            affected += 1
    if affected == 0:
        raise ValueError("No hay documentación pendiente para crear un corte.")

    cuts.append([uuid4().hex, requirement, next_cut, cut_date, user, affected])
    return _save(wb)


def documents_for_cuts(data: bytes, requirement: str, cuts_selected: Iterable[int]) -> list[dict]:
    selected = {int(x) for x in cuts_selected}
    wb = _load(data)
    _ensure_technical_sheets(wb)
    ws = wb[FALTANTES_SHEET]
    result = []
    for r in range(2, ws.max_row + 1):
        cut = ws.cell(r, 9).value
        try:
            cut_int = int(cut)
        except (TypeError, ValueError):
            continue
        if ws.cell(r, 2).value == requirement and cut_int in selected:
            result.append({
                "contrato": ws.cell(r, 3).value,
                "documento": ws.cell(r, 6).value,
                "corte": cut_int,
            })
    return result

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
    "Documento", "Especificacion", "Fecha_deteccion", "Auditor", "Corte"
]
CORTES_HEADERS = ["ID", "Requerimiento", "Corte", "Fecha", "Creado_por", "Documentos"]
OFICIOS_HEADERS = ["ID", "Requerimiento", "Fecha", "Cortes", "Referencia", "Creado_por"]


def _load(data: bytes):
    return load_workbook(BytesIO(data))


def _save(wb) -> bytes:
    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def _ensure_headers(ws, headers: list[str]) -> None:
    current = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    if not any(current):
        ws.append(headers)
        return
    for header in headers:
        if header not in current:
            ws.cell(1, ws.max_column + 1).value = header
            current.append(header)


def _headers(ws) -> dict[str, int]:
    return {str(ws.cell(1, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value}


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
        else:
            ws = wb[name]
            _ensure_headers(ws, headers)
        ws.sheet_state = "hidden"


def _effective_document(documento, especificacion) -> str:
    doc = str(documento or "").strip()
    spec = str(especificacion or "").strip()
    if doc and spec:
        return f"{doc} ({spec})"
    return doc or spec


def _rebuild_visible_cell(wb, requirement: str, contract: str) -> None:
    tech = wb[FALTANTES_SHEET]
    h = _headers(tech)
    docs: list[str] = []
    for r in range(2, tech.max_row + 1):
        if tech.cell(r, h["Requerimiento"]).value == requirement and str(tech.cell(r, h["Contrato"]).value or "").strip() == contract:
            text = _effective_document(
                tech.cell(r, h["Documento"]).value,
                tech.cell(r, h["Especificacion"]).value if "Especificacion" in h else "",
            )
            if text and text not in docs:
                docs.append(text)

    visible_ws = wb[requirement]
    target_row = None
    for r in range(8, visible_ws.max_row + 1):
        if str(visible_ws.cell(r, 2).value or "").strip() == contract:
            target_row = r
            break
    if target_row is None:
        raise ValueError("No se encontró el contrato en la hoja seleccionada.")
    visible_ws.cell(target_row, 5).value = "\n".join(docs)


def _update_cut_count(wb, requirement: str, cut_number: int) -> None:
    falt = wb[FALTANTES_SHEET]
    cuts = wb[CORTES_SHEET]
    fh = _headers(falt)
    ch = _headers(cuts)
    count = 0
    for r in range(2, falt.max_row + 1):
        if falt.cell(r, fh["Requerimiento"]).value == requirement:
            try:
                current = int(falt.cell(r, fh["Corte"]).value)
            except (TypeError, ValueError):
                continue
            if current == int(cut_number):
                count += 1
    for r in range(2, cuts.max_row + 1):
        if cuts.cell(r, ch["Requerimiento"]).value == requirement and int(cuts.cell(r, ch["Corte"]).value) == int(cut_number):
            cuts.cell(r, ch["Documentos"]).value = count
            break


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


def get_faltantes(data: bytes, requirement: str, contract: str | None = None) -> list[dict]:
    wb = _load(data)
    _ensure_technical_sheets(wb)
    ws = wb[FALTANTES_SHEET]
    h = _headers(ws)
    result = []
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, h["Requerimiento"]).value != requirement:
            continue
        if contract is not None and str(ws.cell(r, h["Contrato"]).value or "").strip() != contract:
            continue
        row = {name: ws.cell(r, col).value for name, col in h.items()}
        row["_row"] = r
        row["Documento_efectivo"] = _effective_document(row.get("Documento"), row.get("Especificacion"))
        result.append(row)
    return result


def document_history(data: bytes, requirement: str, contract: str) -> dict[str, dict]:
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
    wb = _load(data)
    _ensure_technical_sheets(wb)
    tech = wb[FALTANTES_SHEET]
    h = _headers(tech)

    pending_codes = set()
    for r in range(2, tech.max_row + 1):
        if tech.cell(r, h["Requerimiento"]).value == requirement and str(tech.cell(r, h["Contrato"]).value or "").strip() == contract:
            if str(tech.cell(r, h["Corte"]).value) == "PENDIENTE":
                pending_codes.add(str(tech.cell(r, h["Codigo_documento"]).value))

    now = datetime.now().replace(microsecond=0)
    for item in selected:
        if item["codigo"] in pending_codes:
            continue
        row = [None] * tech.max_column
        values = {
            "ID": uuid4().hex,
            "Requerimiento": requirement,
            "Contrato": contract,
            "Procedimiento": procedure,
            "Codigo_documento": item["codigo"],
            "Documento": item["documento"],
            "Especificacion": str(item.get("especificacion") or "").strip(),
            "Fecha_deteccion": now,
            "Auditor": auditor,
            "Corte": "PENDIENTE",
        }
        for name, value in values.items():
            row[h[name] - 1] = value
        tech.append(row)
        pending_codes.add(item["codigo"])

    _rebuild_visible_cell(wb, requirement, contract)
    return _save(wb)


def pending_summary(data: bytes, requirement: str) -> list[dict]:
    result = []
    for row in get_faltantes(data, requirement):
        if str(row.get("Corte")) == "PENDIENTE":
            result.append({
                "id": row.get("ID"),
                "contrato": row.get("Contrato"),
                "codigo": row.get("Codigo_documento"),
                "documento": row.get("Documento"),
                "especificacion": row.get("Especificacion") or "",
                "solicitud": row.get("Documento_efectivo"),
                "fecha": row.get("Fecha_deteccion"),
                "auditor": row.get("Auditor"),
            })
    return result


def list_cuts(data: bytes, requirement: str) -> list[dict]:
    wb = _load(data)
    _ensure_technical_sheets(wb)
    ws = wb[CORTES_SHEET]
    h = _headers(ws)
    result = []
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, h["Requerimiento"]).value == requirement:
            result.append({
                "id": ws.cell(r, h["ID"]).value,
                "corte": int(ws.cell(r, h["Corte"]).value),
                "fecha": ws.cell(r, h["Fecha"]).value,
                "creado_por": ws.cell(r, h["Creado_por"]).value,
                "documentos": int(ws.cell(r, h["Documentos"]).value or 0),
            })
    return sorted(result, key=lambda x: x["corte"])


def cut_details(data: bytes, requirement: str, cut_number: int) -> list[dict]:
    result = []
    for row in get_faltantes(data, requirement):
        try:
            current_cut = int(row.get("Corte"))
        except (TypeError, ValueError):
            continue
        if current_cut == int(cut_number):
            result.append({
                "id": row.get("ID"),
                "contrato": row.get("Contrato"),
                "auditor": row.get("Auditor"),
                "codigo": row.get("Codigo_documento"),
                "documento": row.get("Documento"),
                "especificacion": row.get("Especificacion") or "",
                "solicitud": row.get("Documento_efectivo"),
                "fecha": row.get("Fecha_deteccion"),
                "corte": current_cut,
            })
    return result


def create_cut(data: bytes, requirement: str, user: str, cut_date) -> bytes:
    wb = _load(data)
    _ensure_technical_sheets(wb)
    falt = wb[FALTANTES_SHEET]
    cuts = wb[CORTES_SHEET]
    fh = _headers(falt)
    ch = _headers(cuts)

    existing = [
        int(cuts.cell(r, ch["Corte"]).value)
        for r in range(2, cuts.max_row + 1)
        if cuts.cell(r, ch["Requerimiento"]).value == requirement and cuts.cell(r, ch["Corte"]).value
    ]
    next_cut = max(existing, default=0) + 1
    affected = 0
    for r in range(2, falt.max_row + 1):
        if falt.cell(r, fh["Requerimiento"]).value == requirement and str(falt.cell(r, fh["Corte"]).value) == "PENDIENTE":
            falt.cell(r, fh["Corte"]).value = next_cut
            affected += 1
    if affected == 0:
        raise ValueError("No hay documentación pendiente para crear un corte.")

    row = [None] * cuts.max_column
    values = {
        "ID": uuid4().hex, "Requerimiento": requirement, "Corte": next_cut,
        "Fecha": cut_date, "Creado_por": user, "Documentos": affected,
    }
    for name, value in values.items():
        row[ch[name] - 1] = value
    cuts.append(row)
    return _save(wb)


def delete_pending_records(data: bytes, requirement: str, record_ids: Iterable[str]) -> bytes:
    ids = {str(x) for x in record_ids}
    if not ids:
        return data
    wb = _load(data)
    _ensure_technical_sheets(wb)
    ws = wb[FALTANTES_SHEET]
    h = _headers(ws)
    affected_contracts: set[str] = set()
    rows_to_delete = []
    for r in range(2, ws.max_row + 1):
        if (
            ws.cell(r, h["Requerimiento"]).value == requirement
            and str(ws.cell(r, h["ID"]).value) in ids
            and str(ws.cell(r, h["Corte"]).value) == "PENDIENTE"
        ):
            affected_contracts.add(str(ws.cell(r, h["Contrato"]).value or "").strip())
            rows_to_delete.append(r)
    for r in reversed(rows_to_delete):
        ws.delete_rows(r, 1)
    for contract in affected_contracts:
        _rebuild_visible_cell(wb, requirement, contract)
    return _save(wb)


def move_records_to_cut(data: bytes, requirement: str, record_ids: Iterable[str], cut_number: int) -> bytes:
    ids = {str(x) for x in record_ids}
    wb = _load(data)
    _ensure_technical_sheets(wb)
    falt = wb[FALTANTES_SHEET]
    fh = _headers(falt)
    changed = 0
    for r in range(2, falt.max_row + 1):
        if falt.cell(r, fh["Requerimiento"]).value == requirement and str(falt.cell(r, fh["ID"]).value) in ids:
            falt.cell(r, fh["Corte"]).value = int(cut_number)
            changed += 1
    if changed == 0:
        raise ValueError("No se encontraron registros para agregar al corte.")
    _update_cut_count(wb, requirement, cut_number)
    return _save(wb)


def remove_records_from_cut(data: bytes, requirement: str, record_ids: Iterable[str], cut_number: int) -> bytes:
    ids = {str(x) for x in record_ids}
    wb = _load(data)
    _ensure_technical_sheets(wb)
    falt = wb[FALTANTES_SHEET]
    fh = _headers(falt)
    changed = 0
    for r in range(2, falt.max_row + 1):
        if falt.cell(r, fh["Requerimiento"]).value != requirement or str(falt.cell(r, fh["ID"]).value) not in ids:
            continue
        try:
            current_cut = int(falt.cell(r, fh["Corte"]).value)
        except (TypeError, ValueError):
            continue
        if current_cut == int(cut_number):
            falt.cell(r, fh["Corte"]).value = "PENDIENTE"
            changed += 1
    if changed == 0:
        raise ValueError("No se encontraron documentos para retirar del corte.")
    _update_cut_count(wb, requirement, cut_number)
    return _save(wb)


def delete_cut(data: bytes, requirement: str, cut_number: int) -> bytes:
    wb = _load(data)
    _ensure_technical_sheets(wb)
    falt = wb[FALTANTES_SHEET]
    cuts = wb[CORTES_SHEET]
    fh = _headers(falt)
    ch = _headers(cuts)

    for r in range(2, falt.max_row + 1):
        if falt.cell(r, fh["Requerimiento"]).value != requirement:
            continue
        try:
            current_cut = int(falt.cell(r, fh["Corte"]).value)
        except (TypeError, ValueError):
            continue
        if current_cut == int(cut_number):
            falt.cell(r, fh["Corte"]).value = "PENDIENTE"

    deleted = False
    for r in range(2, cuts.max_row + 1):
        if cuts.cell(r, ch["Requerimiento"]).value == requirement and int(cuts.cell(r, ch["Corte"]).value) == int(cut_number):
            cuts.delete_rows(r, 1)
            deleted = True
            break
    if not deleted:
        raise ValueError("No se encontró el corte seleccionado.")
    return _save(wb)


def documents_for_cuts(data: bytes, requirement: str, cuts_selected: Iterable[int]) -> list[dict]:
    selected = {int(x) for x in cuts_selected}
    result = []
    for row in get_faltantes(data, requirement):
        try:
            cut_int = int(row.get("Corte"))
        except (TypeError, ValueError):
            continue
        if cut_int in selected:
            result.append({
                "contrato": row.get("Contrato"),
                "documento": row.get("Documento_efectivo"),
                "corte": cut_int,
            })
    return result

"""Identificación de campos de MS Project (estándar y personalizados)."""
from __future__ import annotations

import re
import unicodedata

# Identificadores (FieldID del XML = constantes PjField de VBA pjTaskTextN /
# pjTaskNumberN) de los campos personalizados de tarea.
TASK_TEXT_IDS: dict[int, int] = {
    1: 188743731, 2: 188743734, 3: 188743737, 4: 188743740, 5: 188743743,
    6: 188743746, 7: 188743747, 8: 188743748, 9: 188743749, 10: 188743750,
    **{i: 188743997 + (i - 11) for i in range(11, 31)},          # Text11..Text30
}
TASK_NUMBER_IDS: dict[int, int] = {
    1: 188743767, 2: 188743768, 3: 188743769, 4: 188743770, 5: 188743771,
    **{i: 188743982 + (i - 6) for i in range(6, 21)},            # Number6..Number20
}


def _norm(text: str) -> str:
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"\s+", "", text).lower()


def builtin_kind(name: str) -> str | None:
    """'WBS' u 'OutlineNumber' si el nombre corresponde a esos campos estándar (nombre en inglés)."""
    n = _norm(name)
    if n == "wbs":
        return "WBS"
    if n in ("outlinenumber", "numerodeesquema", "numeroesquema"):
        return "OutlineNumber"
    return None


def is_spanish_wbs_name(name: str) -> bool:
    """'EDT' es el nombre del campo WBS en MS Project en español.

    Se usa solo si ningún campo personalizado se llama así: es habitual crear un
    campo personalizado con el alias 'EDT' para cargar códigos propios.
    """
    return _norm(name) == "edt"


def is_notes_name(name: str) -> bool:
    """'Notas' (o 'Notes'): el texto de las notas de la tarea. Se usa solo si ningún campo personalizado se llama
    así, como 'EDT'."""
    return _norm(name) in ("notas", "nota", "notes", "note")


def parse_custom_name(name: str) -> tuple[str, int] | None:
    """'Text10', 'Texto10', 'Number1', 'Número1' -> ('Text', 10) / ('Number', 1)."""
    m = re.fullmatch(r"(text|texto|number|numero)(\d{1,2})", _norm(name))
    if not m:
        return None
    kind = "Text" if m.group(1).startswith("tex") else "Number"
    idx = int(m.group(2))
    limit = 30 if kind == "Text" else 20
    if not 1 <= idx <= limit:
        return None
    return kind, idx


def field_id(kind: str, idx: int) -> int | None:
    return (TASK_TEXT_IDS if kind == "Text" else TASK_NUMBER_IDS).get(idx)


def kind_of_field_id(fid: int) -> tuple[str, int] | None:
    for k, v in TASK_TEXT_IDS.items():
        if v == fid:
            return "Text", k
    for k, v in TASK_NUMBER_IDS.items():
        if v == fid:
            return "Number", k
    return None


def same_name(a: str, b: str) -> bool:
    return _norm(a) == _norm(b)


VISIBLE_YES = {"si", "s", "yes", "y", "x", "1", "true", "verdadero"}
VISIBLE_NO = {"no", "n", "0", "false", "falso"}


def parse_visible(value) -> tuple[bool | None, bool]:
    """Valor del campo Visible: (True = Sí, False = No, None = vacío o no reconocido; True si se reconoce).

    Admite Sí, Si, SI, S, X, 1 (y Yes, True, Verdadero) y No, NO, N, 0 (y False, Falso), con o sin tilde.
    """
    if isinstance(value, bool):
        return value, True
    if isinstance(value, (int, float)):
        value = f"{value:g}"
    text = _norm(str(value or ""))
    if not text:
        return None, True
    if text in VISIBLE_YES:
        return True, True
    if text in VISIBLE_NO:
        return False, True
    return None, False


def parse_number(text) -> float | None:
    """Convierte números con formato local ('1.234,5', '1,234.5', '12,5', '12.5')."""
    if text is None:
        return None
    if isinstance(text, (int, float)):
        return float(text)
    s = str(text).strip().replace(" ", "").replace(" ", "")
    s = re.sub(r"[^\d,.\-]", "", s)
    if not s or s in ("-", ".", ","):
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".") if s.count(",") == 1 else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None

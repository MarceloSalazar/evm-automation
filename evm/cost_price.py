"""Depuración de costes: precio de las EDT interna con los recargos de la tabla Recargos ([depuracion_precio]).

La empresa escribe su plan de costes a coste en la hoja PV-depuración: cada EDT interna lleva debajo sus componentes
de coste (líneas con solo la columna D). Con [depuracion_precio], cada mes de una línea con componentes es su precio:
la expresión calculo_precio, con los nombres de los componentes ([depuracion_precio.componentes]: el texto de su
columna D) y de los coeficientes ([depuracion_precio.recargos]: una fila de la tabla Recargos, arriba de la hoja, con
un % por año). Por ejemplo:

    calculo_precio = "(A*(1+KA)+B*(1+KB)+C*(1+KC)+D*(1+KD)+E)*(1+KG)*(1+BE)"

Los textos se comparan sin distinguir mayúsculas, tildes ni espacios de más (norm_label). procesar-costes --calculos
escribe en cada mes la fórmula de Excel equivalente (cost_review.write_calcs, render): cada componente es la suma de
sus líneas en ese mes y cada coeficiente, la celda de su fila y del año del mes. Un componente con otro texto no entra
en el precio (DEP-022). Este módulo no lee ni escribe el libro: cost_review lee la tabla y arma Price.
"""
from __future__ import annotations

import ast
import re
import unicodedata
from collections.abc import Callable
from dataclasses import dataclass, field

from openpyxl.utils import get_column_letter

from evm.config import ConfigError

NAME = re.compile(r"[A-Za-z][A-Za-z0-9_]*\Z")
NUMBER = re.compile(r"(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?\Z")
MAX_RATES = 8                           # filas de la tabla Recargos: 2 a 9 de la hoja PV-depuración
_OPS = {ast.Add: "+", ast.Sub: "-", ast.Mult: "*", ast.Div: "/"}
_PREC = {ast.Add: 1, ast.Sub: 1, ast.Mult: 2, ast.Div: 2}
SYNTAX = "solo admite números (con punto decimal), nombres, + - * / y paréntesis"


def norm_label(text) -> str:
    """Texto sin mayúsculas, tildes ni espacios de más: 'Costes  indirectos de Producción' -> 'costes indirectos de
    produccion'."""
    s = unicodedata.normalize("NFKD", "" if text is None else str(text))
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return " ".join(s.casefold().split())


def parse_formula(text: str, names: set[str]) -> ast.Expression:
    """Expresión de calculo_precio. names: nombres admitidos (en mayúsculas). ConfigError (CFG-055) con el motivo si
    no es una expresión válida o usa otro nombre."""
    source = text.strip()
    try:
        tree = ast.parse(source, mode="eval")
    except SyntaxError:
        raise ConfigError("CFG-055", f"no es una expresión válida: {SYNTAX}") from None
    for node in ast.walk(tree):
        if isinstance(node, (ast.Expression, ast.BinOp, ast.UnaryOp, ast.Load, ast.UAdd, ast.USub, *_OPS)):
            continue
        if isinstance(node, ast.Constant) and type(node.value) in (int, float):
            segment = ast.get_source_segment(source, node) or ""
            if not NUMBER.match(segment):           # 1_000 o 0x10: Python los admite, Excel no
                raise ConfigError("CFG-055", f"el número '{segment}' no es válido: {SYNTAX}")
            continue
        if isinstance(node, ast.Name):
            if not NAME.match(node.id) or node.id.upper() not in names:
                raise ConfigError("CFG-055", f"'{node.id}' no es un nombre de [depuracion_precio.componentes] ni "
                                             "de [depuracion_precio.recargos]")
            continue
        raise ConfigError("CFG-055", SYNTAX)
    return tree


def formula_names(tree: ast.Expression) -> set[str]:
    """Nombres (en mayúsculas) que usa la expresión."""
    return {n.id.upper() for n in ast.walk(tree) if isinstance(n, ast.Name)}


def evaluate(tree: ast.Expression, values: dict[str, float]) -> float:
    """Valor de la expresión; un nombre sin valor es 0, como una celda vacía en Excel. ZeroDivisionError si divide
    por 0 (Excel muestra #¡DIV/0!)."""
    def ev(node) -> float:
        if isinstance(node, ast.Expression):
            return ev(node.body)
        if isinstance(node, ast.Constant):
            return float(node.value)
        if isinstance(node, ast.Name):
            return values.get(node.id.upper(), 0.0)
        if isinstance(node, ast.UnaryOp):
            return -ev(node.operand) if isinstance(node.op, ast.USub) else ev(node.operand)
        a, b = ev(node.left), ev(node.right)
        op = type(node.op)
        return a + b if op is ast.Add else a - b if op is ast.Sub else a * b if op is ast.Mult else a / b
    return ev(tree)


def _number(value) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return repr(value)


def render(tree: ast.Expression, ref: Callable[[str], str]) -> str:
    """Expresión de Excel (sin el '='), con cada nombre cambiado por ref(nombre) y los paréntesis necesarios."""
    def r(node) -> str:
        if isinstance(node, ast.Expression):
            return r(node.body)
        if isinstance(node, ast.Constant):
            return _number(node.value)
        if isinstance(node, ast.Name):
            return ref(node.id.upper())
        if isinstance(node, ast.UnaryOp):
            inner = r(node.operand)
            if isinstance(node.operand, ast.BinOp):
                inner = f"({inner})"
            return ("-" if isinstance(node.op, ast.USub) else "") + inner
        p = _PREC[type(node.op)]
        left, right = r(node.left), r(node.right)
        if isinstance(node.left, ast.BinOp) and _PREC[type(node.left.op)] < p:
            left = f"({left})"
        if isinstance(node.right, ast.BinOp) and (_PREC[type(node.right.op)] < p or (
                _PREC[type(node.right.op)] == p and type(node.op) in (ast.Sub, ast.Div))):
            right = f"({right})"
        return f"{left}{_OPS[type(node.op)]}{right}"
    return r(tree)


@dataclass
class Price:
    """Cálculo del precio de un libro: la expresión, los componentes y la tabla Recargos leída de la hoja."""
    tree: ast.Expression
    components: dict[str, str]                      # texto normalizado -> nombre del componente
    rates: dict[str, str]                           # nombre del coeficiente -> texto de la configuración
    years: list[int]                                # año de cada columna de meses de la hoja PV-depuración
    rate_rows: dict[str, int] = field(default_factory=dict)      # coeficiente -> fila de la tabla
    year_cols: dict[int, int] = field(default_factory=dict)      # año -> columna de la tabla
    values: dict[tuple[str, int], float] = field(default_factory=dict)  # (coeficiente, año) -> valor

    def component(self, label: str) -> str | None:
        """Nombre del componente de una línea (texto de la columna D, normalizado); None si no está en la lista."""
        return self.components.get(label)

    def year(self, i: int) -> int:
        return self.years[min(i, len(self.years) - 1)]

    def used_rates(self) -> list[str]:
        """Coeficientes que usa la expresión, en el orden de la configuración."""
        names = formula_names(self.tree)
        return [n for n in self.rates if n in names]

    def rate_values(self, i: int) -> dict[str, float]:
        """Valor de cada coeficiente en el mes i (base 0); los que faltan en la tabla, 0."""
        y = self.year(i)
        return {n: self.values.get((n, y), 0.0) for n in self.rates}

    def rate_ref(self, name: str, i: int) -> str:
        """Celda del coeficiente en el año del mes i ('$E$2'); '0' si la tabla no tiene su fila o el año."""
        row, c = self.rate_rows.get(name), self.year_cols.get(self.year(i))
        return f"${get_column_letter(c)}${row}" if row and c else "0"

    def month_price(self, i: int, amounts: dict[str, float]) -> float:
        """Precio del mes i con los importes de cada componente."""
        try:
            return evaluate(self.tree, {**self.rate_values(i), **amounts})
        except ZeroDivisionError:
            return 0.0

    def month_formula(self, i: int, refs: dict[str, str]) -> str:
        """Fórmula de Excel del precio del mes i (sin el '='). refs: celdas de cada componente ('G15',
        'SUM(G15:G16)'); un componente sin líneas es 0."""
        return render(self.tree, lambda n: refs.get(n, "0") if n in self.component_names else self.rate_ref(n, i))

    @property
    def component_names(self) -> set[str]:
        return set(self.components.values())


def check_config(formula: str, components: dict[str, str], rates: dict[str, str],
                 max_rates: int = MAX_RATES) -> str | None:
    """Motivo por el que [depuracion_precio] no es válida (None si lo es). components y rates: nombre (en mayúsculas)
    -> texto."""
    if not formula:
        return ("falta calculo_precio, la expresión del precio (por ejemplo, calculo_precio = "
                "\"(A*(1+KA)+B*(1+KB))*(1+KG)*(1+BE)\")")
    if not components:
        return "falta [depuracion_precio.componentes], con un nombre = \"texto de la columna D\" por componente"
    both = sorted(set(components) & set(rates))
    if both:
        return f"{', '.join(both)} está en [depuracion_precio.componentes] y en [depuracion_precio.recargos]"
    if len(rates) > max_rates:
        return f"[depuracion_precio.recargos] admite hasta {max_rates} coeficientes (filas de la tabla Recargos)"
    for section, names in (("componentes", components), ("recargos", rates)):
        seen: dict[str, str] = {}
        for name, label in names.items():
            key = norm_label(label)
            if not key:
                return f"[depuracion_precio.{section}] {name}: falta el texto"
            if key in seen:
                return f"[depuracion_precio.{section}] {seen[key]} y {name} tienen el mismo texto ('{label}')"
            seen[key] = name
    try:
        tree = parse_formula(formula, set(components) | set(rates))
    except ConfigError as exc:
        return f"calculo_precio = \"{formula}\": {exc}"
    if not formula_names(tree) & set(components):
        return f"calculo_precio = \"{formula}\" no usa ningún componente de [depuracion_precio.componentes]"
    return None

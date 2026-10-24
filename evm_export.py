"""Punto de entrada: python evm_export.py <comando> [opciones]

Comandos: validar, generar, actualizar, analizar y generar-fichas. El detalle se obtiene con --help.
"""
from evm.cli import main

if __name__ == "__main__":
    raise SystemExit(main())

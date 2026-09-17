# -*- coding: utf-8 -*-
"""
Barra de progreso. Usa tqdm si esta instalado; si no, imprime una
barra propia para no obligarte a instalar nada.
"""
import sys
import time

try:
    from tqdm import tqdm as _tqdm
    HAY_TQDM = True
except ImportError:
    HAY_TQDM = False


class BarraSimple:
    """Barra de respaldo, sin dependencias."""

    def __init__(self, total, descripcion="Procesando"):
        self.total = max(1, total)
        self.descripcion = descripcion
        self.n = 0
        self.inicio = time.time()
        self._pintar()

    def update(self, cantidad=1):
        self.n += cantidad
        self._pintar()

    def set_postfix_str(self, texto):
        self.postfijo = texto

    def _pintar(self):
        proporcion = min(1.0, self.n / self.total)
        ancho = 34
        lleno = int(ancho * proporcion)
        barra = "#" * lleno + "-" * (ancho - lleno)
        transcurrido = time.time() - self.inicio
        if self.n > 0:
            restante = transcurrido / self.n * (self.total - self.n)
            eta = f"{int(restante // 60):02d}:{int(restante % 60):02d}"
        else:
            eta = "--:--"
        sys.stdout.write(
            f"\r{self.descripcion} |{barra}| {self.n}/{self.total} "
            f"({proporcion*100:5.1f}%) faltan {eta}"
        )
        sys.stdout.flush()

    def close(self):
        sys.stdout.write("\n")
        sys.stdout.flush()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()


def barra(total, descripcion="Procesando"):
    """Devuelve una barra de progreso, la mejor que haya disponible."""
    if HAY_TQDM:
        return _tqdm(total=total, desc=descripcion, unit="correo", ncols=90)
    return BarraSimple(total, descripcion)

"""PORTFOLIO — regra de contagem de cartas omitida."""
from portfolio_omitted import omit


def processar_planilha(*args, **kwargs):
    omit("regra de negocio: contagem de cartas (1-10 autos = 1 carta)")


def exportar_resultados(*args, **kwargs):
    omit("exportacao da contagem de cartas")


def extrair_datas_planilha(*args, **kwargs):
    omit("extrair_datas_planilha")


def __getattr__(name):
    def _m(*a, **k):
        omit(f"contagem_cartas.{name}")

    return _m

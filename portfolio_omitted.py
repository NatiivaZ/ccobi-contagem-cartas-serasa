"""Versao PORTFOLIO — implementacao operacional omitida."""


class PortfolioOmittedError(RuntimeError):
    def __init__(self, feature: str = "implementacao operacional") -> None:
        super().__init__(
            f"OMITIDO PARA PORTFOLIO — {feature}. Codigo completo so no ambiente do autor. "
            "Contato: ruan.natividade1@icloud.com | LinkedIn: ruan-natividade"
        )


def omit(feature: str = "implementacao operacional"):
    raise PortfolioOmittedError(feature)

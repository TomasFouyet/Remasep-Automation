class RemasepError(Exception):
    """Error base del dominio REMASEP."""


class UnknownClassificationError(RemasepError):
    def __init__(self, row_reference: str | None = None):
        message = "Ninguna regla coincide con el registro."
        if row_reference:
            message += f" Registro: {row_reference}"
        super().__init__(message)


class AmbiguousClassificationError(RemasepError):
    def __init__(self, rule_ids: list[str], row_reference: str | None = None):
        message = f"Más de una regla coincide con el registro: {', '.join(rule_ids)}."
        if row_reference:
            message += f" Registro: {row_reference}"
        super().__init__(message)


class SourceValidationError(RemasepError):
    pass


class TemplateValidationError(RemasepError):
    pass

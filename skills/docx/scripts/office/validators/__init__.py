"""Standard-library structural rules and bounded body text revision checks."""

from .structure import StructureValidator
from .redlining import RedliningValidator

__all__ = ["StructureValidator", "RedliningValidator"]

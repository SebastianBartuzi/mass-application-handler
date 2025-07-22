from enum import Enum

class ApplicationPlaceholderEnum(Enum):
    DATE: str = "{{data}}"
    SALUTATION: str = "{{zwrot}}"
    ADDRESSEE_NAME: str = "{{adresatimienazwisko}}"
    ADDRESSEE_TITLE: str = "{{tytulwlodarza}}"
    AUTHORITY_OFFICE_NAME: str = "{{urzadnazwa}}"
from enum import Enum

class AttentionInformation(Enum):
    MAIL_NOT_DELIVERED: str = "E-mail nie dotarł do nadawcy."
    WRONG_ADDRESSEE: str = "Dane nadawcy są złe."
    ACTION_REQUIRED: str = "Wymagana jest akcja."
    DEADLINE_EXTENDED: str = "Gmina zawnioskowała o wydłużenie terminu na odpowiedź."
    REFUSED_TO_ANSWER_FULLY: str = "Gmina odmówiła odpowiedzi na wszystkie pytania."
    REFUSED_TO_ANSWER_PARTIALLY: str = "Gmina odmówiła odpowiedzi na część pytań."
    PART_ANSWERED_SEPARATELY: str = "Gmina wyśle część odpowiedzi w osobnej korespondencji."
    NO_TERYT_MATCHED: str = "Nie dopasowano kodu TERYT do wiadomości."
    MULTIPLE_TERYTS_MATCHED: str = "Dopasowano kilka kodów TERYT do wiadomości."
from typing import List


class TypeConverter:
    def __init__(self) -> None:
        pass

    def str_to_list(self, str_to_convert: str) -> List:
        try:
            return list(str_to_convert)
        except Exception as e:
            raise ValueError(f"Nie udało się przekonwertować string'u {str_to_convert} na listę: {e}")

    def str_to_int(self, str_to_convert: str) -> int:
        try:
            return int(str_to_convert)
        except Exception as e:
            raise ValueError(f"Nie udało się przekonwertować string'u {str_to_convert} na int: {e}")
import json
from typing import Optional, Union, Dict, List

class TypeConverter:

    def str_to_dict_or_list(self, str_to_convert: str) -> Optional[Union[Dict, List]]:

        try:

            parsed_data = json.loads(str_to_convert)

            if isinstance(parsed_data, (dict, list)):
                return parsed_data

            else:
                return None

        except json.JSONDecodeError:
            return None

        except TypeError:
            return None
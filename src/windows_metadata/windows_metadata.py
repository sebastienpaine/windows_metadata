"""This module contains the WindowsAttribute class and all its helper functions"""
from os import PathLike
from typing import Dict, List
from pathlib import Path
from win32com.client.gencache import EnsureDispatch


_attr_list = []


def _fill_attr_list(namespace) -> None:
    """Gets list of all possible metadata attributes
    this is lazily loaded the first time a WindowsAttributes object is created"""
    for col in range(321):
        _attr_list.append(namespace.GetDetailsOf(None, col))


def _parse_attr_name(attr: str) -> str:
    """replaces spaces with underscores"""
    return attr.strip().replace(" ", "_").lower()


def _upper_each(val: str) -> str:
    """Upper Cases Each Word In A Given String"""
    words = val.split(" ")
    return_val = ""
    for word in words:
        word = list(word)
        first_char = word[0].upper()
        return_val.join(first_char).join(word[1:])
    return return_val


def _upper_first(val: str) -> str:
    """Uppercases the first letter of a given string"""
    return_val = ""
    words = list(val)
    first_char = words[0].upper()
    return_val.join(first_char).join(words[1:])
    return return_val


def _remove_underscore(val: str) -> str:
    """replaces underscores with spaces"""
    return val.replace("_", " ")


class WindowsAttributes:
    """This class loads and stores all the Windows metadata information for a given file"""

    def __init__(self, file_path: PathLike) -> None:
        self.__dict__["attr_dict"] = {}
        _path = Path(file_path)
        _sh = EnsureDispatch("Shell.Application", 0)
        self.filename = _path.name
        self.namespace = _sh.NameSpace(str(_path.parent.absolute()))
        self._get_attributes()

    def get_attribute_dict(self) -> Dict[str, str]:
        """get all attributes and their values"""
        return self.attr_dict

    def get_attribute_list(self) -> List[str]:
        """get list of all file attributes"""
        return list(self.attr_dict.keys())

    def _get_attributes(self) -> None:
        """loads all file metadata and adds it to this object"""
        if _attr_list:
            pass
        else:
            _fill_attr_list(self.namespace)

        attr_dict = {}
        for col in range(len(_attr_list)):
            val = self.namespace.GetDetailsOf(
                self.namespace.ParseName(self.filename), col
            )
            if val:
                attr_dict[_parse_attr_name(_attr_list[col])] = str(val)
                attr_dict[_attr_list[col]] = str(val)
                self.attr_dict[_attr_list[col]] = str(val)

        self.__dict__.update(**attr_dict)

    def __getitem__(self, item: str) -> str:
        """allows dict-like access"""
        return self.__dict__[item]

    def __setitem__(self, key: str, value: str) -> None:
        """allows setting values like with a dict"""
        self.__setattr__(key, value)

    def __setattr__(self, key: str, value: str) -> None:
        """ensures values access dict[like] or attribute.like are kept consistent"""
        if list(str(key))[0].isupper():
            self.__dict__[key] = value
            self.attr_dict[key] = value
        else:
            self.__dict__[_upper_first(_remove_underscore(key))] = value
            self.attr_dict[_upper_first(_remove_underscore(key))] = value
        self.__dict__[key.lower()] = value

# Windows Metadata

This package reads the metadata Windows creates for a file (also called "attributes" or "details"). 
These values are then stored in a WindowsAttribute object and can be accessed like a dict or an object attribute.
```python
attr = WindowsAttributes('testfile')
attr["Horizontal resolution"] # dict-like access
attr.horizontal_resolution    # attribute like access
```

To get a dict of all the attributes a file has
```python
attr.get_attribute_dict()
```
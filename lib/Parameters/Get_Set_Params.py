import Autodesk

#FUNCTION TO GET PARAMETER VALUE  change "AsDouble()" to "AsString()" to change data type.
from Autodesk.Revit.DB import StorageType

def set_parameter_by_name(element, parameter_name, value):
    if element is None or not parameter_name:
        return False, "Invalid input"

    param = element.LookupParameter(parameter_name)
    if param is None:
        return False, "Parameter not found"
    if param.IsReadOnly:
        return False, "Parameter is read-only"

    try:
        param.Set(value)
        return True, None
    except Exception as e:
        return False, str(e)

def get_parameter_value_by_name_AsString(element, parameterName):
    return element.LookupParameter(parameterName).AsString()

def get_parameter_value_by_name_AsInteger(element, parameterName):
    return element.LookupParameter(parameterName).AsInteger()

def get_parameter_value_by_name_AsDouble(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

def get_parameter_value_by_name_AsValueString(element, parameterName):
    return element.LookupParameter(parameterName).AsValueString()
import Autodesk


#FUNCTION TO GET PARAMETER VALUE  change "AsDouble()" to "AsString()" to change data type.
def set_parameter_by_name(element, parameterName, value):
    element.LookupParameter(parameterName).Set(value)

def get_parameter_value_by_name_AsString(element, parameterName):
    return element.LookupParameter(parameterName).AsString()

def get_parameter_value_by_name_AsInteger(element, parameterName):
    return element.LookupParameter(parameterName).AsInteger()

def get_parameter_value_by_name_AsDouble(element, parameterName):
    return element.LookupParameter(parameterName).AsDouble()

def get_parameter_value_by_name_AsValueString(element, parameterName):
    return element.LookupParameter(parameterName).AsValueString()
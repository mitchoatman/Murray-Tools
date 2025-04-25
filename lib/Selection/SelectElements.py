from Autodesk.Revit.UI.Selection import ISelectionFilter

class ISelectionFilter_Classes(ISelectionFilter):
    def __init__(self, allowed_types):
        """ ISelectionFilter made to filter with types
        :param allowed_types: list of allowed Types"""
        self.allowed_types = allowed_types

    def AllowElement(self, element):
        if type(element) in self.allowed_types:
            return True


class ISelectionFilter_Categories(ISelectionFilter):
    def __init__(self, allowed_categories):
        """ ISelectionFilter made to filter with categories
        :param allowed_types: list of allowed Categories"""
        self.allowed_categories = allowed_categories

    def AllowElement(self, element):
        if element.Category.BuiltInCategory in self.allowed_categories:
            return True
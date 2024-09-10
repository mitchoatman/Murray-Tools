# -*- coding: utf-8 -*-
# Created by Vasile Corjan
import System
from System import Enum
from pyrevit import forms, script, DB, revit
from category_lists import category_names_type, category_names_instance, built_in_parameter_group

# Get all categories and built-in parameter groups
app = __revit__.Application
doc = revit.doc
all_categories = [category for category in doc.Settings.Categories]
built_in_parameter_groups = [group for group in System.Enum.GetValues(DB.BuiltInParameterGroup)]
built_in_parameter_group_names = [DB.LabelUtils.GetLabelFor(n) for n in built_in_parameter_groups]

# Get the shared parameter file and its groups
sp_file = app.OpenSharedParameterFile()
sp_groups = sp_file.Groups
dict_pg = {g.Name:g for g in sp_groups}

# Select the parameter group
parameter_groups_dict = forms.SelectFromList.show(sorted(dict_pg),
                                        title='Select the Parameter Group',
                                        multiselect=False,
                                        button_name='Select')
if not parameter_groups_dict:
	script.exit()

# Select the parameter
definitions = dict_pg.get(parameter_groups_dict).Definitions
parameter_definitions_dict = {d.Name:d for d in definitions}
selected_parameters_names = forms.SelectFromList.show(sorted(parameter_definitions_dict),
                                        title='Select the Parameter',
                                        multiselect=True,
                                        button_name='Select')
if not selected_parameters_names:
	script.exit()
    
parameter_definitions = [parameter_definitions_dict.get(s) for s in selected_parameters_names]

# Select the built-in parameter group
selected_parameters = forms.SelectFromList.show(sorted(built_in_parameter_group),
                                        title='Select the parameters',
                                        multiselect=False,
                                        button_name='Select')
                                        
if selected_parameters in built_in_parameter_group_names: 
    parameter = built_in_parameter_groups[built_in_parameter_group_names.index(selected_parameters)]
else:
    None

# Select the binding type (Type or Instance)
bindings = ["Type","Instance"]
selected_binding = forms.SelectFromList.show(bindings,
                                        title='Select the binding type',
                                        multiselect=False,
                                        button_name='Select')
if not selected_binding:
    script.exit()
# Get the category names based on the selected binding type
if selected_binding == "Type":
    category_names = sorted(category_names_type)
else:
    category_names = sorted(category_names_instance)

selected_categories = forms.SelectFromList.show(category_names,
                                              title='Select the category',
                                              multiselect=True,
                                              button_name='Select')
if not selected_categories:
    script.exit()
categories = [c for c in all_categories if c.Name in selected_categories]
built_in_categories = [bic.BuiltInCategory for bic in categories]

# Adding the parameter
with revit.Transaction("Add Parameter"):
    category_set = app.Create.NewCategorySet()
    for bic in built_in_categories:
        category_set.Insert(DB.Category.GetCategory(doc,bic))
    if selected_binding == "Type":
        binding = app.Create.NewTypeBinding(category_set)
    else:
        binding = app.Create.NewInstanceBinding(category_set)
    for param_def in parameter_definitions:
        doc.ParameterBindings.Insert(param_def, binding, parameter)
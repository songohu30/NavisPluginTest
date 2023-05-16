using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge = Autodesk.Navisworks.Api.ComApi;

namespace TestNavisPlugin
{
    //Navisworks "AddInPlugin" requires these atrributes
    [PluginAttribute("AddParameter", "Tesla", DisplayName = "AddParameter", ToolTip = "Add new parameter")]
    public class Program : AddInPlugin //Navis will look for the class implementing AddInPlugin interface (if you duplicate this class here with different attributes and name, it will create another button in addins tab
    {
        public override int Execute(params string[] parameters) //this method is required with AddInPlugin interface
        {
            try
            {
                ModelItem item = GetSelectedItem(); //returns selected item or first from selection;
                if (item != null)
                {
                    AddModifyParameter("TeslaProperties", "Bigger Dimension", "oxox", item); //adds new parameter to selected item or modifies value if exists
                }
                else
                {
                    System.Windows.MessageBox.Show("Please select item!");
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.ToString());
            }


            return 0;
        }

        private ModelItem GetSelectedItem()
        {
            return Application.ActiveDocument.CurrentSelection.SelectedItems.FirstOrDefault();
        }

        //todo: create a single method for repeated code
        private void AddModifyParameter(string parCategory, string parName, string parValue, ModelItem item)
        {
            ComApi.InwOpState10 cdoc = ComApiBridge.ComApiBridge.State;
            ComApi.InwOaPath citem = ComApiBridge.ComApiBridge.ToInwOaPath(item);
            ComApi.InwGUIPropertyNode2 cpropcates = (ComApi.InwGUIPropertyNode2)cdoc.GetGUIPropertyNode(citem, true);
            ComApi.InwOaPropertyVec newcate = (ComApi.InwOaPropertyVec)cdoc.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaPropertyVec);

            bool parameterExists = false;
            bool categoryExists = false;
            int index = 1; //user defined property category index to be found (!starts from 1 not 0)

            foreach (ComApi.InwGUIAttribute2 category in cpropcates.GUIAttributes()) //loop through ModelItem property categories
            {
                if (category.UserDefined)
                {
                    if (category.ClassUserName.Equals(parCategory))
                    {
                        categoryExists = true;
                        bool valueChanged = false;
                        foreach (ComApi.InwOaProperty property in category.Properties()) //loop in user defined category properties
                        {                         
                            if(property.name == parName) //if property with parName exists create a copy with modified value and add to new category object
                            {
                                if (property.value != parValue) valueChanged = true;
                                ComApi.InwOaProperty existingPropertyCopy = (ComApi.InwOaProperty)cdoc.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaProperty);
                                existingPropertyCopy.name = property.name;
                                existingPropertyCopy.UserName = property.name;
                                existingPropertyCopy.value = parValue; //update parameter value;
                                newcate.Properties().Add(existingPropertyCopy); //add a copy of existing property with newValue
                                parameterExists = true;
                            }
                            else
                            {
                                ComApi.InwOaProperty existingPropertyCopy = (ComApi.InwOaProperty)cdoc.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaProperty);
                                existingPropertyCopy.name = property.name;
                                existingPropertyCopy.UserName = property.name;
                                existingPropertyCopy.value = property.value;
                                newcate.Properties().Add(existingPropertyCopy); //add a copy of existing property to new category object
                            }
                        }
                        if (parameterExists)
                        {
                            if(valueChanged) cpropcates.SetUserDefined(index, parCategory, parCategory, newcate); //if value was changed for existing parameter, replace existing category object with a modified copy
                        }
                        else
                        {
                            ComApi.InwOaProperty newprop = (ComApi.InwOaProperty)cdoc.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaProperty);
                            newprop.name = parName;
                            newprop.UserName = parName;
                            newprop.value = parValue;
                            newcate.Properties().Add(newprop);
                            cpropcates.SetUserDefined(index, parCategory, parCategory, newcate); //if category exists and parameter does not exist, add new parameter to current category
                        }
                    }
                    index++;
                }
            }

            if (!parameterExists && !categoryExists)
            {
                ComApi.InwOaProperty newprop = (ComApi.InwOaProperty)cdoc.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaProperty);
                newprop.name = parName;
                newprop.UserName = parName;
                newprop.value = parValue;
                newcate.Properties().Add(newprop);
                cpropcates.SetUserDefined(0, parCategory, parCategory, newcate); //if parameter and category do not exists create a new category with new parameter
            }
        }
    }

    [PluginAttribute("RemoveCategory", "Tesla", DisplayName = "RemoveCategory", ToolTip = "Removes entire parameter category")]
    public class Program2 : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            ModelItem item = GetSelectedItem();
            if (item != null)
            {
                RemoveCategory("SomeCategory", item);
            }
            else
            {
                System.Windows.MessageBox.Show("Please select item!");
            }

            return 0;
        }

        private ModelItem GetSelectedItem()
        {
            return Application.ActiveDocument.CurrentSelection.SelectedItems.FirstOrDefault();
        }

        private void RemoveCategory(string parCategory, ModelItem item)
        {
            ComApi.InwOpState10 cdoc = ComApiBridge.ComApiBridge.State;
            ComApi.InwOaPath citem = ComApiBridge.ComApiBridge.ToInwOaPath(item);
            ComApi.InwGUIPropertyNode2 cpropcates = (ComApi.InwGUIPropertyNode2)cdoc.GetGUIPropertyNode(citem, true);

            int index = 1; //user defined property category index to be found

            foreach (ComApi.InwGUIAttribute2 attribute in cpropcates.GUIAttributes()) //loop in property categories
            {
                if (attribute.UserDefined)
                {
                    if (attribute.ClassUserName.Equals(parCategory))
                    {
                        cpropcates.RemoveUserDefined(index);
                        break;                            
                    }
                    index++;
                }
            }
        }
    }

    [PluginAttribute("RemoveParameter", "Tesla", DisplayName = "RemoveParameter", ToolTip = "Removes parameter")]
    public class Program3 : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            try
            {
                ModelItem item = GetSelectedItem();
                if (item != null)
                {
                    RemoveParameter("TeslaProperties", "Bigger Dimension", item);
                }
                else
                {
                    System.Windows.MessageBox.Show("Please select item!");
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.ToString());
            }

            return 0;
        }

        private ModelItem GetSelectedItem()
        {
            return Application.ActiveDocument.CurrentSelection.SelectedItems.FirstOrDefault();
        }

        private void RemoveParameter(string categoryDisplayName, string parameterDisplayName, ModelItem item)
        {
            ComApi.InwOpState10 documentState = ComApiBridge.ComApiBridge.State; // gets main document state
            ComApi.InwOaPath itemObjectPath = ComApiBridge.ComApiBridge.ToInwOaPath(item); //converts model item to internal object path
            ComApi.InwGUIPropertyNode2 itemPropertyNode = (ComApi.InwGUIPropertyNode2)documentState.GetGUIPropertyNode(itemObjectPath, true); // gets model item property node
            ComApi.InwOaPropertyVec newCategory = (ComApi.InwOaPropertyVec)documentState.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaPropertyVec); // creates new empty property category object

            bool parameterExists = false;
            int categoryIndex = 1;

            foreach (ComApi.InwGUIAttribute2 category in itemPropertyNode.GUIAttributes())
            {
                if (category.UserDefined)
                {
                    if (category.ClassUserName.Equals(categoryDisplayName))
                    {
                        foreach (ComApi.InwOaProperty property in category.Properties())
                        {
                            if (property.name == parameterDisplayName)
                            {
                                parameterExists = true;
                            }
                            else
                            {
                                ComApi.InwOaProperty existingPropertyCopy = CreateProperty(documentState, property.name, property.UserName, property.value);
                                newCategory.Properties().Add(existingPropertyCopy);
                            }
                        }
                        if (parameterExists)
                        {
                            if(newCategory.Properties().Count > 0)
                            {
                                itemPropertyNode.SetUserDefined(categoryIndex, categoryDisplayName, categoryDisplayName, newCategory);
                            }
                            else
                            {
                                itemPropertyNode.RemoveUserDefined(categoryIndex);
                            }
                        }
                    }
                    categoryIndex++;
                }
            }
        }

        private ComApi.InwOaProperty CreateProperty(ComApi.InwOpState10 cdoc, string parName, string userName, string parValue)
        {
            ComApi.InwOaProperty newprop = (ComApi.InwOaProperty)cdoc.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaProperty);
            newprop.name = parName;
            newprop.UserName = userName;
            newprop.value = parValue;
            return newprop;
        }
    }
}

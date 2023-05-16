using ComApi = Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge = Autodesk.Navisworks.Api.ComApi;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Threading;

namespace TestNavisPlugin
{
    [PluginAttribute("SetNewDimension", "Tesla", DisplayName = "Set new dimension", ToolTip = "Add new parameter with biggest dimension")]
    public class TeslaNewDimension : AddInPlugin
    {
        public ViewModel ViewModel { get; set; }
        private ManualResetEvent m_ResetEvent = new ManualResetEvent(false);

        public override int Execute(params string[] parameters)
        {
            Application.ActiveDocument.CurrentSelection.Clear(); //command process is slow when there are selected items!
            ViewModel = new ViewModel();
            Thread windowThread = new Thread(delegate ()
            {
                InfoWindow window = new InfoWindow(ViewModel);
                window.Show();
                m_ResetEvent.Set();
                Dispatcher.Run();
            });
            windowThread.SetApartmentState(ApartmentState.STA);
            windowThread.IsBackground = true;
            windowThread.Start();

            m_ResetEvent.WaitOne();
            m_ResetEvent.Reset();

            try
            {
                string parameterCategory = "Element";
                string parameterName = "Category";
                string parameterValue = "Ducts";

                Search search = new Search();
                search.Selection.SelectAll();
                SearchCondition condition = SearchCondition.HasPropertyByDisplayName(parameterCategory, parameterName);
                search.SearchConditions.Add(condition);
                ModelItemCollection collection = search.FindAll(Application.ActiveDocument, false);
                ViewModel.LogInfo += string.Format("Found {0} items with '{1} {2}' property", collection.Count, parameterCategory, parameterName) + Environment.NewLine;

                List<ModelItem> ducts = collection.Where(e => e.PropertyCategories.FindPropertyByDisplayName(parameterCategory, parameterName).Value.ToDisplayString() == parameterValue).ToList();
                ViewModel.LogInfo += string.Format("Found {0} ducts", ducts.Count) + Environment.NewLine;

                int counter = 0;
                int progressStep = 0;
                foreach (ModelItem duct in ducts)
                {
                    string size = duct.PropertyCategories.FindPropertyByDisplayName("Element", "Size").Value.ToDisplayString();
                    int parsed = GetDuctBiggerDimension(size);
                    AddModifyParameter("Tesla", "Bigger Dimension", parsed, duct);


                    counter++;
                    progressStep++;
                    if (progressStep == 100)
                    {
                        ViewModel.LogInfo += string.Format("Processed {0} items.", counter) + Environment.NewLine;
                        progressStep = 0;
                    }
                }

                ViewModel.LogInfo += "FINISHED!!!";
            }
            catch (Exception ex)
            {
                ViewModel.LogInfo += ex.ToString() + Environment.NewLine;
                ViewModel.LogInfo += "FAILED!";
            }


            return 0;
        }

        //predicted value: DN10xDN25 or ø100 or 100x100
        public int GetDuctBiggerDimension(string size)
        {
            int result = 0;
            try
            {
                int[] values = size.Split('x').Select(e => int.Parse(GetNumbers(e))).ToArray();
                result = values.Max();
            }
            catch (Exception ex)
            {
                ViewModel.LogInfo += ex.ToString() + Environment.NewLine;
            }
            return result;
        }

        private string GetNumbers(string input)
        {
            return new string(input.Where(c => char.IsDigit(c)).ToArray());
        }

        private void AddModifyParameter(string categoryDisplayName, string parameterDisplayName, int parameterValue, ModelItem item)
        {
            ComApi.InwOpState10 documentState = ComApiBridge.ComApiBridge.State; // gets main document state
            ComApi.InwOaPath itemObjectPath = ComApiBridge.ComApiBridge.ToInwOaPath(item); //converts model item to internal object path
            ComApi.InwGUIPropertyNode2 itemPropertyNode = (ComApi.InwGUIPropertyNode2)documentState.GetGUIPropertyNode(itemObjectPath, true); // gets model item property node
            ComApi.InwOaPropertyVec newCategory = (ComApi.InwOaPropertyVec)documentState.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaPropertyVec); // creates new empty property category object

            bool parameterExists = false;
            bool categoryExists = false;
            int index = 1; //user defined property category index to be found (!starts from 1 not 0)

            foreach (ComApi.InwGUIAttribute2 category in itemPropertyNode.GUIAttributes()) //loop through ModelItem property categories
            {
                if (category.UserDefined)
                {
                    if (category.ClassUserName.Equals(categoryDisplayName))
                    {
                        categoryExists = true;
                        bool valueChanged = false;
                        foreach (ComApi.InwOaProperty property in category.Properties()) //loop in user defined category properties
                        {
                            if (property.name == parameterDisplayName) //if property with parName exists create a copy with modified value and add to new category object
                            {
                                if (property.value != parameterValue) valueChanged = true;
                                ComApi.InwOaProperty existingPropertyCopy = CreateProperty(documentState, property.name, property.UserName, parameterValue);  //update parameter value;
                                newCategory.Properties().Add(existingPropertyCopy); //add a copy of existing property with newValue
                                parameterExists = true;
                            }
                            else
                            {
                                ComApi.InwOaProperty existingPropertyCopy = CreateProperty(documentState, property.name, property.UserName, property.value);
                                newCategory.Properties().Add(existingPropertyCopy); //add a copy of existing property to new category object
                            }
                        }
                        if (parameterExists)
                        {
                            if (valueChanged) itemPropertyNode.SetUserDefined(index, categoryDisplayName, categoryDisplayName, newCategory); //if value was changed for existing parameter, replace existing category object with a modified copy
                        }
                        else
                        {
                            ComApi.InwOaProperty newprop = CreateProperty(documentState, parameterDisplayName, parameterDisplayName, parameterValue);
                            newCategory.Properties().Add(newprop);
                            itemPropertyNode.SetUserDefined(index, categoryDisplayName, categoryDisplayName, newCategory); //if category exists and parameter does not exist, add new parameter to current category
                        }
                    }
                    index++;
                }
            }

            if (!parameterExists && !categoryExists)
            {
                ComApi.InwOaProperty newprop = CreateProperty(documentState, parameterDisplayName, parameterDisplayName, parameterValue);
                newCategory.Properties().Add(newprop);
                itemPropertyNode.SetUserDefined(0, categoryDisplayName, categoryDisplayName, newCategory); //if parameter and category do not exists create a new category with new parameter
            }
        }

        private ComApi.InwOaProperty CreateProperty(ComApi.InwOpState10 cdoc, string parName, string userName, int parValue)
        {
            ComApi.InwOaProperty newprop = (ComApi.InwOaProperty)cdoc.ObjectFactory(ComApi.nwEObjectType.eObjectType_nwOaProperty);
            newprop.name = parName;
            newprop.UserName = userName;
            newprop.value = parValue;
            return newprop;
        }
    }
}

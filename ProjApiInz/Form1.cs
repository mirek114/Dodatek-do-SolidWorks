using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SldWorks;
using SWPublished;
using SwConst;
using SwCommands;

//using SolidWorksTools;
//using SolidWorksTools.File;


using System.Runtime.InteropServices;

namespace ProjApiInz
{
    public partial class Form1 : Form
    {
        #region Variables
        SldWorks.SldWorks swApp;
        ModelDoc2 swModel;


        SelectionMgr swSelMgr;

        public DispatchWrapper[] bodyArray;

        //kolorki
        DisplayStateSetting swDisplayStateSetting = null;
        string[] displayStateNames = new string[1];

        object appearances = null;
        object[] appearancesArray = null;

        AppearanceSetting swAppearanceSetting = default(AppearanceSetting);
        AppearanceSetting[] newAppearanceSetting = new AppearanceSetting[1];

        int red_rgb = 0;
        int green_rgb = 0;
        int blue_rgb = 0;
        int newColor = 0;

        double minVolume = 0.0;
        double maxVolume = 0.0;

        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            #region Min/Max Volume

            if (textBox1 != null && textBox2 != null)
            {
                minVolume = Convert.ToDouble(textBox1.Text);
                maxVolume = Convert.ToDouble(textBox2.Text);
            }
            else
            {
                MessageBox.Show("Musisz podać zakres objętości!");
                return;
            }
            #endregion

            #region SetColor
            if (comboBox1.SelectedItem.ToString() != "")
            {
                SetNewColor(comboBox1.SelectedItem.ToString());
            }
            else
            {
                MessageBox.Show("Musisz wybrać kolor!");
                return;
            }
            #endregion

            #region SWConnection
            try
            {
                swApp = (SldWorks.SldWorks)Marshal.GetActiveObject("SldWorks.Application");

            }
            catch
            {
                MessageBox.Show("Błąd Połączenia z SW!");
                return;
            }
            #endregion

            #region DocumentModel
            swModel = (ModelDoc2)swApp.ActiveDoc;

            if (swModel == null)
            {
                MessageBox.Show("Bład przy otwieraniu");
                return;
            }

            if (swModel.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                MessageBox.Show("Otwarty dokument nie jest złożeniem!");
                return;
            }
            #endregion


            if (checkBox1.Checked)
            {
                AnalyzeSelectedParts();
            }
            else
            {
                AnalyzeAllParts();
            }

            //swModel.ClearSelection2(true);
            swModel.EditRebuild3();
        }

        private void AnalyzeAllParts()
        {
            Configuration swConf = (Configuration)swModel.GetActiveConfiguration();
            Component2 swRootComponent = (Component2)swConf.GetRootComponent();
            TraverseComponent(swRootComponent);

        }

        private void AnalyzeSelectedParts()
        {
            swSelMgr = (SelectionMgr)swModel.SelectionManager;

            for (int i = 1; i <= swSelMgr.GetSelectedObjectCount2(-1); i++)
            {
                var swComponent = (Component2)swSelMgr.GetSelectedObject6(i, 0);
                Component2[] swComponents = new Component2[1];
                swComponents[0] = (Component2)swComponent;
                if (GetVolume(swComponent) >= minVolume && GetVolume(swComponent) <= maxVolume)
                {
                    ChangeColor(swComponents);
                }

            }
        }

        private void TraverseComponent(Component2 component)
        {
            object[] children = (object[])component.GetChildren();
            if (children.Length > 0)
            {
                foreach (Component2 comp in children)
                {
                    TraverseComponent(comp);

                    Component2[] swComponents = new Component2[1];
                    swComponents[0] = (Component2)comp;
                    if (GetVolume(comp) >= minVolume && GetVolume(comp) <= maxVolume)
                    {
                        ChangeColor(swComponents);
                    }
                }
            }
        }

        private void ChangeColor(Component2[] swComponents)
        {
             ModelDocExtension swModelDocExt = (ModelDocExtension)swModel.Extension;
            //Get display state
            swDisplayStateSetting = swModelDocExt.GetDisplayStateSetting((int)swDisplayStateOpts_e.swAllDisplayState);
            swDisplayStateSetting.Entities = swComponents;
            swDisplayStateSetting.Option = (int)swDisplayStateOpts_e.swSpecifyDisplayState;
            displayStateNames[0] = "<Default>_Display State 1";
            swDisplayStateSetting.Names = displayStateNames;
            //Set to part level
            swDisplayStateSetting.PartLevel = true;

            appearances = (object)swModelDocExt.get_DisplayStateSpecMaterialPropertyValues(swDisplayStateSetting);
            appearancesArray = (object[])appearances;
            swAppearanceSetting = (AppearanceSetting)appearancesArray[0];
            
            swAppearanceSetting.Color = newColor;
            newAppearanceSetting[0] = swAppearanceSetting;
            swModelDocExt.set_DisplayStateSpecMaterialPropertyValues(swDisplayStateSetting, newAppearanceSetting);
        }

        private double GetVolume(Component2 swComponent)
        {
            object bodyInfo = null;

            object[] bodies = (object[])swComponent.GetBodies3((int)swBodyType_e.swAllBodies, out bodyInfo);

            MassProperty swMass = (MassProperty)swModel.Extension.CreateMassProperty();

            bodyArray = (DispatchWrapper[])ObjectArrayToDispatchWrapperArray(bodies);

            bool boolstatus = swMass.AddBodies((bodyArray));
            swMass.UseSystemUnits = false;

            return swMass.Volume;
        }

        private DispatchWrapper[] ObjectArrayToDispatchWrapperArray(object[] SwObjects)
        {
            int arraySize = 0;
            arraySize = SwObjects.GetUpperBound(0);
            DispatchWrapper[] dispwrap = new DispatchWrapper[arraySize + 1];
            int arrayIndex = 0;
            for (arrayIndex = 0; arrayIndex <= arraySize; arrayIndex++)
            {
                dispwrap[arrayIndex] = new DispatchWrapper(SwObjects[arrayIndex]);
            }
            return dispwrap;
        }

        private void SetNewColor(string colorName)
        {
            switch (colorName)
            {
                case "Czerwony":
                    red_rgb = 255;
                    green_rgb = 0;
                    blue_rgb = 0;
                    break;
                case "Zielony":
                    red_rgb = 0;
                    green_rgb = 255;
                    blue_rgb = 0;
                    break;
                case "Niebieski":
                    red_rgb = 0;
                    green_rgb = 0;
                    blue_rgb = 255;
                    break;
                case "Żółty":
                    red_rgb = 255;
                    green_rgb = 255;
                    blue_rgb = 0;
                    break;
                case "Różowy":
                    red_rgb = 255;
                    green_rgb = 0;
                    blue_rgb = 255;
                    break;
                case "Czarny":
                    red_rgb = 0;
                    green_rgb = 0;
                    blue_rgb = 0;
                    break;
            }
            newColor = Math.Max(Math.Min(red_rgb, 255), 0) + Math.Max(Math.Min(green_rgb, 255), 0) * 16 * 16 + Math.Max(Math.Min(blue_rgb, 255), 0) * 16 * 16 * 16 * 16;
        }

    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Display;

namespace DesktopWindowsApplication2
{
    public partial class Form1 : Form
    {
      
        int CheckLoading = 0;
        string Path;
        string DBPath;
        int DF;
        IMapDocument Map = new MapDocument();
        Dictionary<string, IMapDocument> Dic = new Dictionary<string, IMapDocument>();
        DataTable table = new DataTable();
        Dictionary<string, List<string>> LayerCond = new Dictionary<string, List<string>>();
        OpenFileDialog openFileDialog1;
        List<string> ListofUniqueValues = new List<string>();
        string LastCondition;
        public Form1()
        {
            InitializeComponent();
        }

        public void ChangeDataFrame(int DFIndex)
        {
            int cv = 0;
            axMapControl1.Map = Map.get_Map(DFIndex);

            axMapControl1.Refresh();

            checkedListBox1.Items.Clear();
            comboBox3.Items.Clear();
            comboBox2.Items.Clear();
            comboBox8.Items.Clear();

            IEnumLayer AllLayers = axMapControl1.Map.Layers;
            ILayer OneLayer = AllLayers.Next();
            while (OneLayer != null)
            {
                checkedListBox1.Items.Add(OneLayer.Name);
                comboBox3.Items.Add(OneLayer.Name);
                comboBox2.Items.Add(OneLayer.Name);
                comboBox8.Items.Add(OneLayer.Name);
                comboBox13.Items.Add(OneLayer.Name);
                comboBox15.Items.Add(OneLayer.Name);
                comboBox18.Items.Add(OneLayer.Name);
                comboBox22.Items.Add(OneLayer.Name);
                comboBox24.Items.Add(OneLayer.Name);
                comboBox28.Items.Add(OneLayer.Name);
                if (Map.get_Map(DFIndex).get_Layer(cv).Visible == true)
                {
                    checkedListBox1.SetItemChecked(cv, true);
                }
                cv++;
                OneLayer = AllLayers.Next();
            }
            comboBox3.Text = comboBox3.Items[0].ToString();
            comboBox2.Text = comboBox2.Items[0].ToString();
            comboBox8.Text = comboBox2.Items[0].ToString();
            comboBox13.Text = comboBox13.Items[0].ToString();
            comboBox15.Text = comboBox15.Items[0].ToString();
            comboBox18.Text = comboBox18.Items[0].ToString();
            comboBox22.Text = comboBox22.Items[0].ToString();
            comboBox24.Text = comboBox24.Items[0].ToString();
            comboBox28.Text = comboBox28.Items[0].ToString();
            DF = comboBox1.SelectedIndex;
            Dic[Path] = Map;
        }

        public void HideLayer(int LayerIndex)
        {
            try
            {
                if (Map.get_Map(DF).get_Layer(LayerIndex).Visible == true)
                {
                    axMapControl1.Map.get_Layer(LayerIndex).Visible = false;
                    axMapControl1.Refresh();
                    Dic[Path] = Map;
                }
                else
                {
                    axMapControl1.Map.get_Layer(LayerIndex).Visible = true;
                    axMapControl1.Refresh();
                    Dic[Path] = Map;
                    
                }
            }
            catch
            { }
        }

        public bool LoadMap()
        {
            try
            {
                openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Title = "Browse Mxd Files";
                openFileDialog1.DefaultExt = "mxd";
                openFileDialog1.Filter = "mxd Diles (*.mxd)|*.mxd";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    Path = openFileDialog1.FileName;
                    if (Dic.ContainsKey(Path) == true)
                    {
                        checkedListBox1.Items.Clear();
                        Map = Dic[Path];
                        axMapControl1.Map = Map.get_Map(0);
                        axMapControl1.Refresh();
                    }
                    else
                    {
                        checkedListBox1.Items.Clear();
                        Map.Open(Path);
                        axMapControl1.Map = Map.get_Map(0);
                        axMapControl1.Refresh();
                        Dic.Add(Path, Map);

                    }
                    comboBox1.Items.Clear();
                    for (int i = 0; i < Map.MapCount; i++)
                    {
                        comboBox1.Items.Add(Map.get_Map(i).Name);
                    }
                    comboBox1.Text = comboBox1.Items[0].ToString();

                    toolStripStatusLabel2.Text = "Scale: " + axMapControl1.MapScale.ToString();
                    
                    groupBox1.Enabled = true;
                    groupBox2.Enabled = true;
                    groupBox3.Enabled = true;
                    groupBox6.Enabled = true;
                    groupBox7.Enabled = true;
                    groupBox8.Enabled = true;
                    groupBox4.Enabled = true;
                    groupBox5.Enabled = true;

                    return true;
                }

            }
            catch
            {
                return false;
            }
            return true;

        }

        public void SetInitialValues()
        {
            toolStripStatusLabel1.Text = "Map X Y";
            toolStripStatusLabel2.Text = "Map Scale";
            groupBox1.Enabled = false;
            groupBox2.Enabled = false;
            groupBox3.Enabled = false;
            groupBox6.Enabled = false;
            groupBox7.Enabled = false;
            groupBox8.Enabled = false;
            groupBox4.Enabled = false;
            groupBox5.Enabled = false;
            comboBox5.Text = comboBox5.Items[0].ToString();
            comboBox6.Text = comboBox6.Items[0].ToString();
            comboBox11.Text = comboBox11.Items[0].ToString();
            comboBox26.Text = comboBox26.Items[0].ToString();

        }

        public void LoadDB()
        {
            try
            {
                openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Title = "Browse mdb Files";
                openFileDialog1.DefaultExt = "mdb";
                openFileDialog1.Filter = "mdb Diles (*.mdb)|*.mdb";
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    DBPath = openFileDialog1.FileName; ;
                }
               
                IWorkspaceFactory AccessWorkSpace = new AccessWorkspaceFactoryClass();
                IWorkspace MyWorkSpace = AccessWorkSpace.OpenFromFile(DBPath, 0);
                IFeatureWorkspace featWorkspace = MyWorkSpace as IFeatureWorkspace;

                IEnumDataset AllLayerFromDB = MyWorkSpace.get_Datasets(esriDatasetType.esriDTFeatureClass);
                IDataset SingleLayer = AllLayerFromDB.Next();
                while (SingleLayer != null)
                {
                    string LayerText = SingleLayer.Name;
                    comboBox12.Items.Add(LayerText);
                    SingleLayer = AllLayerFromDB.Next();
                }
                comboBox12.Text = comboBox12.Items[0].ToString();
            }
            catch
            {
            }
        }

        public void AddLayer(string LayerName)
        {
            IMap FocusedMap = axMapControl1.ActiveView.FocusMap;
            IWorkspaceFactory AccessWorkSpace = new AccessWorkspaceFactoryClass();
            IWorkspace MyWorkSpace = AccessWorkSpace.OpenFromFile(DBPath, 0);
            IFeatureWorkspace FeatureOfWorkSpace = MyWorkSpace as IFeatureWorkspace;
            IFeatureClass FeatureClassOfLayer;


            try
            {

                FeatureClassOfLayer = FeatureOfWorkSpace.OpenFeatureClass(LayerName);
                IFeatureLayer LayerToBeAdded = new FeatureLayerClass();
                LayerToBeAdded.FeatureClass = FeatureClassOfLayer;
                LayerToBeAdded.Name = LayerName;
                LayerToBeAdded.ShowTips = true;
                FocusedMap.AddLayer(LayerToBeAdded);
                FocusedMap.MoveLayer(LayerToBeAdded, 2);



                if (!checkedListBox1.Items.Contains(LayerName))
                {
                    checkedListBox1.Items.Add(LayerName);
                    checkedListBox1.SetItemChecked(checkedListBox1.Items.Count - 1, true);
                    comboBox3.Items.Add(LayerName);
                    comboBox2.Items.Add(LayerName);
                    comboBox8.Items.Add(LayerName);

                }
                
            }
            catch
            {
                
            }
            
            
        }

        public void RemoveLayer(int LayerIndex,string LayerName)
        {
            try
            {

                axMapControl1.DeleteLayer(LayerIndex);
                axMapControl1.Refresh();

                checkedListBox1.Items.Remove(LayerName);
                comboBox8.Items.Remove(LayerName);
                comboBox8.Text = comboBox8.Items[0].ToString();
                comboBox2.Items.Remove(LayerName);
                comboBox2.Text = comboBox2.Items[0].ToString();
                comboBox3.Items.Remove(LayerName);
                comboBox3.Text = comboBox3.Items[0].ToString();
            }
            catch
            { }
        }

        public void ApplyDefinisionQuery(string Field, string Condition, string Value)
        {
            try
            {

                IFeatureLayer FocusedLayer = axMapControl1.ActiveView.FocusMap.get_Layer(comboBox2.SelectedIndex) as IFeatureLayer;

                IFeatureLayerDefinition DefinisionQuery = FocusedLayer as IFeatureLayerDefinition;
                string ConditionToBeApplyed;
                char FirstCharOfValue = comboBox9.Text[0];
                int FirstNumAsci=48;
                int SecondNumAsci=57;
                if (FirstCharOfValue >= FirstNumAsci && FirstCharOfValue <= SecondNumAsci)
                {
                    ConditionToBeApplyed = Field + Condition + Value;
                }
                else
                {
                    string ValueinQoutes = "'" + Value + "'";
                    ConditionToBeApplyed = Field + Condition + ValueinQoutes;
                }
                DefinisionQuery.DefinitionExpression = ConditionToBeApplyed;
                axMapControl1.ActiveView.Refresh();
            }
            catch
            {
                MessageBox.Show("Error");
            }
        }

        public void RemoveDefinisionQuery(int LayerIndex)
        {
            try
            {

                IFeatureLayer FocusedLayer = axMapControl1.ActiveView.FocusMap.get_Layer(LayerIndex) as IFeatureLayer;
                IFeatureLayerDefinition DefinisionQuery = FocusedLayer as IFeatureLayerDefinition;

                DefinisionQuery.DefinitionExpression = "";
                axMapControl1.ActiveView.Refresh();
            }
            catch
            {

            }
        }

        public void SelectByAttribute(string Field, string Condition, string Value,string Method)
        {


            IFeatureLayer FoucsedLayer = axMapControl1.ActiveView.FocusMap.get_Layer(comboBox8.SelectedIndex) as IFeatureLayer;
            IFeatureCursor CursorOfFeatures;
            IQueryFilter FilterOfSelection = new QueryFilterClass();
            string SelectCondition;
            int FirstNumberAsci = 48;
            int LastNumberAsci = 57;
            char FirstCharofValue = comboBox10.Text[0];
            string EmptyS = "Empty";
            if (FirstCharofValue >= FirstNumberAsci && FirstCharofValue <= LastNumberAsci)
            {

                SelectCondition = Field + Condition + Value;
                LastCondition = SelectCondition;

            }
            else
            {
                string ValueInQoutes = "'" + Value + "'";
                SelectCondition = Field + Condition + ValueInQoutes;
                LastCondition = SelectCondition;
            }

            if (LayerCond.ContainsKey(FoucsedLayer.Name))
            {
                LayerCond[FoucsedLayer.Name].Add(SelectCondition);
            }
            else
            {
                List<string> Tmp = new List<string>();
                Tmp.Add(SelectCondition);
                LayerCond.Add(FoucsedLayer.Name, Tmp);
            }


            FilterOfSelection.WhereClause = SelectCondition;
            CursorOfFeatures = FoucsedLayer.FeatureClass.Search(FilterOfSelection, true);
            IFeatureSelection featSelection = FoucsedLayer as IFeatureSelection;
            if (Method == "Add")
            {
                featSelection.SelectFeatures(FilterOfSelection, esriSelectionResultEnum.esriSelectionResultAdd, false);
                IFeatureCursor CursorForGridView =
                FoucsedLayer.FeatureClass.Search(FilterOfSelection, true);

                IFeature feature = CursorForGridView.NextFeature();

                int NumOfFieldsinLayer=FoucsedLayer.FeatureClass.Fields.FieldCount;
                string[] InsertedArr = new string[NumOfFieldsinLayer];
                int NumOfFieldsinGV = CursorForGridView.Fields.FieldCount;
                int FirstLoopCounter = 0;
                int SecondLoopCounter = 0;
                while(FirstLoopCounter<NumOfFieldsinGV)      
                {

                    if (feature == null)
                        break;
                    SecondLoopCounter = 0;
                    
                    while(SecondLoopCounter<NumOfFieldsinLayer)    
                    {
                        string CurFeature = feature.get_Value(SecondLoopCounter).ToString();
                        if (CurFeature == "")
                        {
                            InsertedArr[SecondLoopCounter] = EmptyS;
                        }
                        else
                        {
                            InsertedArr[SecondLoopCounter] = CurFeature;
                        }
                        SecondLoopCounter = SecondLoopCounter + 1;

                    }


                    table.Rows.Add(InsertedArr);

                    feature = CursorForGridView.NextFeature();
                    FirstLoopCounter = FirstLoopCounter + 1;

                }

                dataGridView1.DataSource = table;

            }
            else if (Method == "New")
            {
                dataGridView1.DataSource = null;
                table.Clear();
                featSelection.SelectFeatures(FilterOfSelection, esriSelectionResultEnum.esriSelectionResultNew, false);
                IFeatureCursor CursorForGridView =
                FoucsedLayer.FeatureClass.Search(FilterOfSelection, true);
                IFeature feature = CursorForGridView.NextFeature();
                dataGridView1.DataSource = table;
                int NumOfFieldsinLayer = FoucsedLayer.FeatureClass.Fields.FieldCount;
                string[] InsertedArr = new string[NumOfFieldsinLayer];
                int NumOfFieldsinGV = CursorForGridView.Fields.FieldCount;
                int FirstLoopCounter = 0;
                int SecondLoopCounter = 0;
                while (FirstLoopCounter < NumOfFieldsinGV)
                {

                    if (feature == null)
                        break;
                    SecondLoopCounter = 0;
                   
                    while (SecondLoopCounter < NumOfFieldsinLayer)
                    {
                        string CurFeature = feature.get_Value(SecondLoopCounter).ToString();
                        if (CurFeature == "")
                        {
                            InsertedArr[SecondLoopCounter] = EmptyS;
                        }
                        else
                        {
                            InsertedArr[SecondLoopCounter] = CurFeature;
                        }
                        SecondLoopCounter = SecondLoopCounter + 1;

                    }


                    table.Rows.Add(InsertedArr);

                    feature = CursorForGridView.NextFeature();
                    FirstLoopCounter = FirstLoopCounter + 1;

                }

                dataGridView1.DataSource = table;
            }
            else if (Method == "Remove")
            {
                try
                {
                    featSelection.SelectFeatures(FilterOfSelection, esriSelectionResultEnum.esriSelectionResultSubtract, false);

                    IFeatureCursor CursorForGridView =
                    FoucsedLayer.FeatureClass.Search(FilterOfSelection, true);
                    IFeature feature = CursorForGridView.NextFeature();
                    DataRow[] Rows;
                    Rows = table.Select(SelectCondition);
                    foreach (DataRow r in Rows)
                        r.Delete();





                }
                catch
                { }
            }

            axMapControl1.ActiveView.Refresh();

        }

        public void GetFieldsOfLayer(int IndexOfLayer,ComboBox CB)
        {
            CB.Items.Clear();
            IFeatureLayer FocusedLayer = axMapControl1.ActiveView.FocusMap.get_Layer(IndexOfLayer) as IFeatureLayer;
            int LayerFields = int.Parse(FocusedLayer.FeatureClass.Fields.FieldCount.ToString());
            int FirstLoopCounter = 0;
            while (FirstLoopCounter < LayerFields)
            {
                string FieldValue = FocusedLayer.FeatureClass.Fields.get_Field(FirstLoopCounter).Name;
                CB.Items.Add(FieldValue);
                FirstLoopCounter = FirstLoopCounter + 1;
            }
            CB.Text = CB.Items[0].ToString();
        }

        public void WriteHeadersInGridView(int LayerIndex,DataTable Table)
        {
            IFeatureLayer FocusedLayer = axMapControl1.ActiveView.FocusMap.get_Layer(LayerIndex) as IFeatureLayer;
            Table.Rows.Clear();
            Table.Columns.Clear();
            int NumOfFieldsinLayer=FocusedLayer.FeatureClass.Fields.FieldCount;
            int FirstLoopCounter = 0;
            while(FirstLoopCounter<NumOfFieldsinLayer)
            {
                string FieldValue=FocusedLayer.FeatureClass.Fields.get_Field(FirstLoopCounter).Name;
                Table.Columns.Add(FieldValue, typeof(string));
                FirstLoopCounter = FirstLoopCounter + 1;
            }
        }

        public void GetUniqueValuesOfFiled(int LayerIndex, int FieldIndex, ComboBox CB)
        {
            CB.Items.Clear();
            IFeatureLayer FocusedLayer = axMapControl1.ActiveView.FocusMap.get_Layer(LayerIndex) as IFeatureLayer;
            IQueryFilter qfilter = new QueryFilterClass();

            qfilter.WhereClause = "";

            IFeatureCursor featCursor = FocusedLayer.FeatureClass.Search(qfilter, true);
            IFeature FeatureByFeature = featCursor.NextFeature();

            while (FeatureByFeature != null)
            {
                string ValueOfFeature = FeatureByFeature.get_Value(FieldIndex).ToString();
                if (!CB.Items.Contains(ValueOfFeature))
                {
                    CB.Items.Add(ValueOfFeature);
                }
                FeatureByFeature = featCursor.NextFeature();
            }
            CB.Text = CB.Items[0].ToString();
        }

        public void SetXY(string X, string Y,int CheckIfMapLoaded)
        {

            if (CheckIfMapLoaded == 0)
            {
                toolStripStatusLabel1.Text = "Map X Y";
            }
            else
            toolStripStatusLabel1.Text = "X=" + X + "," + "Y=" + X;
        }

        public void ClearSelection(string Field, string Condition, string Value)
        {
            try
            {

                IFeatureLayer FocusedLayer = axMapControl1.ActiveView.FocusMap.get_Layer(comboBox8.SelectedIndex) as IFeatureLayer;
                IFeatureCursor featCursor;
                IQueryFilter qFilter = new QueryFilterClass();
                string ConditionOfSelection;
                char FirstCharinValue=comboBox10.Text[0];
                int FirstCharAsci = 48;
                int SecondCharAsci = 57;
                if (FirstCharinValue >= FirstCharAsci && FirstCharinValue <= SecondCharAsci)
                {
                    ConditionOfSelection = Field + Condition + Value;
                }
                else
                {
                    string ValueinQoutes="'" + Value + "'";
                    ConditionOfSelection = Field + Condition + ValueinQoutes;
                }
                qFilter.WhereClause = ConditionOfSelection;
                featCursor = FocusedLayer.FeatureClass.Search(qFilter, true);
                IFeatureSelection featSelection = FocusedLayer as IFeatureSelection;
                featSelection.SelectFeatures(qFilter, esriSelectionResultEnum.esriSelectionResultNew, false);
                featSelection.Clear();
                dataGridView1.DataSource = null;
                table.Clear();
                axMapControl1.ActiveView.Refresh();
            }
            catch
            {

            }

        }

        public void SelectPoint(string X,string Y,int CheckIfMapLoaded)
        {
            if (CheckIfMapLoaded == 0)
            {
                toolStripStatusLabel1.Text = "Map X Y";
            }
            else
            {
                toolStripStatusLabel1.Text = "X=" + X + "," + "Y=" + Y;

                int ConstantNum = 255;
                IMarkerElement MyIMarkerElement = new MarkerElementClass();
                ISimpleMarkerSymbol MyISimpleMarkerSymbol = new SimpleMarkerSymbolClass();
                IRgbColor MyIRGBColor = new RgbColorClass();
                MyIRGBColor.Blue = ConstantNum;
                MyISimpleMarkerSymbol.Color = MyIRGBColor;
                MyISimpleMarkerSymbol.Style = esriSimpleMarkerStyle.esriSMSCross;
                MyIMarkerElement.Symbol = MyISimpleMarkerSymbol;
                IPoint MyIPoint = new PointClass();
                MyIPoint.X = double.Parse(X);
                MyIPoint.Y = double.Parse(Y);
                IElement MyIElement = MyIMarkerElement as IElement;
                MyIElement.Geometry = MyIPoint;
                axMapControl1.ActiveView.GraphicsContainer.AddElement(MyIElement, 0);
                axMapControl1.ActiveView.Refresh();

                ISpatialFilter MyISpatialFilter = new SpatialFilterClass();
                MyISpatialFilter.Geometry = MyIPoint;
                MyISpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelWithin;
                IFeatureLayer SelectedLayer = axMapControl1.ActiveView.FocusMap.get_Layer(comboBox13.SelectedIndex) as IFeatureLayer;
                IFeatureSelection MyIFeatureSelection = SelectedLayer as IFeatureSelection;
                MyIFeatureSelection.SelectFeatures(MyISpatialFilter, esriSelectionResultEnum.esriSelectionResultNew, false);
                axMapControl1.ActiveView.Refresh();


            }
        }

        public void CategoryUniqueValues(IGeoFeatureLayer geoLayer, string aFieldName)
        {
            IUniqueValueRenderer MyIUniqueValueRenderer;
            MyIUniqueValueRenderer = new UniqueValueRenderer();
            MyIUniqueValueRenderer.FieldCount = 1;
            MyIUniqueValueRenderer.set_Field(0, aFieldName);
            int intFieldIndex;
            intFieldIndex = geoLayer.FeatureClass.FindField(aFieldName);
            IFeatureCursor MyIFeatureCursor;
            IQueryFilter MyQueryFilter;
            MyQueryFilter = new QueryFilter();
            MyQueryFilter.SubFields = aFieldName;
            MyIFeatureCursor = geoLayer.FeatureClass.Search(MyQueryFilter, true);
            IRandomColorRamp MyIRandomColorRamp;
            MyIRandomColorRamp = new RandomColorRamp();
            MyIRandomColorRamp.Size = 16;
            bool CheckColors;
            MyIRandomColorRamp.CreateRamp(out CheckColors);
            if (!CheckColors)
            {
                return;
            }
            IEnumColors MyIEnumColors= MyIRandomColorRamp.Colors;
            ISimpleMarkerSymbol MyISimpleMarkerSymbol;
            IColor MyIColor;
            IFeature MyIFeature;
            MyIFeature = MyIFeatureCursor.NextFeature();
            while (MyIFeature != null)
            {
                MyISimpleMarkerSymbol = new SimpleMarkerSymbolClass();
                MyIColor = MyIEnumColors.Next();
                if ((MyIColor == null))
                {
                    MyIEnumColors.Reset();
                    MyIColor = MyIEnumColors.Next();
                }
                MyISimpleMarkerSymbol.Style = esriSimpleMarkerStyle.esriSMSCross;
                MyISimpleMarkerSymbol.Color = MyIColor;
                MyISimpleMarkerSymbol.Size = 22;
                MyIUniqueValueRenderer.AddValue(MyIFeature.get_Value(intFieldIndex).ToString(), "", ((ISymbol)(MyISimpleMarkerSymbol)));
                MyIFeature = MyIFeatureCursor.NextFeature();
            }
            geoLayer.Renderer = MyIUniqueValueRenderer as IFeatureRenderer;
            axMapControl1.ActiveView.Refresh();
        }

        public void UniqueValueMultiCondition(IGeoFeatureLayer geoLayer, string aFieldName,string Value)
        {
            IUniqueValueRenderer MyIUniqueValueRenderer;
            MyIUniqueValueRenderer = new UniqueValueRenderer();
            MyIUniqueValueRenderer.FieldCount = 1;
            MyIUniqueValueRenderer.set_Field(0, aFieldName);
            int intFieldIndex;
            intFieldIndex = geoLayer.FeatureClass.FindField(aFieldName);
            IFeatureCursor MyIFeatureCursor;
            IQueryFilter MyQueryFilter;
            MyQueryFilter = new QueryFilter();
            int ValuesCount = ListofUniqueValues.Count;
           
            for(int i=0;i<ValuesCount;i++)
            {

                MyQueryFilter.WhereClause = ListofUniqueValues[i];
                MyIFeatureCursor = geoLayer.FeatureClass.Search(MyQueryFilter, true);
            IRandomColorRamp MyIRandomColorRamp;
            MyIRandomColorRamp = new RandomColorRamp();
            MyIRandomColorRamp.Size = 16;
            bool CheckColor;
            MyIRandomColorRamp.CreateRamp(out CheckColor);
            if (!CheckColor)
            {
                return;
            }
            IEnumColors MyIEnumColors;
            MyIEnumColors = MyIRandomColorRamp.Colors;
            ISimpleMarkerSymbol MyISimpleMarkerSymbol;
            IColor MyIColor;
            IFeature MyIFeature=MyIFeatureCursor.NextFeature();
            while (MyIFeature != null)
            {
                MyISimpleMarkerSymbol = new SimpleMarkerSymbolClass();
                MyIColor = MyIEnumColors.Next();
                if ((MyIColor == null))
                {
                    MyIEnumColors.Reset();
                    MyIColor = MyIEnumColors.Next();
                }
                MyISimpleMarkerSymbol.Style = esriSimpleMarkerStyle.esriSMSCross;
                MyISimpleMarkerSymbol.Color = MyIColor;
                MyISimpleMarkerSymbol.Size = 22;
                MyIUniqueValueRenderer.AddValue(MyIFeature.get_Value(intFieldIndex).ToString(), "", ((ISymbol)(MyISimpleMarkerSymbol)));
                MyIFeature = MyIFeatureCursor.NextFeature();
            }
            geoLayer.Renderer = MyIUniqueValueRenderer as IFeatureRenderer;
            axMapControl1.ActiveView.Refresh();
        }
        }

        public void SelectByPolygon(int IndexOfLayer,RubberPolygonClass Polygon)
        {
            int ConstantNum = 255;
            IActiveView CurrentView = axMapControl1.ActiveView;
            IScreenDisplay MyScreenDispaly = CurrentView.ScreenDisplay;
            MyScreenDispaly.StartDrawing(MyScreenDispaly.hDC, (System.Int16)esriScreenCache.esriNoScreenCache); 
            IRgbColor MYRGBCOLOR = new RgbColorClass();
            MYRGBCOLOR.Red = ConstantNum;
            IColor MyColor = MYRGBCOLOR; 
            ISimpleFillSymbol MySimpleFillPolygon = new SimpleFillSymbolClass();
            MySimpleFillPolygon.Color = MyColor;
            ISymbol MySymbol = MySimpleFillPolygon as ISymbol;
            IRubberBand MyIRubberBand = Polygon;
            IGeometry MyGeometry = MyIRubberBand.TrackNew(MyScreenDispaly, MySymbol);
            MyScreenDispaly.SetSymbol(MySymbol);
            MyScreenDispaly.DrawPolygon(MyGeometry);
            MyScreenDispaly.FinishDrawing();
            ISpatialFilter MyISpatialFilter = new SpatialFilterClass();
            MyISpatialFilter.Geometry = MyGeometry;
            MyISpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects;
            IFeatureLayer SelectedLayer = axMapControl1.ActiveView.FocusMap.get_Layer(IndexOfLayer) as IFeatureLayer;
            IFeatureSelection SelectedFeature = SelectedLayer as IFeatureSelection;
            SelectedFeature.SelectFeatures(MyISpatialFilter, esriSelectionResultEnum.esriSelectionResultNew, false);
            ISelectionSet MyISelectionSet = SelectedFeature.SelectionSet;
            axMapControl1.ActiveView.Refresh();
        }

        public void SelectByLine(int IndexOfLayer,RubberLineClass Line)
        {
            int ConstantNum = 255;
            IActiveView CurrentView = axMapControl1.ActiveView;
            IScreenDisplay CurScreenDisplay = CurrentView.ScreenDisplay;
            CurScreenDisplay.StartDrawing(CurScreenDisplay.hDC, (System.Int16)esriScreenCache.esriNoScreenCache); 
            IRgbColor RGBCOLORS = new ESRI.ArcGIS.Display.RgbColorClass();
            RGBCOLORS.Red = ConstantNum;
            IColor MyColor = RGBCOLORS; 
            ISimpleFillSymbol MySimpleFillSymbol = new SimpleFillSymbolClass();
            MySimpleFillSymbol.Color = MyColor;
            ISymbol MySymbol = MySimpleFillSymbol as ISymbol;
            IRubberBand MyIRubberBand = Line;
            IGeometry MyGeometry = MyIRubberBand.TrackNew(CurScreenDisplay, MySymbol);
            CurScreenDisplay.SetSymbol(MySymbol);
            CurScreenDisplay.DrawPolygon(MyGeometry);
            CurScreenDisplay.FinishDrawing();
            ISpatialFilter MySpatialFilter = new SpatialFilterClass();
            MySpatialFilter.Geometry = MyGeometry;
            MySpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects;
            IFeatureLayer SelectedLayer = axMapControl1.ActiveView.FocusMap.get_Layer(IndexOfLayer) as IFeatureLayer;
            IFeatureSelection SelectedFeatures = SelectedLayer as IFeatureSelection;
            SelectedFeatures.SelectFeatures(MySpatialFilter, esriSelectionResultEnum.esriSelectionResultNew, false);
            ISelectionSet FinalSelection = SelectedFeatures.SelectionSet;
            axMapControl1.ActiveView.Refresh();

        }

        public void SelectByLocayion(ComboBox C28,ComboBox C24,ComboBox C25,ComboBox C26,ComboBox C27)
        {

            IFeatureLayer IntersectedLayer = axMapControl1.ActiveView.FocusMap.get_Layer(C28.SelectedIndex) as IFeatureLayer;
            IFeatureLayer BigLayer = axMapControl1.ActiveView.FocusMap.get_Layer(C24.SelectedIndex) as IFeatureLayer;
            IQueryFilter MyQueryFilter = new QueryFilter();
            string ValueOfCond = C27.SelectedItem.ToString();
            char FirstChar = ValueOfCond[0];
            int Value1 = 48;
            int Value2 = 57;
            if (FirstChar >= Value1 && FirstChar <= Value2)
                MyQueryFilter.WhereClause = C25.SelectedItem.ToString() + C26.SelectedItem.ToString() + C27.SelectedItem.ToString();
            else
                MyQueryFilter.WhereClause = C25.SelectedItem.ToString() + C26.SelectedItem.ToString() + "'" + C27.SelectedItem.ToString() + "'";

            IFeatureCursor ResultOfSearch = BigLayer.FeatureClass.Search(MyQueryFilter, true);
            IFeature SingleFeature = ResultOfSearch.NextFeature();
            ISpatialFilter MySpatialFilter = new SpatialFilter();
            IFeatureSelection MyIFeatureSelection = IntersectedLayer as IFeatureSelection;
            while (SingleFeature != null)
            {

                MySpatialFilter.Geometry = SingleFeature.Shape;
                MySpatialFilter.SpatialRel = esriSpatialRelEnum.esriSpatialRelIntersects;
                MyIFeatureSelection.SelectFeatures(MySpatialFilter, esriSelectionResultEnum.esriSelectionResultAdd, false);
                axMapControl1.Refresh();
                SingleFeature = ResultOfSearch.NextFeature();

            }
        
        }

        
        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CheckLoading = 0;
            if (LoadMap() == true)
            {
                CheckLoading = 1;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            SetInitialValues();
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void axMapControl1_OnMouseDown(object sender, ESRI.ArcGIS.Controls.IMapControlEvents2_OnMouseDownEvent e)
        {
            SelectPoint(e.mapX.ToString(), e.mapY.ToString(), CheckLoading);
        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (CheckLoading == 1)
            {
                HideLayer(checkedListBox1.SelectedIndex);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ChangeDataFrame(comboBox1.SelectedIndex);
        }

        private void axMapControl1_OnKeyDown(object sender, ESRI.ArcGIS.Controls.IMapControlEvents2_OnKeyDownEvent e)
        {


        }

        private void axMapControl1_LocationChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {


        }

        private void button2_Click(object sender, EventArgs e)
        {


        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {


        }


        private void axMapControl1_OnMouseMove(object sender, ESRI.ArcGIS.Controls.IMapControlEvents2_OnMouseMoveEvent e)
        {
            
                SetXY(e.mapX.ToString(), e.mapY.ToString(), CheckLoading);  

        }

        private void closeToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {

                RemoveLayer(comboBox3.SelectedIndex, comboBox3.SelectedItem.ToString());  
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            AddLayer(comboBox12.SelectedItem.ToString());
        }

        

        private void button3_Click_1(object sender, EventArgs e)
        {
            LoadDB();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            
            ApplyDefinisionQuery(comboBox4.SelectedItem.ToString(), comboBox5.SelectedItem.ToString(), comboBox9.SelectedItem.ToString());
            

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetFieldsOfLayer(comboBox2.SelectedIndex, comboBox4);
          
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            

            GetFieldsOfLayer(comboBox8.SelectedIndex, comboBox7);
            WriteHeadersInGridView(comboBox8.SelectedIndex, table);

        }

        private void button5_Click_1(object sender, EventArgs e)
        {

            SelectByAttribute(comboBox7.SelectedItem.ToString(), comboBox6.SelectedItem.ToString(), comboBox10.SelectedItem.ToString(),comboBox11.SelectedItem.ToString());

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            GetUniqueValuesOfFiled(comboBox2.SelectedIndex, comboBox4.SelectedIndex, comboBox9);
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetUniqueValuesOfFiled(comboBox8.SelectedIndex, comboBox7.SelectedIndex, comboBox10);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

       

        private void button7_Click(object sender, EventArgs e)
        {
            ClearSelection(comboBox7.SelectedItem.ToString(), comboBox6.SelectedItem.ToString(), comboBox10.SelectedItem.ToString());
        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            RemoveDefinisionQuery(comboBox2.SelectedIndex);
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        

        

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {
           if(comboBox14.SelectedItem.ToString()=="Polygon")
           {
               SelectByPolygon(comboBox13.SelectedIndex,new RubberPolygonClass());
           }
           else if (comboBox14.SelectedItem.ToString() == "Line")
           {
               SelectByLine(comboBox13.SelectedIndex,new RubberLineClass());
           }
        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            IFeatureLayer FoucsedLayer = axMapControl1.ActiveView.FocusMap.get_Layer(comboBox15.SelectedIndex) as IFeatureLayer;
            IGeoFeatureLayer GLayer = FoucsedLayer as IGeoFeatureLayer;
            CategoryUniqueValues(GLayer, comboBox16.SelectedItem.ToString());
        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetFieldsOfLayer(comboBox15.SelectedIndex, comboBox16);
        }

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetUniqueValuesOfFiled(comboBox18.SelectedIndex, comboBox17.SelectedIndex, comboBox19);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            IFeatureLayer FoucsedLayer = axMapControl1.ActiveView.FocusMap.get_Layer(comboBox18.SelectedIndex) as IFeatureLayer;
            IGeoFeatureLayer GLayer = FoucsedLayer as IGeoFeatureLayer;
            UniqueValueMultiCondition(GLayer, comboBox17.SelectedItem.ToString(), comboBox19.SelectedItem.ToString());

        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetFieldsOfLayer(comboBox18.SelectedIndex, comboBox17);
        }

        private void comboBox22_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetFieldsOfLayer(comboBox22.SelectedIndex, comboBox21);
            GetFieldsOfLayer(comboBox22.SelectedIndex, comboBox20);
            GetFieldsOfLayer(comboBox22.SelectedIndex, comboBox23);


        }

        private void button10_Click(object sender, EventArgs e)
        {
            
        }

        private void button12_Click(object sender, EventArgs e)
        {
            ListofUniqueValues.Clear();
            button9.Enabled = false;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            string FirstChar = comboBox19.SelectedItem.ToString();
            char FChar = FirstChar[0];
            string Cond = "";
            int FirstNum = 48;
            int SecondNum = 57;
            if (FChar >= FirstNum && FChar <= SecondNum)
                Cond = comboBox17.SelectedItem.ToString() + "=" + comboBox19.SelectedItem.ToString();
            else
                Cond = comboBox17.SelectedItem.ToString() + "=" +"'"+ comboBox19.SelectedItem.ToString()+"'";
            if(!ListofUniqueValues.Contains(Cond))
            ListofUniqueValues.Add(Cond);
            
            button9.Enabled = true;
        }

        private void button13_Click(object sender, EventArgs e)
        {

            SelectByLocayion(comboBox28, comboBox24, comboBox25, comboBox26, comboBox27);

        
        }

        private void comboBox24_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetFieldsOfLayer(comboBox24.SelectedIndex, comboBox25);
        }

        private void comboBox25_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetUniqueValuesOfFiled(comboBox24.SelectedIndex, comboBox25.SelectedIndex, comboBox27);
        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

       
    }
}
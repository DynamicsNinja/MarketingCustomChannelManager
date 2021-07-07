using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Serialization;
using Fic.XTB.MarketingCustomChannelManager.Helper;
using Fic.XTB.MarketingCustomChannelManager.Model;
using Fic.XTB.MarketingCustomChannelManager.Proxy;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using XrmToolBox.Extensibility;
using ComboBox = System.Windows.Forms.ComboBox;
using Label = Fic.XTB.MarketingCustomChannelManager.Model.Label;
using Encoding = System.Text.Encoding;


namespace Fic.XTB.MarketingCustomChannelManager
{
    public partial class MarketingCustomChannelManager : PluginControlBase
    {
        private Settings settings;

        public TileXml TileXml = new TileXml();
        public TileCss TileCss;

        public List<int> LocaleIds;

        public List<AttributeMetadata> ContactFields;

        private string _responseTypeIdOldValue;
        private string _responseTypeIdNewValue;

        private string _fileSuffix = "CustomerJourneyDesignerTileConfig";
        private string _publisherPrefix = "";
        private string _solutionUniqueName = "";
        private string _tileName = "";

        private PrivateFontCollection _pfc = new PrivateFontCollection();


        public MarketingCustomChannelManager()
        {
            InitializeComponent();
        }

        private void MarketingCustomChannelManager_Load(object sender, EventArgs e)
        {
            if (!SettingsManager.Instance.TryLoad(GetType(), out settings))
            {
                settings = new Settings();

                LogWarning("Settings not found => a new settings file has been created!");
            }
            else
            {
                LogInfo("Settings found and loaded");
            }
        }

        private void tsbClose_Click(object sender, EventArgs e)
        {
            CloseTool();
        }

        private void GetSolutions()
        {
            WorkAsync(
                new WorkAsyncInfo("Loading solutions...",
                (eventargs) =>
                {
                    ContactFields = MetadataHelper.LoadEntityDetails(Service, "contact")
                        .EntityMetadata.FirstOrDefault()
                        ?.Attributes.ToList();

                    LoadFont();

                    var query = new QueryExpression("solution")
                    {
                        ColumnSet = { AllColumns = true },

                    };

                    query.AddOrder("friendlyname", OrderType.Ascending);
                    query.Criteria.AddCondition("ismanaged", ConditionOperator.Equal, false);

                    var queryPublisher = query.AddLink("publisher", "publisherid", "publisherid");
                    queryPublisher.Columns.AddColumns("customizationprefix");
                    queryPublisher.LinkCriteria.AddCondition("customizationprefix", ConditionOperator.NotIn, "new");
                    queryPublisher.LinkCriteria.AddCondition("customizationprefix", ConditionOperator.NotNull);
                    eventargs.Result = Service.RetrieveMultiple(query);
                })
                {
                    PostWorkCallBack = (completedargs) =>
                    {
                        if (completedargs.Error != null)
                        {
                            MessageBox.Show(completedargs.Error.Message);
                        }
                        else
                        {
                            if (!(completedargs.Result is EntityCollection result)) { return; }

                            var solutions = result.Entities.Select(f => new SolutionProxy(f)).OrderBy(f => f.ToString()).ToArray();
                            cmbSolution.Items.AddRange(solutions);

                            cmbSolution.Enabled = true;
                        }
                    }
                });
        }

        private void GetCustomChannels()
        {
            var selectedSolution = (SolutionProxy)cmbSolution.SelectedItem;
            cmbCustomChannel.Items.Clear();
            cmbCssFile.Items.Clear();

            WorkAsync(
                new WorkAsyncInfo("Loading web resources...",
                    (eventargs) =>
                    {
                        var componentQe = new QueryExpression("solutioncomponent");
                        componentQe.ColumnSet = new ColumnSet("objectid");
                        componentQe.Criteria.AddCondition("solutionid", ConditionOperator.Equal, selectedSolution.Entity.Id.ToString("D"));
                        componentQe.Criteria.AddCondition("componenttype", ConditionOperator.Equal, 61);

                        var solutionComponents = Service.RetrieveMultiple(componentQe).Entities;

                        if (solutionComponents.Count == 0)
                        {
                            eventargs.Result = new EntityCollection();
                        }
                        else
                        {
                            var objectIds = solutionComponents.Select(c => (Guid)c["objectid"]).ToArray();

                            var query = new QueryExpression("webresource");
                            query.ColumnSet.AddColumns("webresourcetype", "name", "displayname", "content");
                            query.Criteria.AddCondition("webresourcetype", ConditionOperator.In, 2, 4);

                            var filter = new FilterExpression();
                            query.Criteria.AddFilter(filter);
                            filter.FilterOperator = LogicalOperator.Or;

                            foreach (var objectId in objectIds)
                            {
                                filter.AddCondition("webresourceid", ConditionOperator.Equal, objectId.ToString("D"));
                            }

                            query.AddOrder("name", OrderType.Ascending);

                            var webresources = Service.RetrieveMultiple(query);

                            eventargs.Result = webresources;
                        }
                    })
                {
                    PostWorkCallBack = (completedargs) =>
                    {
                        if (completedargs.Error != null)
                        {
                            MessageBox.Show(completedargs.Error.Message);
                        }
                        else
                        {
                            if (!(completedargs.Result is EntityCollection result)) { return; }

                            cmbCustomChannel.Items.Add(new WebresourceProxy
                            {
                                Entity = new Entity
                                {
                                    ["name"] = "Create New"
                                }
                            });

                            var xmls = result.Entities
                                .Where(w => ((OptionSetValue)w["webresourcetype"]).Value == 4)
                                .Select(w => new WebresourceProxy(w)).OrderBy(w => w.ToString()).ToArray();
                            cmbCustomChannel.Items.AddRange(xmls);

                            cmbCssFile.Items.Add(new WebresourceProxy
                            {
                                Entity = new Entity
                                {
                                    ["name"] = "Create New"
                                }
                            });

                            var css = result.Entities
                                .Where(w => ((OptionSetValue)w["webresourcetype"]).Value == 2)
                                .Select(w => new WebresourceProxy(w)).OrderBy(w => w.ToString()).ToArray();
                            cmbCssFile.Items.AddRange(css);
                        }
                    }
                });
        }

        private void GetEntities()
        {
            WorkAsync(
                new WorkAsyncInfo("Loading entities...",
                    (eventargs) =>
                    {
                        eventargs.Result = MetadataHelper.LoadEntities(Service);
                    })
                {
                    PostWorkCallBack = (completedargs) =>
                    {
                        if (completedargs.Error != null)
                        {
                            MessageBox.Show(completedargs.Error.Message);
                        }
                        else
                        {
                            if (!(completedargs.Result is RetrieveMetadataChangesResponse result)) { return; }

                            var entites = result.EntityMetadata.Select(f => new EntityMetadataProxy(f)).OrderBy(f => f.ToString()).ToArray();
                            cmbEntity.Items.AddRange(entites);
                        }
                    }
                });
        }

        private void GetFields()
        {
            var selectedEntity = ((EntityMetadataProxy)cmbEntity.SelectedItem).Entity;
            cmbComplianceField.Items.Clear();
            cmbTitleField.Items.Clear();

            WorkAsync(
                new WorkAsyncInfo("Loading fields...",
                    (eventargs) =>
                    {
                        eventargs.Result = MetadataHelper.LoadEntityDetails(Service, selectedEntity.LogicalName);
                    })
                {
                    PostWorkCallBack = (completedargs) =>
                    {
                        if (completedargs.Error != null)
                        {
                            MessageBox.Show(completedargs.Error.Message);
                        }
                        else
                        {
                            if (!(completedargs.Result is RetrieveMetadataChangesResponse result)) { return; }

                            var boolFields = ContactFields
                                .Where(m => m.AttributeType == AttributeTypeCode.Boolean)
                                .Select(f => new AttributeMetadataProxy(f)).OrderBy(f => f.ToString()).ToArray();
                            cmbComplianceField.Items.AddRange(boolFields);

                            var stringFields = result.EntityMetadata.FirstOrDefault()?.Attributes
                                .Where(m => m.AttributeType == AttributeTypeCode.String)
                                .Select(f => new AttributeMetadataProxy(f)).OrderBy(f => f.ToString()).ToArray();
                            cmbTitleField.Items.AddRange(stringFields);

                            foreach (AttributeMetadataProxy amp in cmbComplianceField.Items)
                            {
                                if (amp.AttributeMetadata.LogicalName != TileXml.ChannelProperties.ComplianceField) { continue; }
                                cmbComplianceField.SelectedItem = amp;
                                break;
                            }

                            foreach (AttributeMetadataProxy amp in cmbTitleField.Items)
                            {
                                if (amp.AttributeMetadata.LogicalName != TileXml.ChannelProperties.TitleFieldName) { continue; }
                                cmbTitleField.SelectedItem = amp;
                                break;
                            }
                        }
                    }
                });
        }

        private void GetViews()
        {
            var selectedEntity = ((EntityMetadataProxy)cmbEntity.SelectedItem).Entity;
            cmbLookupView.Items.Clear();

            WorkAsync(
                new WorkAsyncInfo("Loading views...",
                    (eventargs) =>
                    {
                        var query = new QueryExpression("savedquery");
                        query.ColumnSet.AddColumns("name");
                        query.AddOrder("name", OrderType.Ascending);
                        query.Criteria.AddCondition("returnedtypecode", ConditionOperator.Equal, selectedEntity.ObjectTypeCode);
                        query.Criteria.AddCondition("querytype", ConditionOperator.Equal, 64);

                        eventargs.Result = Service.RetrieveMultiple(query);
                    })
                {
                    PostWorkCallBack = (completedargs) =>
                    {
                        if (!(completedargs.Result is EntityCollection result)) { return; }

                        var views = result.Entities.Select(f => new ViewProxy(f)).OrderBy(f => f.ToString()).ToArray();
                        cmbLookupView.Items.AddRange(views);

                        foreach (ViewProxy v in cmbLookupView.Items)
                        {
                            if (v.Entity.Id.ToString("D").ToLower() != TileXml.ChannelProperties.LookupViewId?.ToLower()) { continue; }
                            cmbLookupView.SelectedItem = v;
                            return;
                        }

                        if (cmbLookupView.Items.Count == 1)
                        {
                            cmbLookupView.SelectedIndex = 0;
                        }
                    }
                });
        }

        private void GetForms()
        {
            var selectedEntity = ((EntityMetadataProxy)cmbEntity.SelectedItem).Entity;
            cmbQuickViewForm.Items.Clear();

            WorkAsync(
                new WorkAsyncInfo("Loading forms...",
                    (eventargs) =>
                    {
                        var query = new QueryExpression("systemform");
                        query.ColumnSet.AddColumns("name");
                        query.Criteria.AddCondition("type", ConditionOperator.Equal, 6);
                        query.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, selectedEntity.ObjectTypeCode);

                        eventargs.Result = Service.RetrieveMultiple(query);
                    })
                {
                    PostWorkCallBack = (completedargs) =>
                    {
                        if (!(completedargs.Result is EntityCollection result)) { return; }

                        var forms = result.Entities.Select(f => new FormProxy(f)).OrderBy(f => f.ToString()).ToArray();
                        cmbQuickViewForm.Items.AddRange(forms);

                        foreach (FormProxy f in cmbQuickViewForm.Items)
                        {
                            if (f.Entity.Id.ToString("D").ToLower() != TileXml.ChannelProperties.QuickViewFormId?.ToLower()) { continue; }
                            cmbQuickViewForm.SelectedItem = f;
                            return;
                        }

                        if (cmbQuickViewForm.Items.Count == 1)
                        {
                            cmbQuickViewForm.SelectedIndex = 0;
                        }
                    }
                });
        }

        private void MyPluginControl_OnCloseTool(object sender, EventArgs e)
        {
            SettingsManager.Instance.Save(GetType(), settings);
        }

        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);

            if (settings == null || detail == null) { return; }

            settings.LastUsedOrganizationWebappUrl = detail.WebApplicationUrl;
            LogInfo("Connection has changed to: {0}", detail.WebApplicationUrl);
        }

        private void MarketingCustomChannelManager_ConnectionUpdated(object sender, ConnectionUpdatedEventArgs e)
        {
            ExecuteMethod(GetLocales);
            ExecuteMethod(GetSolutions);
            ExecuteMethod(GetEntities);
        }

        private void LoadFont()
        {
            var fontAsStream = this.GetType().Assembly.GetManifestResourceStream("Fic.XTB.MarketingCustomChannelManager.CRMMDL2.ttf");
            var fontAsByte = new byte[fontAsStream.Length];
            fontAsStream.Read(fontAsByte, 0, (int)fontAsStream.Length);
            fontAsStream.Close();
            var memPointer = System.Runtime.InteropServices.Marshal.AllocHGlobal(System.Runtime.InteropServices.Marshal.SizeOf(typeof(byte)) * fontAsByte.Length);
            System.Runtime.InteropServices.Marshal.Copy(fontAsByte, 0, memPointer, fontAsByte.Length);
            _pfc.AddMemoryFont(memPointer, fontAsByte.Length);
        }

        private void GetLocales()
        {
            var request = new RetrieveAvailableLanguagesRequest();
            var response = (RetrieveAvailableLanguagesResponse)Service.Execute(request);

            var query = new QueryExpression("languagelocale");
            query.ColumnSet.AllColumns = true;


            var filter = new FilterExpression();
            query.Criteria.AddFilter(filter);

            filter.FilterOperator = LogicalOperator.Or;

            LocaleIds = response.LocaleIds.ToList();

            foreach (var lcid in response.LocaleIds)
            {
                filter.AddCondition("localeid", ConditionOperator.Equal, lcid);
            }

            var locales = Service.RetrieveMultiple(query).Entities;

            foreach (var loc in locales)
            {
                var displayName = $"{loc["language"]} ({loc["localeid"]})";
                var lcid = $"{loc["localeid"]}";

                var localeProxy = new LocaleProxy(lcid, displayName);

                ((DataGridViewComboBoxColumn)dgvLabels.Columns["labelLcid"])?.Items.Add(localeProxy);
                ((DataGridViewComboBoxColumn)dgvTooltips.Columns["tooltipLcid"])?.Items.Add(localeProxy);
                ((DataGridViewComboBoxColumn)dgvResponseLabels.Columns["responseLcid"])?.Items.Add(localeProxy);
            }
        }

        private void cmbSolution_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetCustomChannels();

            var selectedSolution = ((SolutionProxy)cmbSolution.SelectedItem).Entity;
            _publisherPrefix = ((AliasedValue)selectedSolution["publisher1.customizationprefix"]).Value.ToString();
            _solutionUniqueName = selectedSolution["uniquename"].ToString();

            cmbCustomChannel.Enabled = true;
        }

        private void cmbEntity_SelectedIndexChanged(object sender, EventArgs e)
        {
            GetFields();
            GetViews();
            GetForms();

            //cmbComplianceField.SelectedIndex = cmbComplianceField.FindString(tileXml.ChannelProperties.ComplianceField);

            cmbTitleField.SelectedIndex = cmbTitleField.FindString(TileXml.ChannelProperties.TitleFieldName);
            cmbLookupView.SelectedIndex = cmbLookupView.FindString(TileXml.ChannelProperties.LookupViewId);
            cmbQuickViewForm.SelectedIndex = cmbQuickViewForm.FindString(TileXml.ChannelProperties.QuickViewFormId);
        }

        private void cmbCssFile_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedCss = (WebresourceProxy)cmbCssFile.SelectedItem;

            if (!selectedCss.Entity.Contains("content"))
            {
                tbFont.ReadOnly = true;
                tbIcon.ReadOnly = true;

                TileCss = new TileCss(_tileName);

                TileCss.FontFamilyClass = $"{_tileName}TileSymbolFont";
                TileCss.IconClass = $"{_tileName}Tile";

                tbCssFile.Text = $"{_publisherPrefix}_/css/{_tileName}{_fileSuffix}.css"; ;
            }
            else
            {
                tbFont.ReadOnly = false;
                tbIcon.ReadOnly = false;

                var content = selectedCss.Entity["content"].ToString();
                var data = Convert.FromBase64String(content);
                var cssContent = Encoding.Default.GetString(data);

                TileCss = new TileCss(_tileName, cssContent);

                tbCssFile.Text = selectedCss.Entity["name"].ToString();
            }


            tbFont.Text = TileCss.FontFamilyClass;
            tbIcon.Text = TileCss.IconClass;

            tbTileBackgroundColor.Text = TileCss.TileBackgroundColor;
            tbIconBackgroundColor.Text = TileCss.IconBackgroundColor;
            tbTileBorderColor.Text = TileCss.TileBorderColor;
            tbLeftBorderColor.Text = TileCss.LeftBorderColor;
        }

        private void cmbFontFamily_SelectedIndexChanged(object sender, EventArgs e)
        {
            tbFont.Text = cmbFontFamily.SelectedItem.ToString();
        }

        private void cmbCustomChannel_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedChannel = ((WebresourceProxy)cmbCustomChannel.SelectedItem).Entity;

            if (selectedChannel.Id == Guid.Empty)
            {
                TileXml = new TileXml();
                InitResponseTypes();
                InitLabels();
                InitTooltips();
                InitDefinition();
                InitChannelProperties();
            }
            else
            {
                var base64Content = selectedChannel["content"].ToString();

                var base64EncodedBytes = Convert.FromBase64String(base64Content);
                var xmlContent = Encoding.UTF8.GetString(base64EncodedBytes);

                var serializer = new XmlSerializer(typeof(TileXml));
                using (TextReader reader = new StringReader(xmlContent))
                {
                    TileXml = (TileXml)serializer.Deserialize(reader);
                }

                PopulateLabelsGrid();
                PopulateTooltipsGrid();
                PopulateResponseTypesGrid();
                PopulateDefinitionSection();
                PopulateChannelPropertiesSection();
            }

            tabControl.Enabled = true;
            tsbPublish.Enabled = true;
        }

        private void PopulateDefinitionSection()
        {
            tbIcon.Text = TileXml.Definition.Icon;
            tbFont.Text = TileXml.Definition.FontFamily;

            foreach (WebresourceProxy wrp in cmbCssFile.Items)
            {
                if (wrp.Entity["name"].ToString() != TileXml.Definition.CssFileName) { continue; }
                cmbCssFile.SelectedItem = wrp;
                break;
            }
        }

        private void PopulateChannelPropertiesSection()
        {
            var entityName = TileXml.ChannelProperties.EntityType;
            foreach (EntityMetadataProxy emp in cmbEntity.Items)
            {
                if (emp.Entity.LogicalName != entityName) { continue; }
                cmbEntity.SelectedItem = emp;
                break;
            }
        }

        private void InitLabels()
        {
            TileXml.Labels.Label = new List<Label>();

            foreach (var lcid in LocaleIds)
            {
                TileXml.Labels.Label.Add(new Label
                {
                    Text = "New Tile",
                    LocId = lcid.ToString()
                });
            }

            PopulateLabelsGrid();
        }

        private void PopulateLabelsGrid()
        {
            dgvLabels.Rows.Clear();

            foreach (var label in TileXml.Labels.Label)
            {
                var rowId = dgvLabels.Rows.Add();
                var row = dgvLabels.Rows[rowId];
                row.Cells["labelText"].Value = label.Text;

                var localeDropdown = (DataGridViewComboBoxCell)row.Cells["labelLcid"];

                foreach (LocaleProxy lp in localeDropdown.Items)
                {
                    if (lp.Lcid != label.LocId) { continue; }
                    localeDropdown.Value = lp;
                    break;
                }
            }

            _tileName = TileXml.Labels.Label.FirstOrDefault()?.Text.Replace(" ", "");
        }

        private void InitTooltips()
        {
            TileXml.Tooltips.Tooltip = new List<Tooltip>();
            foreach (var lcid in LocaleIds)
            {
                TileXml.Tooltips.Tooltip.Add(new Tooltip
                {
                    Text = "New Tooltip",
                    LocId = lcid.ToString()
                });
            }

            PopulateTooltipsGrid();
        }

        private void InitChannelProperties()
        {
            cmbEntity.SelectedIndex = 0;
        }

        private void InitDefinition()
        {
            cmbCssFile.SelectedIndex = 0;
        }

        private void PopulateTooltipsGrid()
        {

            dgvTooltips.Rows.Clear();

            foreach (var tooltip in TileXml.Tooltips.Tooltip)
            {
                var rowId = dgvTooltips.Rows.Add();
                var row = dgvTooltips.Rows[rowId];
                row.Cells["tooltipText"].Value = tooltip.Text;

                var localeDropdown = (DataGridViewComboBoxCell)row.Cells["tooltipLcid"];

                foreach (LocaleProxy lp in localeDropdown.Items)
                {
                    if (lp.Lcid != tooltip.LocId) { continue; }
                    localeDropdown.Value = lp;
                    break;
                }
            }
        }

        private void InitResponseTypes()
        {
            TileXml.ResponseTypes.ResponseType = new List<ResponseType>();
            TileXml.ResponseTypes.ResponseType.Add(new ResponseType("sent", LocaleIds));
            TileXml.ResponseTypes.ResponseType.Add(new ResponseType("delivered", LocaleIds));

            PopulateResponseTypesGrid();
        }

        private void PopulateResponseTypesGrid()
        {
            dgvResponseTypes.Rows.Clear();

            foreach (var responseType in TileXml.ResponseTypes.ResponseType)
            {
                var rowId = dgvResponseTypes.Rows.Add();
                var row = dgvResponseTypes.Rows[rowId];
                row.ReadOnly = true;
                row.Cells["ID"].Value = responseType.Id;
            }

            PopulateResponseLabelsGrid(TileXml.ResponseTypes.ResponseType.FirstOrDefault()?.Id);
        }

        private void dgvResponseTypes_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            var rowId = e.RowIndex;

            if (rowId == -1) { return; }

            var row = dgvResponseTypes.Rows[rowId];

            var responseId = row.Cells["ID"].Value?.ToString();

            PopulateResponseLabelsGrid(responseId);
        }

        private void PopulateResponseLabelsGrid(string responseId)
        {
            dgvResponseLabels.Rows.Clear();

            if (responseId == null) { return; }

            var responseLabels = TileXml.ResponseTypes.ResponseType.FirstOrDefault(r => r.Id == responseId)?.Labels.Label;

            foreach (var responseLabel in responseLabels)
            {
                var rowId = dgvResponseLabels.Rows.Add();
                var row = dgvResponseLabels.Rows[rowId];
                row.Cells["responseValue"].Value = responseLabel.Text;

                var localeDropdown = (DataGridViewComboBoxCell)row.Cells["responseLcid"];

                foreach (LocaleProxy lp in localeDropdown.Items)
                {
                    if (lp.Lcid != responseLabel.LocId) { continue; }
                    localeDropdown.Value = lp;
                    break;
                }
            }
        }

        private void dgvResponseTypes_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            _responseTypeIdOldValue = dgvResponseTypes[e.ColumnIndex, e.RowIndex].Value?.ToString();
        }

        private void dgvResponseTypes_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            _responseTypeIdNewValue = dgvResponseTypes[e.ColumnIndex, e.RowIndex].Value?.ToString();

            if (_responseTypeIdNewValue == _responseTypeIdOldValue) { return; }

            if (_responseTypeIdOldValue == null)
            {
                var responseType = new ResponseType(_responseTypeIdNewValue, LocaleIds);
                responseType.Custom = "True";
                TileXml.ResponseTypes.ResponseType.Add(responseType);
            }
            else
            {
                var responseType = TileXml.ResponseTypes.ResponseType.FirstOrDefault(r => r.Id == _responseTypeIdOldValue);
                responseType.Id = _responseTypeIdNewValue;
            }
        }

        private void dgvResponseLabels_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            var selectedResponseTypeValue = dgvResponseTypes.SelectedCells[0].Value.ToString();

            var responseType = TileXml.ResponseTypes.ResponseType.FirstOrDefault(r => r.Id == selectedResponseTypeValue);

            var label = dgvResponseLabels.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            var lcid = ((LocaleProxy)dgvResponseLabels.Rows[e.RowIndex].Cells[0].Value).Lcid;

            var responseTypeLabel = responseType.Labels.Label.FirstOrDefault(l => l.LocId == lcid);

            responseTypeLabel.Text = label;
        }

        private string GenerateTileXml()
        {
            TileXml.Definition = new Definition
            {
                CssFileName = tbCssFile.Text,
                FontFamily = tbFont.Text,
                Icon = tbIcon.Text,
            };

            TileXml.ChannelProperties = new ChannelProperties
            {
                ComplianceField = ((AttributeMetadataProxy)cmbComplianceField.SelectedItem).AttributeMetadata.LogicalName,
                EntitySetName = ((EntityMetadataProxy)cmbEntity.SelectedItem).Entity.LogicalCollectionName,
                EntityType = ((EntityMetadataProxy)cmbEntity.SelectedItem).Entity.LogicalName,
                LookupViewId = ((ViewProxy)cmbLookupView.SelectedItem).Entity.Id.ToString("D"),
                QuickViewFormId = ((FormProxy)cmbQuickViewForm.SelectedItem).Entity.Id.ToString("D"),
                TitleFieldName = ((AttributeMetadataProxy)cmbTitleField.SelectedItem).AttributeMetadata.LogicalName,
            };

            var xmlSerializer = new XmlSerializer(TileXml.GetType());

            using (var textWriter = new StringWriter())
            {
                xmlSerializer.Serialize(textWriter, TileXml);
                var xmlContent = textWriter.ToString();

                return xmlContent;
            }
        }

        private void dgvLabels_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            var label = dgvLabels.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            var lcid = ((LocaleProxy)((DataGridViewComboBoxCell)dgvLabels.Rows[e.RowIndex].Cells[0]).Value).Lcid;
            var tileLabel = TileXml.Labels.Label.FirstOrDefault(l => l.LocId == lcid);

            if (lcid == "1033")
            {
                _tileName = label.Replace(" ", "");

                tbFont.Text = $"{_tileName}TileSymbolFont";
                tbIcon.Text = $"{_tileName}Tile";
            }

            tileLabel.Text = label;
        }

        private void dgvTooltips_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            var tooltip = dgvTooltips.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            var lcid = ((LocaleProxy)((DataGridViewComboBoxCell)dgvLabels.Rows[e.RowIndex].Cells[0]).Value).Lcid;

            var tileTooltip = TileXml.Tooltips.Tooltip.FirstOrDefault(t => t.LocId == lcid);

            tileTooltip.Text = tooltip;
        }

        private void AddWebresourceToCombobox(ComboBox cbx, WebresourceProxy cmbItem)
        {
            Invoke(new Action(() =>
            {
                cbx.Items.Add(cmbItem);
                cbx.SelectedItem = cmbItem;
            }));
        }

        private void tsbPublish_Click(object sender, EventArgs e)
        {
            var selectedCustomChannel = ((WebresourceProxy)cmbCustomChannel.SelectedItem).Entity;

            var name = TileXml.Labels.Label.FirstOrDefault()?.Text.Replace(" ", "");

            var selectedCssFile = ((WebresourceProxy)cmbCssFile.SelectedItem).Entity;

            var cssFilename = selectedCssFile.Contains("content") 
                ? (string)selectedCssFile["name"]
                : $"{_publisherPrefix}_/css/{name}{_fileSuffix}.css";

            var cssContent = TileCss.GenerateFileContent();
            var plainTextBytes = Encoding.UTF8.GetBytes(cssContent);
            var cssBase64 = Convert.ToBase64String(plainTextBytes);

            var cssFile = new Entity("webresource");
            cssFile["content"] = cssBase64;
            cssFile["webresourcetype"] = new OptionSetValue(2);
            cssFile["name"] = cssFilename;
            cssFile["displayname"] = cssFilename;
            cssFile["description"] = cssFilename;

            TileXml.Definition.CssFileName = cssFilename;

            var xml = GenerateTileXml();
            plainTextBytes = Encoding.UTF8.GetBytes(xml);
            var base64Content = Convert.ToBase64String(plainTextBytes);

            var xmlFilename = $"{_publisherPrefix}_/xml/{name}{_fileSuffix}.xml";

            var xmlFile = new Entity("webresource");
            xmlFile["content"] = base64Content;
            xmlFile["webresourcetype"] = new OptionSetValue(4);
            xmlFile["name"] = xmlFilename;
            xmlFile["displayname"] = xmlFilename;
            xmlFile["description"] = xmlFilename;

            WorkAsync(
                new WorkAsyncInfo("Publishing custom channel...",
                    (eventargs) =>
                    {
                        // XML File

                        if (selectedCustomChannel.Id == Guid.Empty)
                        {
                            var cr = new CreateRequest { Target = xmlFile };
                            cr.Parameters.Add("SolutionUniqueName", _solutionUniqueName);
                            var cresp = (CreateResponse)Service.Execute(cr);

                            PublishWebresource(cresp.id.ToString("D"));

                            var entity = Service.Retrieve("webresource", cresp.id,
                                new ColumnSet("webresourcetype", "name", "displayname", "content"));

                            var cmbItem = new WebresourceProxy(entity);
                            AddWebresourceToCombobox(cmbCustomChannel,cmbItem);
                        }
                        else
                        {
                            xmlFile.Id = selectedCustomChannel.Id;
                            var ur = new UpdateRequest { Target = xmlFile };
                            ur.Parameters.Add("SolutionUniqueName", _solutionUniqueName);
                            var cresp = (UpdateResponse)Service.Execute(ur);

                            PublishWebresource(xmlFile.Id.ToString("D"));
                        }

                        // CSS File

                        if (selectedCssFile.Id == Guid.Empty)
                        {
                            var cr = new CreateRequest { Target = cssFile };
                            cr.Parameters.Add("SolutionUniqueName", _solutionUniqueName);
                            var cresp = (CreateResponse)Service.Execute(cr);

                            PublishWebresource(cresp.id.ToString("D"));

                            var entity = Service.Retrieve("webresource", cresp.id,
                                new ColumnSet("webresourcetype", "name", "displayname", "content"));

                            var cmbItem = new WebresourceProxy(entity);

                            AddWebresourceToCombobox(cmbCssFile,cmbItem);
                        }
                        else
                        {
                            cssFile.Id = selectedCssFile.Id;
                            var ur = new UpdateRequest { Target = cssFile };
                            ur.Parameters.Add("SolutionUniqueName", _solutionUniqueName);
                            var cresp = (UpdateResponse)Service.Execute(ur);

                            PublishWebresource(cssFile.Id.ToString("D"));
                        }
                    })
                {
                    PostWorkCallBack = (completedargs) =>
                    {
                        if (completedargs.Error != null)
                        {
                            MessageBox.Show(completedargs.Error.Message);
                        }
                    }
                });
        }

        private void PublishWebresource(string webresourceId)
        {
            var webResctag = "<webresource>" + webresourceId + "</webresource>";
            var webrescXml = "<importexportxml><webresources>" + webResctag + "</webresources></importexportxml>";
            var publishxmlrequest = new PublishXmlRequest { ParameterXml = string.Format(webrescXml) };
            var presp = (PublishXmlResponse)Service.Execute(publishxmlrequest);
        }

        private void ChangeColor(TextBox tb, Button btn)
        {
            var colorDialog = new ColorDialog
            {
                AllowFullOpen = true,
                ShowHelp = false,
                FullOpen = true,
                AnyColor = true,
            };

            if (colorDialog.ShowDialog() != DialogResult.OK) { return; }

            tb.Text = ColorTranslator.ToHtml(colorDialog.Color);
            btn.BackColor = colorDialog.Color;
        }

        private void btnTileBackgroundColor_Click(object sender, EventArgs e)
        {
            ChangeColor(tbTileBackgroundColor, btnTileBackgroundColor);
        }

        private void btnIconBackgroundColor_Click(object sender, EventArgs e)
        {
            ChangeColor(tbIconBackgroundColor, btnIconBackgroundColor);
        }

        private void btnTileBorderColor_Click(object sender, EventArgs e)
        {
            ChangeColor(tbTileBorderColor, btnTileBorderColor);
        }

        private void btnLeftBorderColor_Click(object sender, EventArgs e)
        {
            ChangeColor(tbLeftBorderColor, btnLeftBorderColor);
        }

        private void tbTileBackgroundColor_TextChanged(object sender, EventArgs e)
        {
            var colorString = tbTileBackgroundColor.Text;

            if (colorString == string.Empty) { return; }

            var color = GetColor(colorString);
            btnTileBackgroundColor.BackColor = color;

            TileCss.TileBackgroundColor = colorString;
        }

        private void tbIconBackgroundColor_TextChanged(object sender, EventArgs e)
        {
            var colorString = tbIconBackgroundColor.Text;

            if (colorString == string.Empty) { return; }

            var color = GetColor(colorString);
            btnIconBackgroundColor.BackColor = color;

            TileCss.IconBackgroundColor = colorString;
        }

        private void tbTileBorderColor_TextChanged(object sender, EventArgs e)
        {
            var colorString = tbTileBorderColor.Text;

            if (colorString == string.Empty) { return; }

            var color = GetColor(colorString);
            btnTileBorderColor.BackColor = color;

            TileCss.TileBorderColor = colorString;
        }

        private void tbLeftBorderColor_TextChanged(object sender, EventArgs e)
        {
            var colorString = tbLeftBorderColor.Text;

            if (colorString == string.Empty) { return; }

            var color = GetColor(colorString);
            btnLeftBorderColor.BackColor = color;

            TileCss.LeftBorderColor = colorString;
        }

        private Color GetColor(string colorString)
        {
            var color = new Color();

            try
            {
                if (colorString.Contains("rgb"))
                {
                    var values = colorString
                        .Replace("rgb(", "")
                        .Replace(")", "")
                        .Split(',')
                        .Select(d => int.Parse(d.Trim()))
                        .ToList();

                    color = Color.FromArgb(255, values[0], values[1], values[2]);
                }
                else if (colorString.StartsWith("#"))
                {
                    color = ColorTranslator.FromHtml(colorString);
                }
                else
                {
                    color = Color.FromName(colorString);
                }
            }
            catch (Exception)
            {
                // ignored
            }

            return color;
        }
    }
}
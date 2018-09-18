/******************************************************************
** auth: wei.huazhong
** date: 9/17/2018 12:05:58 PM
** desc:
******************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace ExcelService
{
    public sealed class ExcelImportConfigHandler
    {
        private Dictionary<string, ExcelImportData> importDatas;

        public ExcelImportConfigHandler()
        {
            importDatas = new Dictionary<string, ExcelImportData>();
            string configPath = GetConfigPath();
            LoadConfig(configPath, importDatas);
        }

        private String GetConfigPath()
        {
            Uri baseUri = new Uri(System.Reflection.Assembly.GetCallingAssembly().CodeBase);

            Uri sheetConfigUri = new Uri(baseUri, @"..\..\Configs\ExcelImport.cfg.xml");

            return sheetConfigUri.LocalPath;
        }

        private void LoadConfig(string path, Dictionary<string, ExcelImportData> importDatas)
        {
            var xdoc = XDocument.Load(path);
            var dataElements = xdoc.Root.Elements();
            ExcelImportData importData;
            foreach (var dataElement in dataElements)
            {
                importData = new ExcelImportData();
                importData.SheetIndex = int.Parse(dataElement.Attribute("sheetIndex").Value);
                importData.DataStartRow = int.Parse(dataElement.Attribute("dataStartRow").Value);
                importData.Entity = dataElement.Attribute("entity").Value;
                if (dataElement.Attribute("checkEndCol") != null)
                {
                    importData.CheckEndCol = int.Parse(dataElement.Attribute("checkEndCol").Value);
                }
                if (dataElement.Attribute("checkEndValue") != null)
                {
                    importData.CheckEndValue = dataElement.Attribute("checkEndValue").Value;
                }
                if (dataElement.Attribute("fileTypeCol") != null)
                {
                    importData.FileTypeColumn = int.Parse(dataElement.Attribute("fileTypeCol").Value);
                }
                if (dataElement.Attribute("titleRow") != null)
                {
                    importData.TitleRow = int.Parse(dataElement.Attribute("titleRow").Value);
                }

                importData.Columns = GetDataColumns(dataElement, Type.GetType(importData.Entity));
                importDatas.Add(dataElement.Name.LocalName, importData);
            }
        }

        private List<ExcelImportColumn> GetDataColumns(XElement dataElement, Type entityType)
        {
            var columns = new List<ExcelImportColumn>();
            var columnElements = dataElement.Elements("column");
            ExcelImportColumn column;
            foreach (var e in columnElements)
            {
                column = new ExcelImportColumn();
                column.Name = e.Attribute("name").Value;
                column.PropertyInfo = entityType.GetProperty(e.Attribute("property").Value);
                column.Col = int.Parse(e.Attribute("col").Value);
                column.DataType = Type.GetType(e.Attribute("type").Value);

                if (e.Attribute("required") != null)
                {
                    column.Required = bool.Parse(e.Attribute("required").Value);
                }
                if (e.Attribute("maxlength") != null)
                {
                    column.MaxLength = int.Parse(e.Attribute("maxlength").Value);
                }
                if (e.Attribute("regexp") != null)
                {
                    column.Regexp = e.Attribute("regexp").Value;
                }
                if (e.Attribute("coltorow") != null)
                {
                    column.ColToRow = bool.Parse(e.Attribute("coltorow").Value);
                }
                if (e.Attribute("headerproperty") != null)
                {
                    column.HeaderPropertyInfo = entityType.GetProperty(e.Attribute("headerproperty").Value);
                }
                if (e.Attribute("headervalue") != null)
                {
                    column.HeaderValue = e.Attribute("headervalue").Value;
                }
                if (e.Attribute("headertype") != null)
                {
                    column.HeaderDataType = Type.GetType(e.Attribute("headertype").Value);
                }
                if (e.Attribute("valuemapping") != null)
                {
                    column.ValueMapping = bool.Parse(e.Attribute("valuemapping").Value);
                    if (column.ValueMapping)
                    {
                        column.ValueType = Type.GetType(e.Attribute("valuetype").Value);
                        column.InitValueMapping();
                        FillColumnValueMappings(column, e);
                    }
                }

                columns.Add(column);
            }

            return columns;
        }

        private void FillColumnValueMappings(ExcelImportColumn column, XElement columnElement)
        {
            var mappingElements = columnElement.Descendants("mapping");
            foreach (var mappingElement in mappingElements)
            {
                var key = mappingElement.Attribute("key").Value;
                var value = mappingElement.Attribute("value").Value;
                column.AddMappingValue(Convert.ChangeType(key, column.DataType), Convert.ChangeType(value, column.ValueType));
            }
        }

        class Nested
        {
            public static ExcelImportConfigHandler instance;
            static Nested()
            {
                if (instance == null)
                    instance = new ExcelImportConfigHandler();
            }
            //static Nested() { }
            //internal static readonly ExcelImportConfigHandler instance = new ExcelImportConfigHandler();
        }

        public ExcelImportConfigHandler Instance
        {
            get
            {
                return Nested.instance;
            }
        }

        public ExcelImportData GetExcelImportDataConfig(string importDataName)
        {
            if (importDatas.ContainsKey(importDataName))
            {
                return importDatas[importDataName];
            }
            else
            {
                throw new NullReferenceException(string.Format("no import data config of '{0}'", importDataName));
            }
        }
    }

    public class ExcelImportData
    {
        public int SheetIndex
        {
            set;
            get;
        }

        public int DataStartRow
        {
            set;
            get;
        }

        public string Entity
        {
            set;
            get;
        }

        public int CheckEndCol
        {
            set;
            get;
        }

        public string CheckEndValue
        {
            set;
            get;
        }

        public int FileTypeColumn
        {
            set;
            get;
        }

        public int TitleRow
        {
            set;
            get;
        }

        public List<ExcelImportColumn> Columns
        {
            set;
            get;
        }
    }

    public class ExcelImportColumn
    {
        #region properties

        public string Name { get; set; }
        public PropertyInfo PropertyInfo { get; set; }
        public int Col { get; set; }
        public Type DataType { get; set; }
        public bool Required { get; set; }
        public bool ColToRow { get; set; }
        public PropertyInfo HeaderPropertyInfo { get; set; }
        public string HeaderValue { get; set; }
        public Type HeaderDataType { get; set; }
        public bool ValueMapping { get; set; }
        public Type ValueType { get; set; }
        public int MaxLength { get; set; }
        public string Regexp { get; set; }

        #endregion

        private object _valuemappings;
        private MethodInfo _containKeyMethod;
        private MethodInfo _getValueMethod;
        private MethodInfo _addMethod;

        internal void InitValueMapping()
        {
            var dicyType = typeof(Dictionary<,>);
            Type[] typeArgs = { DataType, ValueType };
            var dictCreateype = dicyType.MakeGenericType(typeArgs);
            _valuemappings = Activator.CreateInstance(dictCreateype);

            _containKeyMethod = dictCreateype.GetMethod("ContainsKey", BindingFlags.Instance | BindingFlags.Public);
            _getValueMethod = dictCreateype.GetMethod("get_Item", BindingFlags.Instance | BindingFlags.Public);
            _addMethod = dictCreateype.GetMethod("Add", BindingFlags.Instance | BindingFlags.Public);
        }

        internal void AddMappingValue(object key, object value)
        {
            _addMethod.Invoke(_valuemappings, new object[] { key, value });
        }

        public object GetMapingValue(object key)
        {
            if ((bool)_containKeyMethod.Invoke(_valuemappings, new object[] { key }))
            {
                return _getValueMethod.Invoke(_valuemappings, new object[] { key });
            }
            return null;
        }

    }
}

/******************************************************************
** auth: wei.huazhong
** date: 9/19/2018 6:01:08 PM
** desc:
******************************************************************/

using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelService
{
    public class ExcelImportService<T> where T : ImportEntityBase
    {
        /// <summary>
        /// will close stream after read data
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="errors"></param>
        /// <returns></returns>
        public IList<T> GetParsedPositionImport(Stream stream, IList<ImportError> errors, string cfgNodeName) =>
        (
             ParseImport(stream, cfgNodeName, errors)
        );

        private IList<T> ParseImport(Stream stream, string configName, IList<ImportError> errors)
        {
            var dataConfig = ExcelImportConfigHandler.Instance.GetExcelImportDataConfig(configName);

            using (var p = new ExcelPackage(stream))
            {
                p.Compatibility.IsWorksheets1Based = true;
                var sheet = p.Workbook.Worksheets[dataConfig.SheetIndex];

                var entityList = new List<T>();
                if (errors == null) errors = new List<ImportError>();

                T entity;

                int row = dataConfig.DataStartRow;

                IDictionary<string, List<T>> conflictData = new Dictionary<string, List<T>>();
                IDictionary<int, ImportError> conflictErrors = new Dictionary<int, ImportError>();

                while (true)
                {
                    string endColValue = sheet.Cells[row, dataConfig.CheckEndCol].GetValue<string>();
                    if (string.IsNullOrEmpty(dataConfig.CheckEndValue) && string.IsNullOrEmpty(endColValue))
                        break;
                    else
                    {
                        if (endColValue.Equals(dataConfig.CheckEndValue, StringComparison.InvariantCultureIgnoreCase))
                            break;
                    }

                    entity = Activator.CreateInstance<T>();
                    var accessor = TypeAccessor.Create(typeof(T));
                    entity.Line = row;
                    var columns = dataConfig.Columns;

                    foreach (var column in columns)
                    {
                        Type type = column.DataType;
                        var method = ReflectMethodProvider.Instance.GetCellValueMethod(type);
                        var result = method.Invoke(sheet.Cells[row, column.Col], null);// TODO: use delegate
                        string resultStr = Convert.ToString(result);

                        if (column.Required)
                        {
                            if (type == typeof(string))
                            {
                                if (string.IsNullOrEmpty(resultStr))
                                {
                                    var error = new ImportError { Line = row };
                                    error.ErrorMsg = ErrorMessageHandler.Instance.GetErrorMessage("EmptyColumnValue", column.Name);
                                    errors.Add(error);
                                    entity.IsError = true;
                                    continue;
                                }
                            }
                            else
                            {
                                if (result == null)
                                {
                                    var error = new ImportError { Line = row + 1 };
                                    error.ErrorMsg = ErrorMessageHandler.Instance.GetErrorMessage("InvalidDataFormatOrEmptyValue", column.Name);
                                    errors.Add(error);
                                    entity.IsError = true;
                                    continue;
                                }
                            }
                        }

                        if (column.MaxLength > 0 && result != null)
                        {
                            if (resultStr.Length > column.MaxLength)
                            {
                                errors.Add(new ImportError
                                {
                                    Line = row,
                                    ErrorMsg = ErrorMessageHandler.Instance.GetErrorMessage("OutOfLength", column.Name)
                                });

                                entity.IsError = true;
                                continue;
                            }
                        }

                        if (!string.IsNullOrEmpty(column.Regexp))
                        {
                            Regex rx = new Regex(column.Regexp);
                            if (!rx.IsMatch(resultStr))
                            {
                                errors.Add(new ImportError
                                {
                                    Line = row,
                                    ErrorMsg = ErrorMessageHandler.Instance.GetErrorMessage("DataFormatError", column.Name)
                                });

                                entity.IsError = true;
                                continue;
                            }
                        }

                        if (column.ValueMapping)
                        {
                            if (column.DataType == typeof(string))
                            {
                                if (result == null) result = string.Empty;
                                else
                                    result = (result as string).Trim().ToUpper();
                            }

                            var mappingValue = column.GetMapingValue(result);
                            if (mappingValue == null)
                            {
                                errors.Add(new ImportError
                                {
                                    Line = row,
                                    ErrorMsg = ErrorMessageHandler.Instance.GetErrorMessage("MappingKeyNotExists", column.Name, result)
                                });
                                entity.IsError = true;
                                continue;
                            }
                            else
                            {
                                accessor[entity, column.Name] = mappingValue;
                            }
                        }
                        else
                        {
                            accessor[entity, column.Name] = result;
                        }
                    }

                    entityList.Add(entity);
                    row++;
                }

                return entityList;
            }
        }
    }
}

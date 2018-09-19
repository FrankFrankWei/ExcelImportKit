/******************************************************************
** auth: wei.huazhong
** date: 9/17/2018 5:43:22 PM
** desc:
******************************************************************/

using ExcelService;
using ModelImport;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelImportKit
{
    public class SampleImportService
    {
        public IList<SampleImport> GetParsedPositionImport(Stream stream, IList<ImportError> errors)
        {
            var importList = ParseImport(stream, "Sample", errors);

            if (importList.Count <= 0) return null;

            FilterConflictData(importList, errors);

            if (errors.Count > 0) LogImportErrors();

            return importList;
        }

        private void LogImportErrors()
        { //(new ExportErrorDataService()).SaveErrorXlsFile(savedFile, errors, lines.Item1, lines.Item2);

        }


        private void FilterConflictData(IList<SampleImport> importList, IList<ImportError> errors)
        {
            //TODO: group importList by unique key, grouped count > 1 meanings multiple records, mark them as error
        }

        private IList<SampleImport> ParseImport(Stream stream, string configName, IList<ImportError> errors)
        {
            var dataConfig = new ExcelImportConfigHandler().GetExcelImportDataConfig(configName);

            using (var p = new ExcelPackage(stream))
            {
                p.Compatibility.IsWorksheets1Based = true;
                var sheet = p.Workbook.Worksheets[dataConfig.SheetIndex];

                var entityList = new List<SampleImport>();
                if (errors == null) errors = new List<ImportError>();

                SampleImport entity;

                int row = dataConfig.DataStartRow;

                IDictionary<string, List<SampleImport>> conflictData = new Dictionary<string, List<SampleImport>>();
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

                    entity = Activator.CreateInstance<SampleImport>();
                    entity.Line = row;
                    var columns = dataConfig.Columns;

                    foreach (var column in columns)
                    {
                        Type type = column.DataType;
                        var method = ReflectMethodProvider.Instance.GetCellValueMethod(type);
                        //var result = method.Invoke(sheet, new object[] { row, column.Col });
                        var result = method.Invoke(sheet.Cells[row, column.Col], null);
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
                                // TODO: fastmember optimize
                                column.PropertyInfo.SetValue(entity, mappingValue, null);
                            }
                        }
                        else
                        {
                            // TODO: fastmember optimize
                            column.PropertyInfo.SetValue(entity, result, null);
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
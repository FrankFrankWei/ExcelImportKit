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
        private IList<SampleImport> ParseImport(Stream stream, string configName, IList<ImportError> errors, Func<IDictionary<string, List<SampleImport>>, SampleImport, IDictionary<int, ImportError>, bool> checkConflictFunc)
        {
            var dataConfig = new ExcelImportConfigHandler().GetExcelImportDataConfig(configName);

            using (var p = new ExcelPackage(stream))
            {
                var sheet = p.Workbook.Worksheets[0];

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
                        var result = method.Invoke(sheet, new object[] { row, column.Col });
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

                        column.PropertyInfo.SetValue(entity, result, null);
                    }

                    if (checkConflictFunc != null && checkConflictFunc(conflictData, entity, conflictErrors))
                        entity.IsError = true;

                    entityList.Add(entity);
                    row++;
                }

                foreach (var conflictError in conflictErrors.Values)
                {
                    conflictError.FillConflictErrorMsg();
                    errors.Add(conflictError);
                }

                return entityList;
            }
        }

        public IList<SampleImport> GetParsedPositionImport(Stream stream, IList<ImportError> errors)
        {
            Func<IDictionary<string, List<SampleImport>>, SampleImport, IDictionary<int, ImportError>, bool> checkConflictFunc = (conflictData, talents, conflictErrors) =>
            {
                return CheckSampleImportDataConflict(conflictData, talents, conflictErrors);
            };

            var importList = ParseImport(stream, "Sample", errors, checkConflictFunc);

            if (importList.Count <= 0)
                return null;

            return importList;
        }

        private bool CheckSampleImportDataConflict(IDictionary<string, List<SampleImport>> conflictData, SampleImport comparedEntity, IDictionary<int, ImportError> conflictErrors)
        {
            return true;
        }
    }
}
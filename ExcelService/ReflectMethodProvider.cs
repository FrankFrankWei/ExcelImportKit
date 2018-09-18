/******************************************************************
** auth: wei.huazhong
** date: 9/17/2018 12:06:52 PM
** desc:
******************************************************************/

using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelService
{
    public class ReflectMethodProvider
    {
        private MethodInfo _getCellValueMethodMaker;
        private Dictionary<Type, MethodInfo> _getCellValueMethods;
        private MethodInfo _setCellValueMethodMaker;
        private Dictionary<Type, MethodInfo> _setCellValueMethods;

        private ReflectMethodProvider()
        {
            //_getCellValueMethodMaker = typeof(ExcelRange).GetMethod("GetValue", new Type[] { typeof(int), typeof(int) });
            _getCellValueMethodMaker = typeof(ExcelRange).GetMethod("GetValue", System.Type.EmptyTypes);
            _getCellValueMethods = new Dictionary<Type, MethodInfo>();

            _setCellValueMethodMaker = typeof(ExcelRange).GetMethod("SetValue", new Type[] { typeof(int), typeof(int), typeof(object), typeof(int) });
            _setCellValueMethods = new Dictionary<Type, MethodInfo>();
        }

        class Nested
        {
            static Nested() { }
            internal static readonly ReflectMethodProvider instance = new ReflectMethodProvider();
        }

        public static ReflectMethodProvider Instance
        {
            get
            {
                return Nested.instance;
            }
        }

        public MethodInfo GetCellValueMethod(Type dataType)
        {
            if (_getCellValueMethods.ContainsKey(dataType))
            {
                return _getCellValueMethods[dataType];
            }

            MethodInfo method = _getCellValueMethodMaker.MakeGenericMethod(dataType);
            _getCellValueMethods.Add(dataType, method);
            return method;
        }

        public MethodInfo GetSetCellValueMethod(Type dataType)
        {
            if (_setCellValueMethods.ContainsKey(dataType))
            {
                return _setCellValueMethods[dataType];
            }
            MethodInfo method = _setCellValueMethodMaker.MakeGenericMethod(dataType);
            _setCellValueMethods.Add(dataType, method);
            return method;
        }
    }
}

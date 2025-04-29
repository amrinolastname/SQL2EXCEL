using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data;
using System.Drawing;

namespace SQL2XL
{
    class excel
    {
        static DataGridView grid;

        static object xls;
        static object workbooks;
        static object workbook;
        static object sheet;
        static object range;
        static object cell1;
        static object cell2;
        static object borders;
        static object font;
        static object interior;
        static object entireColumn;
        static object style;
        static object activeWindow;
        static object activeSheet;

        private static void releaseInteropObject()
        {
            releaseObject(style);
            releaseObject(entireColumn);
            releaseObject(interior);
            releaseObject(font);
            releaseObject(borders);
            releaseObject(cell2);
            releaseObject(cell1);
            releaseObject(range);
            releaseObject(sheet);
            releaseObject(workbook);
            releaseObject(workbooks);
            releaseObject(xls);
            releaseObject(activeWindow);
            releaseObject(activeSheet);

            grid = null;

            GC.Collect();
        }
        private static int releaseObject(object obj)
        {
            if (obj != null) return System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            return 0;

        }

        public static void exportToExel(DataGridView inputGrid,bool withGraph=false)
        {
            grid = inputGrid;
            string error = "";
            try
            {
                if(grid.RowCount>0) startExport(withGraph);
            }
            catch (Exception e)
            {
                error = e.Message;
            }
            finally
            {
                releaseInteropObject();
                if (error != "") MessageBox.Show(error);
            }

        }

        private static void startExport(bool withGraph)
        {
            Type type;
            type = Type.GetTypeFromProgID("Excel.Application");
            xls = Activator.CreateInstance(type);

            workbooks = getObject(xls, "Workbooks", null);
            workbook = command(workbooks, "Add", null);
            sheet = getObject(workbook, "ActiveSheet", null);
            activeWindow = getObject(xls, "ActiveWindow", null);
            activeSheet = getObject(workbook,"ActiveSheet",null);

            style = getObject(cell(1, 1), "Style", null);

            setObject(xls, "Visible", true);
            setObject(xls, "UserControl", true);

            copyFromGrid(withGraph);

        }

        private static void copyFromGrid(bool withGraph)
        {
            int columns = grid.ColumnCount;
            int rowCount = grid.RowCount;
            int headerRowNumber;

            List<int> columnList = new List<int>();

            headerRowNumber = 3;


            List<object> textList = new List<object>();

            for (int column = 0; column < columns; column++)
            {
                if (grid.Columns[column].Visible)
                {
                    columnList.Add(column);
                    textList.Add(grid.Columns[column].HeaderText.Replace('\r', (char)10));
                }
            }

            columns = columnList.Count;
            object[] textArray = new object[columns];

            //====================================================
            // merge cell
            //----------------------------------------------------

            cell1 = cell(1, 1);
            cell2 = cell(columns, 1);
            range = setRange(cell1, cell2);

            command(range, "Merge", null);

            //====================================================
            // set font header
            //----------------------------------------------------

            setFont(cell1, "Size", 20);

            //====================================================
            // set border
            //----------------------------------------------------

            cell1 = cell(1, headerRowNumber);
            cell2 = cell(columns, rowCount-1 + headerRowNumber);
            range = setRange(cell1, cell2);
            borders = getObject(range, "Borders", null);
            setObject(borders, "Weight", 1);

            //====================================================

            for (int column = 0; column < columns; column++)
            {
                textArray[column] = textList[column];
            }
            cell1 = cell(1, headerRowNumber);
            cell2 = cell(columns, headerRowNumber);
            range = setRange(cell1, cell2);
            setObject(range, "VerticalAlignment", -4108);

            setFont(range, "Bold", true);
            setBgColor(range, 6);
            setObject(range, "WrapText", true);

            object[] parameter = new object[1];
            parameter[0] = textArray;

            writeToCell(range, parameter);

            for (int row = 0; row < rowCount-1; row++)
            {

                for (int column = 0; column < columns; column++)
                {
                    textArray[column] = grid[columnList[column], row].Value.ToString().Replace('\r', (char)10);
                }

                cell1 = cell(1, row + headerRowNumber + 1);
                cell2 = cell(columns, row + headerRowNumber + 1);
                range = setRange(cell1, cell2);
                writeToCell(range, parameter);

                writeToCell(1, 1, "Wait a moment, " + (Convert.ToDouble(row + 1) / Convert.ToDouble(rowCount) * 100d).ToString("0") + "% complete");

            }

            for (int column = 0; column < columns; column++)
            {
                entireColumn = getObject(cell(column + 1, 1), "EntireColumn", null);

                if (grid.Columns[columnList[column]].ValueType == typeof(decimal))
                {
                    setObject(entireColumn, "NumberFormat", "#,##0.00;-#,##0.00; ");

                }
                else if (grid.Columns[columnList[column]].ValueType == typeof(int))
                {
                    setObject(entireColumn, "NumberFormat", "0");
                }
                else if (grid.Columns[columnList[column]].ValueType == typeof(DateTime))
                {
                    setObject(entireColumn, "NumberFormat", "dd/MM/yyyy");
                    setObject(entireColumn, "HorizontalAlignment", -4108);
                }
                else
                {
                    setObject(entireColumn, "NumberFormat", "@");
                }

                command(entireColumn, "AutoFit", null);
            }

            range = setRange(cell(1, 1), cell(1, 2));
            range = getObject(range, "EntireRow", null);
            command(range, "Delete", -4162);
            command(cell(1, 1), "Select", null);
            //setObject(activeWindow, "FreezePanes", true);

            if (withGraph)
            {
                object shapes;
                object activeChart;
                object chart;
                shapes = getObject(sheet, "Shapes", null);
                activeChart = command(shapes, "AddChart", null);
                command(activeChart, "Select", null);
                chart = getObject(workbook, "ActiveChart", null);
                setObject(chart, "ChartType", -4102); //pie chart

            }
        }

        private static void writeToCell(object range, object[] value)
        {
            setObject(range, "Value", value);
        }
        private static void writeToCell(object column, object row, object value)
        {
            setObject(cell(column, row), "Value", value);
        }

        private static void setFont(object range, string properties, object parameter)
        {
            font = getObject(range, "Font", null);
            setObject(font, properties, parameter);
        }
        private static void setBgColor(object range, object color)
        {
            interior = getObject(range, "Interior", null);
            setObject(interior, "ColorIndex", color);
        }

        private static object setRange(object begin, object end)
        {
            object[] parameter;
            if (end == Missing.Value)
            {
                parameter = new object[1];
                parameter[0] = begin;
            }
            else
            {
                parameter = new object[2];
                parameter[0] = begin;
                parameter[1] = end;
            }
            return sheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, sheet, parameter);
        }
        private static object cell(object column, object row)
        {
            object[] parameter = new object[2];
            parameter[0] = row;
            parameter[1] = column;

            return sheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, sheet, parameter);
        }

        private static object command(object obj, string command, object single_parameter)
        {
            object[] parameter;
            if (single_parameter == null)
            {
                parameter = null;
            }
            else
            {
                parameter = new object[1];
                parameter[0] = single_parameter;
            }


            return obj.GetType().InvokeMember(command, BindingFlags.InvokeMethod, null, obj, parameter);
        }
        private static void setObject(object obj, string command, object single_parameter)
        {
            object[] parameter;
            if (single_parameter == null)
            {
                parameter = null;
            }
            else
            {
                parameter = new object[1];
                parameter[0] = single_parameter;
            }


            obj.GetType().InvokeMember(command, BindingFlags.SetProperty, null, obj, parameter);
        }
        private static void setObject(object obj, string command, object[] parameter)
        {
            obj.GetType().InvokeMember(command, BindingFlags.SetProperty, null, obj, parameter);
        }
        private static object getObject(object obj, string command, object[] single_parameter)
        {
            object[] parameter;
            if (single_parameter == null)
            {
                parameter = null;
            }
            else
            {
                parameter = new object[1];
                parameter[0] = single_parameter;
            }

            return obj.GetType().InvokeMember(command, BindingFlags.GetProperty, null, obj, parameter);
        }
    }
}

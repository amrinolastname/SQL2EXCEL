using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using System.ComponentModel;
using System.IO;

namespace SQL2XL
{
    class sql
    {
        public static string connectionString;
        public static SqlCommand command = new SqlCommand();
        public static SqlConnection connection = new SqlConnection();

        public static int readToGrid(string query, DataGridView grid, int alignment, int rowHeight, List<Type> newColumn)
        {

            int recordCount = 0;
            SqlConnection sqlConnection = new SqlConnection(connectionString);
            SqlDataReader sqlData = null;
            grid.DataSource = null;

            int columnCount = 0;
            DataTable result = new DataTable();

            try
            {

                sqlConnection.Open();
                command.Connection = sqlConnection;

                command.CommandText = "set dateformat 'dmy';" + query;
                sqlData = command.ExecuteReader();
                if (sqlData.HasRows == true)
                {
                    //grid.DataSource = null;

                    columnCount = sqlData.FieldCount;
                    object[] content = new object[columnCount];
                    for (int xx = 0; xx < columnCount; xx++)
                    {
                        result.Columns.Add(Convert.ToString(xx), sqlData.GetFieldType(xx));
                    }
                    recordCount = 0;
                    while (sqlData.Read())
                    {
                        recordCount = ++recordCount;
                        for (int xx = 0; xx < columnCount; xx++)
                        {
                            content[xx] = sqlData[xx];
                        }

                        result.Rows.Add(content);
                    }

                    if (newColumn != null)
                    {
                        for (int xx = 0; xx < newColumn.Count; xx++)
                        {
                            result.Columns.Add(result.Columns.Count.ToString(), newColumn[xx]); ;
                        }
                    }

                    if (rowHeight > 0) grid.RowTemplate.Height = rowHeight;

                    grid.RowTemplate.MinimumHeight = grid.RowTemplate.Height;
                    grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

                    grid.DataSource = result;

                    DataGridViewContentAlignment numberAlignment, stringAlignment;
                    if (alignment == 0)
                    {
                        numberAlignment = DataGridViewContentAlignment.MiddleRight;
                        stringAlignment = DataGridViewContentAlignment.MiddleLeft;
                    }
                    else
                    {
                        numberAlignment = DataGridViewContentAlignment.TopRight;
                        stringAlignment = DataGridViewContentAlignment.TopLeft;
                    }

                    grid.DefaultCellStyle.Alignment = stringAlignment;

                    for (int xx = 0; xx < columnCount; xx++)
                    {
                        string aa = sqlData.GetFieldType(xx).ToString();
                        grid.Columns[xx].HeaderText = sqlData.GetName(xx);
                        if (sqlData.GetFieldType(xx) == typeof(decimal))
                        {
                            grid.Columns[xx].DefaultCellStyle.Format = "#,##0.00;-#,##0.00; ";
                            grid.Columns[xx].DefaultCellStyle.Alignment = numberAlignment;
                        }
                        else if (sqlData.GetFieldType(xx) == typeof(int))
                        {
                            grid.Columns[xx].DefaultCellStyle.Alignment = numberAlignment;
                        }
                        else if (sqlData.GetFieldType(xx) == typeof(DateTime))
                        {
                            grid.Columns[xx].DefaultCellStyle.Format = "dd/MM/yyyy";
                        }
                    }

                }
            }

            catch (Exception e)
            {
                MessageBox.Show("Error: " + e.Message);
            }

            try
            {
                sqlData.Close();
                sqlData.Dispose();
            }
            catch
            {
            }

            sqlConnection.Close();
            sqlConnection.Dispose();


            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            int a = grid.ColumnHeadersHeight;
            grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            grid.ColumnHeadersHeight = a + 10;

            grid.AutoResizeColumns();

            return recordCount;
        }
    }


}

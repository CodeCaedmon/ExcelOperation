﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using System.Data.OleDb;

namespace Caedmon
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

        }

        private void InBtn_Click(object sender, RoutedEventArgs e)
        {
            string Path = "";
            Path =@PathText.Text+ NameText.Text+".xlsx";
            DataSet ds = new DataSet();
            ExcelOperation eco = new ExcelOperation();
            ds = eco.ExcelToDS(Path);
            DataTable dt = new DataTable();
            dt = ds.Tables[0];
            DataGrid1.ItemsSource = dt.DefaultView;
        }
    }
}

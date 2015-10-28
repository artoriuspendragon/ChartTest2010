using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Windows.Forms.DataVisualization.Charting;
using Tools.Excel;
using FundamentalClass;

namespace ChartTest2010
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            chart1.Series.Clear();            
        }

        //图的不同类型表：
        //Point  点图类型。  
        // FastPoint  快速点图类型。  
        // Bubble  气泡图类型。  
        // Line  折线图类型。  
        // Spline  样条图类型。  
        // StepLine  阶梯线图类型。  
        // FastLine  快速扫描线图类型。  
        // Bar  条形图类型。  
        // StackedBar  堆积条形图类型。  
        // StackedBar100  百分比堆积条形图类型。  
        // Column  柱形图类型。  
        // StackedColumn  堆积柱形图类型。  
        // StackedColumn100  百分比堆积柱形图类型。  
        // Area  面积图类型。  
        // SplineArea  样条面积图类型。  
        // StackedArea  堆积面积图类型。  
        // StackedArea100  百分比堆积面积图类型。  
        // Pie  饼图类型。  
        // Doughnut  圆环图类型。  
        // Stock  股价图类型。  
        // Candlestick  K 线图类型。  
        // Range  范围图类型。  
        // SplineRange  样条范围图类型。  
        // RangeBar  范围条形图类型。  
        // RangeColumn  范围柱形图类型。  
        // Radar  雷达图类型。  
        // Polar  极坐标图类型。  
        // ErrorBar  误差条形图类型。  
        // BoxPlot  盒须图类型。  
        // Renko  砖形图类型。  
        // ThreeLineBreak  新三值图类型。  
        // Kagi  卡吉图类型。  
        // PointAndFigure  点数图类型。  
        // Funnel  漏斗图类型。  
        // Pyramid  棱锥图类型。 
        private void AddIntoChart(clsChartPoint[] chartPoints, Chart chart)
        {
            Series series = new Series("生猪价格");
            series.ChartType = SeriesChartType.Line;//设置画图类型
            series.BorderWidth = 7;
            series.ShadowOffset = 2;

            for (int i = 0; i < chartPoints.Length; i++)
            {
                DataPoint dp = new DataPoint();
                dp.Label = chartPoints[i].PointValue.ToString();
                dp.AxisLabel = chartPoints[i].PointName;
                dp.YValues[0] = chartPoints[i].PointValue;
                series.Points.Add(dp);
            }

            chart.Series.Add(series);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            clsExcelReader er = new clsExcelReader();

            //打开文件对话框
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "c:\\";//初始路径
            openFileDialog.Filter = "文本文件|*.*|C#文件|*.cs|所有文件|*.*";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                er.FileName = openFileDialog.FileName;
            }
            //

            textBox1.Text = er.FileName;
            er.SheetNumber = 1;

            if (er.OpenFileContinuously() == false)
            {
                textBox1.Text = er.ErrorString;
                return;
            }

            int iCount = er.RowCount;
            
            string[] listMonth = new string[iCount - 2];
            clsChartPoint[] listPrice = new clsChartPoint[iCount - 3];

            for (int i = 0; i < iCount - 3; i++)
            {
                listPrice[i] = new clsChartPoint();
                //列号是指定的
                listPrice[i].PointValue = Convert.ToDouble(er.getTextInOneCell(i + 3, 3));
                listPrice[i].PointName = er.getTextInOneCell(i + 3, 1);                
            }

            er.CloseFile();

            this.AddIntoChart(listPrice, chart1);

            //设置Chart的滚动条:
            //设置滚动条是在外部显示
            chart1.ChartAreas["ChartArea1"].AxisX.ScrollBar.IsPositionedInside= false;
            //设置滚动条的宽度
            chart1.ChartAreas["ChartArea1"].AxisX.ScrollBar.Size = 20;
            //滚动条只显示向前的按钮，主要是为了不显示取消显示的按钮
            chart1.ChartAreas["ChartArea1"].AxisX.ScrollBar.ButtonStyle = ScrollBarButtonStyles.SmallScroll;
            //设置图表可视区域数据点数，说白了一次可以看到多少个X轴区域
            chart1.ChartAreas["ChartArea1"].AxisX.ScaleView.Size = 10;
            //设置滚动一次，移动几格区域
            chart1.ChartAreas["ChartArea1"].AxisX.ScaleView.MinSize = 2;
            //设置X轴的间隔，设置它是为了看起来方便点，也就是要每个X轴的记录都显示出来
            chart1.ChartAreas["ChartArea1"].AxisX.Interval=2;
            //X轴起始点
            chart1.ChartAreas["ChartArea1"].AxisX.Minimum = 0;
            //X轴结束点，一般这个是应该在后台设置的，
            chart1.ChartAreas["ChartArea1"].AxisX.Maximum = 19;
            //对于我而言，是用的第一列作为X轴，那么有多少行，就有多少个X轴的刻度，所以最大值应该就等于行数；
            //该值设置大了，会在后边出现一推空白，设置小了，会出后边多出来的数据在图表中不显示，所以最好是在后台根据你的数据列来设置.要实现显示滚动条，就不能设置成自动显示刻度，必须要有值才可以。
        }
    }
}

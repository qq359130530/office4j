package com.office.chart4j.core;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtils;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.StandardChartTheme;
import org.jfree.chart.labels.StandardPieSectionLabelGenerator;
import org.jfree.chart.plot.PiePlot;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;

import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.NumberFormat;

public class Test {

    public static void main(String[] args) throws IOException {
        // 创建主题样式
        StandardChartTheme theme = new StandardChartTheme("defaultTheme");
        // 设置标题字体
        theme.setExtraLargeFont(new Font("SimSun", Font.BOLD,20));
        // 设置图例字体
        theme.setRegularFont(new Font("SimSun", Font.PLAIN,15));
        // 设置轴向字体
        theme.setLargeFont(new Font("SimSun", Font.PLAIN,15));
        // 应用主题样式
        ChartFactory.setChartTheme(theme);
        OutputStream outputStream = new FileOutputStream(new File("C:\\Users\\Rzxuser\\Desktop\\chart.jpg"));
        ChartUtils.writeChartAsJPEG(outputStream, createBarChart(), 500, 500);
        outputStream.close();
    }

    // 条形图
    public static JFreeChart createBarChart() {
        DefaultCategoryDataset dateset = new DefaultCategoryDataset();
        dateset.addValue(1.0, "宝马", "速度");
        dateset.addValue(2.0, "奥迪", "速度");
        dateset.addValue(3.0, "奔驰", "速度");
        dateset.addValue(2.0, "宝马", "价格");
        dateset.addValue(2.0, "奥迪", "价格");
        dateset.addValue(1.0, "奔驰", "价格");
        dateset.addValue(2.0, "宝马", "空间");
        dateset.addValue(2.0, "奥迪", "空间");
        dateset.addValue(3.0, "奔驰", "空间");

        JFreeChart chart = ChartFactory.createBarChart("汽车对照表", "参数", "分值", dateset);
        return chart;
    }

    // 饼图
    public static JFreeChart createPieChart() {
        DefaultPieDataset dataset = new DefaultPieDataset();
        dataset.setValue("iPhone 12", 200);
        dataset.setValue("三星", 240);
        dataset.setValue("华为", 500);
        dataset.setValue("OPPO", 150);
        JFreeChart chart = ChartFactory.createPieChart("手机销售份额", dataset);
        // 设置背景颜色
        chart.setBackgroundPaint(Color.WHITE);
        PiePlot plot = (PiePlot) chart.getPlot();
        // 设置百分比显示格式
        plot.setLabelGenerator(
                new StandardPieSectionLabelGenerator(
                        "{0}={1}({2})",
                        NumberFormat.getNumberInstance(),
                        new DecimalFormat("0.00%")
                )
        );
        // 设置section轮廓线颜色
        plot.setDefaultSectionOutlinePaint(new Color(0xF7, 0x79,0xED));
        // 设置section轮廓线厚度
        plot.setDefaultSectionOutlineStroke(new BasicStroke(0));
        // 设置section颜色
        plot.setDefaultSectionPaint(new Color(0xF7, 0x79, 0xED));
        plot.setNoDataMessage("暂无数据！");
        plot.setNoDataMessagePaint(Color.blue);
        plot.setCircular(true);
        plot.setLabelGap(0.01D);// 间距
        plot.setBackgroundPaint(Color.white);
        plot.setLabelFont(new Font("SimSun", Font.TRUETYPE_FONT,12));
        // 设置背景透明度（0~1）
        plot.setBackgroundAlpha(0.6f);
        // 设置前景透明度（0~1）
        plot.setForegroundAlpha(0.8f);
        return chart;
    }

}

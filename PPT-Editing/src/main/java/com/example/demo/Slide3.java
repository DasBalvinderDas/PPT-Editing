package com.example.demo;

import java.awt.Color;
import java.awt.Rectangle;
import java.awt.geom.Rectangle2D;
import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.sl.usermodel.TableCell.BorderEdge;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.AxisCrossBetween;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.AxisTickMark;
import org.apache.poi.xddf.usermodel.chart.BarDirection;
import org.apache.poi.xddf.usermodel.chart.BarGrouping;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFBarChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFGraphicFrame;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

/**
 * Build a chart without reading template file
 */
@SuppressWarnings({"java:S106","java:S4823","java:S1192"})
public final class Slide3 {
    private Slide3() {}

    private static void usage(){
        System.out.println("Usage: DoughnutChartFromScratch <bar-chart-data.txt>");
        System.out.println("    bar-chart-data.txt          the model to set. First line is chart title, " +
                "then go pairs {axis-label value}");
    }

    public static void main(String[] args) throws Exception {
        if(args.length < 1) {
            usage();
            return;
        }
        
        
        try (BufferedReader modelReader = Files.newBufferedReader(Paths.get(args[0]), StandardCharsets.UTF_8)) {

            String chartTitle = modelReader.readLine();  // first line is chart title
            String seriesText = modelReader.readLine();
            String[] series = seriesText == null ? new String[0] : seriesText.split(",");

            // Category Axis Data
            List<String> listLanguages = new ArrayList<>(10);

            // Values
            List<Double> listCountries = new ArrayList<>(10);
            List<Double> listSpeakers = new ArrayList<>(10);

            // set model
            String ln;
            while((ln = modelReader.readLine()) != null) {
                String[] vals = ln.split(",");
                listCountries.add(Double.valueOf(vals[0]));
                listSpeakers.add(Double.valueOf(vals[1]));
                listLanguages.add(vals[2]);
            }

            String[] categories = listLanguages.toArray(new String[0]);
            Double[] values1 = listCountries.toArray(new Double[0]);
            Double[] values2 = listSpeakers.toArray(new Double[0]);

            try (XMLSlideShow ppt = new XMLSlideShow()) {
                createSlideWithChart(ppt, chartTitle, series, categories, values1, COLUMN_COUNTRIES);
                //createSlideWithChart(ppt, chartTitle, series, categories, values2, COLUMN_SPEAKERS);
                // save the result
                try (OutputStream out = new FileOutputStream("slide3.pptx")) {
                    ppt.write(out);
                }
            }
            try (FileInputStream is = new FileInputStream("slide3.pptx")) {
                try (XMLSlideShow ppt = new XMLSlideShow(is)) {
                    for (XSLFSlide slide : ppt.getSlides()) {
                        for (XSLFShape shape : slide.getShapes()) {
                            if (shape instanceof XSLFGraphicFrame) {
                                XSLFGraphicFrame frame = (XSLFGraphicFrame) shape;
                                if (frame.hasChart()) {
                                    System.out.println(frame.getChart().getTitleShape().getText());
                                }
                            }
                        }
                    }
                }
            }
        }
        System.out.println("Done");
    }

    private static void createSlideWithChart(XMLSlideShow ppt, String chartTitle, String[] series, String[] categories,
                                             Double[] values, int valuesColumn) throws IOException{
    	
    	ppt.setPageSize(new java.awt.Dimension(1920, 1080));
    	
        XSLFSlide slide = ppt.createSlide();
        XSLFTextBox topHeading = slide.createTextBox();
    	
        //Top Heading
        String headindText = "High level findings";
        PPTUtil.topHeading(topHeading,new Rectangle(10, 50, 550, 100),new Color(51,51,51),headindText);

        //Top Para
        String introParaText = "This section represents the high-level findings we summarized from across your in-scope environments. You may access additional information and analytics by logging into your HCLVertex® portal account.";
        PPTUtil.createParagraph(slide,new Rectangle(10, 100, 1900, 100),introParaText,30d,-20d,100d);
        
        //Pie Chart
        XSLFChart chart = ppt.createChart();
        Rectangle2D rect2D = new Rectangle(fromCM(2), fromCM(9), fromCM(20), fromCM(10));
        slide.addChart(chart, rect2D);
        setDoughnutData(chart, chartTitle, series, categories, values, valuesColumn);
        
        XSLFChart chart2 = ppt.createChart();
        Rectangle2D rect22D = new Rectangle(fromCM(22), fromCM(9), fromCM(20), fromCM(10));
        slide.addChart(chart2, rect22D);
        setDoughnutData(chart2, chartTitle, series, categories, values, valuesColumn);
        
 //---------------------------Bar Chart ------------------------------
        
        


        String chartTitleBar = "Bar Chart";  // first line is chart title

        // Category Axis Data
        List<String> listLanguages = new ArrayList<>(10);

        // Values
        List<Double> listCountries = new ArrayList<>(10);
        List<Double> listSpeakers = new ArrayList<>(10);

        // set model
		listCountries.add(Double.valueOf(4));
		listCountries.add(Double.valueOf(38));
		listCountries.add(Double.valueOf(118));
		listCountries.add(Double.valueOf(4));
		listCountries.add(Double.valueOf(2));
		listCountries.add(Double.valueOf(15));
		listCountries.add(Double.valueOf(6));
		listCountries.add(Double.valueOf(18));
		listCountries.add(Double.valueOf(31));

		listSpeakers.add(Double.valueOf(243));
		listSpeakers.add(Double.valueOf(1219));
		listSpeakers.add(Double.valueOf(378));
		listSpeakers.add(Double.valueOf(260));
		listSpeakers.add(Double.valueOf(128));
		listSpeakers.add(Double.valueOf(223));
		listSpeakers.add(Double.valueOf(119));
		listSpeakers.add(Double.valueOf(154));
		listSpeakers.add(Double.valueOf(442));

		listLanguages.add("বাংলা");
		listLanguages.add("中文");
		listLanguages.add("English");
		listLanguages.add("हिन्दी");
		listLanguages.add("日本語");
		listLanguages.add("português");
		listLanguages.add("ਪੰਜਾਬੀ");
		listLanguages.add("Русский язык");
		listLanguages.add("español");

        String[] categoriesBar = listLanguages.toArray(new String[0]);
        Double[] values1 = listCountries.toArray(new Double[0]);
        Double[] values2 = listSpeakers.toArray(new Double[0]);

        try {
            createSlideWithChartBar(slide,ppt, chartTitleBar, series, categoriesBar, values1, values2);
            // save the result
        }catch(Exception e) {
        	System.out.println(e);
        }
//        try (FileInputStream is = new FileInputStream("chart-from-scratch.pptx")) {
//            try (XMLSlideShow ppt = new XMLSlideShow(is)) {
//                for (XSLFSlide slide : ppt.getSlides()) {
//                    for (XSLFShape shape : slide.getShapes()) {
//                        if (shape instanceof XSLFGraphicFrame) {
//                            XSLFGraphicFrame frame = (XSLFGraphicFrame) shape;
//                            if (frame.hasChart()) {
//                                System.out.println(frame.getChart().getTitleShape().getText());
//                            }
//                        }
//                    }
//                }
//            }
//        }
    
        
        
// ------------------------------Bar chart ends-------------------------------------        
        
      //---------------------------------Create Table ----------------------------------------
		
      		XSLFTable tbl = slide.createTable();
      		tbl.setAnchor(new Rectangle(100, 600, 1000, 300));

              int numColumns = 3;
              int numRows = 2;
              XSLFTableRow headerRow = tbl.addRow();
              headerRow.setHeight(50);
              // header
              for (int i = 0; i < numColumns; i++) {
                  XSLFTableCell th = headerRow.addCell();
                  XSLFTextParagraph p = th.addNewTextParagraph();
                  p.setTextAlign(TextAlign.CENTER);
                  XSLFTextRun r = p.addNewTextRun();
                  if(i==0)
                	  r.setText("Assets ");
                  r.setBold(true);
                  r.setFontColor(Color.white);
                  th.setFillColor(new Color(79, 129, 189));
                  th.setBorderWidth(BorderEdge.bottom, 2.0);
                  th.setBorderColor(BorderEdge.bottom, Color.white);

                  tbl.setColumnWidth(i, 250);  // all columns are equally sized
              }

              // rows

              for (int rownum = 0; rownum < numRows; rownum++) {
                  XSLFTableRow tr = tbl.addRow();
                  tr.setHeight(50);
                  // header
                  for (int i = 0; i < numColumns; i++) {
                      XSLFTableCell cell = tr.addCell();
                      XSLFTextParagraph p = cell.addNewTextParagraph();
                      XSLFTextRun r = p.addNewTextRun();

                      r.setText("Cell " + (i + 1));
                      if (rownum % 2 == 0)
                          cell.setFillColor(new Color(208, 216, 232));
                      else
                          cell.setFillColor(new Color(233, 247, 244));

                  }
              }

              //--------------------------------------------------------------------------------------
              
              
            //---------------------------------Create Table 2----------------------------------------
      		
        		XSLFTable tbl2 = slide.createTable();
        		tbl2.setAnchor(new Rectangle(900, 600, 1000, 300));

                int numColumns2 = 3;
                int numRows2 = 2;
                XSLFTableRow headerRow2 = tbl2.addRow();
                headerRow2.setHeight(50);
                // header
                for (int i = 0; i < numColumns2; i++) {
                    XSLFTableCell th = headerRow2.addCell();
                    XSLFTextParagraph p = th.addNewTextParagraph();
                    p.setTextAlign(TextAlign.CENTER);
                    XSLFTextRun r = p.addNewTextRun();
                    if(i==0)
                        r.setText("HCLVertex Fit Score");
                    r.setBold(true);
                    r.setFontColor(Color.white);
                    th.setFillColor(new Color(79, 129, 189));
                    th.setBorderWidth(BorderEdge.bottom, 2.0);
                    th.setBorderColor(BorderEdge.bottom, Color.white);

                    tbl2.setColumnWidth(i, 250);  // all columns are equally sized
                }

                // rows

                for (int rownum = 0; rownum < numRows2; rownum++) {
                    XSLFTableRow tr = tbl2.addRow();
                    tr.setHeight(50);
                    // header
                    for (int i = 0; i < numColumns; i++) {
                        XSLFTableCell cell = tr.addCell();
                        XSLFTextParagraph p = cell.addNewTextParagraph();
                        XSLFTextRun r = p.addNewTextRun();

                        r.setText("Cell " + (i + 1));
                        if (rownum % 2 == 0)
                            cell.setFillColor(new Color(208, 216, 232));
                        else
                            cell.setFillColor(new Color(233, 247, 244));

                    }
                }

                //--------------------------------------------------------------------------------------
                
                
              //---------------------------------Create Table 3----------------------------------------
          		
        		XSLFTable tbl3 = slide.createTable();
        		tbl3.setAnchor(new Rectangle(100, 760, 1000, 300));

                int numColumns3 = 1;
                int numRows3 = 1;
                XSLFTableRow headerRow3 = tbl3.addRow();
                headerRow3.setHeight(50);
                // header
                for (int i = 0; i < numColumns3; i++) {
                    XSLFTableCell th = headerRow3.addCell();
                    XSLFTextParagraph p = th.addNewTextParagraph();
                    p.setTextAlign(TextAlign.CENTER);
                    XSLFTextRun r = p.addNewTextRun();
                    r.setText("CPU Cores");
                    r.setBold(true);
                    r.setFontColor(Color.white);
                    th.setFillColor(new Color(79, 129, 189));
                    th.setBorderWidth(BorderEdge.bottom, 2.0);
                    th.setBorderColor(BorderEdge.bottom, Color.white);

                    tbl3.setColumnWidth(i, 750);  // all columns are equally sized
                }

                // rows

                for (int rownum = 0; rownum < numRows3; rownum++) {
                    XSLFTableRow tr = tbl3.addRow();
                    tr.setHeight(50);
                    // header
                    for (int i = 0; i < numColumns3; i++) {
                        XSLFTableCell cell = tr.addCell();
                        XSLFTextParagraph p = cell.addNewTextParagraph();
                        XSLFTextRun r = p.addNewTextRun();

                        r.setText("Cell " + (i + 1));
                        if (rownum % 2 == 0)
                            cell.setFillColor(new Color(208, 216, 232));
                        else
                            cell.setFillColor(new Color(233, 247, 244));

                    }
                }

                //--------------------------------------------------------------------------------------
                
                
              //---------------------------------Create Table 4----------------------------------------
          		
        		XSLFTable tbl4 = slide.createTable();
        		tbl4.setAnchor(new Rectangle(900, 760, 1000, 300));

                int numColumns4 = 1;
                int numRows4 = 1;
                XSLFTableRow headerRow4 = tbl4.addRow();
                headerRow4.setHeight(50);
                // header
                for (int i = 0; i < numColumns4; i++) {
                    XSLFTableCell th = headerRow4.addCell();
                    XSLFTextParagraph p = th.addNewTextParagraph();
                    p.setTextAlign(TextAlign.CENTER);
                    XSLFTextRun r = p.addNewTextRun();
                    r.setText("Installed Software");
                    r.setBold(true);
                    r.setFontColor(Color.white);
                    th.setFillColor(new Color(79, 129, 189));
                    th.setBorderWidth(BorderEdge.bottom, 2.0);
                    th.setBorderColor(BorderEdge.bottom, Color.white);

                    tbl4.setColumnWidth(i, 750);  // all columns are equally sized
                }

                // rows

                for (int rownum = 0; rownum < numRows4; rownum++) {
                    XSLFTableRow tr = tbl4.addRow();
                    tr.setHeight(50);
                    // header
                    for (int i = 0; i < numColumns4; i++) {
                        XSLFTableCell cell = tr.addCell();
                        XSLFTextParagraph p = cell.addNewTextParagraph();
                        XSLFTextRun r = p.addNewTextRun();

                        r.setText("Cell " + (i + 1));
                        if (rownum % 2 == 0)
                            cell.setFillColor(new Color(208, 216, 232));
                        else
                            cell.setFillColor(new Color(233, 247, 244));

                    }
                }

                //--------------------------------------------------------------------------------------
                
                //---------------------------------Create Table 5----------------------------------------
          		
          		XSLFTable tbl5 = slide.createTable();
          		tbl5.setAnchor(new Rectangle(100, 870, 1000, 300));

                  int numColumns5 = 1;
                  int numRows5 = 1;
                  XSLFTableRow headerRow5 = tbl5.addRow();
                  headerRow5.setHeight(50);
                  // header
                  for (int i = 0; i < numColumns5; i++) {
                      XSLFTableCell th = headerRow5.addCell();
                      XSLFTextParagraph p = th.addNewTextParagraph();
                      p.setTextAlign(TextAlign.CENTER);
                      XSLFTextRun r = p.addNewTextRun();
                      r.setText("CPU Cores");
                      r.setBold(true);
                      r.setFontColor(Color.white);
                      th.setFillColor(new Color(79, 129, 189));
                      th.setBorderWidth(BorderEdge.bottom, 2.0);
                      th.setBorderColor(BorderEdge.bottom, Color.white);

                      tbl5.setColumnWidth(i, 750);  // all columns are equally sized
                  }

                  // rows

                  for (int rownum = 0; rownum < numRows5; rownum++) {
                      XSLFTableRow tr = tbl5.addRow();
                      tr.setHeight(50);
                      // header
                      for (int i = 0; i < numColumns5; i++) {
                          XSLFTableCell cell = tr.addCell();
                          XSLFTextParagraph p = cell.addNewTextParagraph();
                          XSLFTextRun r = p.addNewTextRun();

                          r.setText("Cell " + (i + 1));
                          if (rownum % 2 == 0)
                              cell.setFillColor(new Color(208, 216, 232));
                          else
                              cell.setFillColor(new Color(233, 247, 244));

                      }
                  }

                  //--------------------------------------------------------------------------------------
                  
                  
                //---------------------------------Create Table 6----------------------------------------
            		
          		XSLFTable tbl6 = slide.createTable();
          		tbl6.setAnchor(new Rectangle(900, 870, 1000, 300));

                  int numColumns6 = 1;
                  int numRows6 = 1;
                  XSLFTableRow headerRow6 = tbl6.addRow();
                  headerRow6.setHeight(50);
                  // header
                  for (int i = 0; i < numColumns6; i++) {
                      XSLFTableCell th = headerRow6.addCell();
                      XSLFTextParagraph p = th.addNewTextParagraph();
                      p.setTextAlign(TextAlign.CENTER);
                      XSLFTextRun r = p.addNewTextRun();
                      r.setText("Installed Software");
                      r.setBold(true);
                      r.setFontColor(Color.white);
                      th.setFillColor(new Color(79, 129, 189));
                      th.setBorderWidth(BorderEdge.bottom, 2.0);
                      th.setBorderColor(BorderEdge.bottom, Color.white);

                      tbl6.setColumnWidth(i, 750);  // all columns are equally sized
                  }

                  // rows

                  for (int rownum = 0; rownum < numRows6; rownum++) {
                      XSLFTableRow tr = tbl6.addRow();
                      tr.setHeight(50);
                      // header
                      for (int i = 0; i < numColumns6; i++) {
                          XSLFTableCell cell = tr.addCell();
                          XSLFTextParagraph p = cell.addNewTextParagraph();
                          XSLFTextRun r = p.addNewTextRun();

                          r.setText("Cell " + (i + 1));
                          if (rownum % 2 == 0)
                              cell.setFillColor(new Color(208, 216, 232));
                          else
                              cell.setFillColor(new Color(233, 247, 244));

                      }
                  }

                  //--------------------------------------------------------------------------------------
                
                //Footer
                PPTUtil.footerImage(ppt, slide);
                
                String footerTextNote = "Copyright © 2022 HCL Technologies Limited | www.hcltech.com";
                PPTUtil.createParagraph(slide,new Rectangle(1300, 1000, 700, 100),footerTextNote,20d,-20d,100d);
        
    }

    private static int fromCM(double cm) {
        return (int) (Math.rint(cm * Units.EMU_PER_CENTIMETER));
    }

    private static void setDoughnutData(XSLFChart chart, String chartTitle, String[] series, String[] categories,
                                        Double[] values, int valuesColumn) {
        final int numOfPoints = categories.length;
        final String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, COLUMN_LANGUAGES, COLUMN_LANGUAGES));
        final String valuesDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, valuesColumn, valuesColumn));
        final XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, COLUMN_LANGUAGES);
        final XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(values, valuesDataRange, valuesColumn);
        valuesData.setFormatCode("General");

        XDDFDoughnutChartData data = (XDDFDoughnutChartData) chart.createData(ChartTypes.DOUGHNUT, null, null);
        XDDFDoughnutChartData.Series series1 = (XDDFDoughnutChartData.Series) data.addSeries(categoriesData, valuesData);
        series1.setTitle(series[0], chart.setSheetTitle(series[valuesColumn - 1], valuesColumn));

        data.setVaryColors(true);
        //data.setHoleSize(42);
        //data.setFirstSliceAngle(90);
        chart.plot(data);

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.LEFT);
        legend.setOverlay(false);

        chart.setTitleText(chartTitle);
        chart.setTitleOverlay(false);
        chart.setAutoTitleDeleted(false);
        
    }

    private static final int COLUMN_LANGUAGES = 0;
    private static final int COLUMN_COUNTRIES = 1;
    private static final int COLUMN_SPEAKERS = 2;
    
    //------------------------------- Bar chart Additional Methods ---------------------------------------
    
    private static void createSlideWithChartBar(XSLFSlide slide,XMLSlideShow ppt, String chartTitle, String[] series, String[] categories,
            Double[] values1, Double[] values2) {
        XSLFChart chart = ppt.createChart();
        Rectangle2D rect2D = new Rectangle(fromCM(44), fromCM(9), fromCM(15), fromCM(10));
        slide.addChart(chart, rect2D);
        setBarData(chart, chartTitle, series, categories, values1, values2);
    }

    private static void setBarData(XSLFChart chart, String chartTitle, String[] series, String[] categories, Double[] values1, Double[] values2) {
        // Use a category axis for the bottom axis.
        XDDFChartAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(series[2]);
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(series[0]+","+series[1]);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        leftAxis.setMajorTickMark(AxisTickMark.OUT);
        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

        final int numOfPoints = categories.length;
        final String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, COLUMN_LANGUAGES, COLUMN_LANGUAGES));
        final String valuesDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, COLUMN_COUNTRIES, COLUMN_COUNTRIES));
        final String valuesDataRange2 = chart.formatRange(new CellRangeAddress(1, numOfPoints, COLUMN_SPEAKERS, COLUMN_SPEAKERS));
        final XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, COLUMN_LANGUAGES);
        final XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(values1, valuesDataRange, COLUMN_COUNTRIES);
        valuesData.setFormatCode("General");
        values1[6] = 16.0; // if you ever want to change the underlying data, it has to be done before building the data source
        final XDDFNumericalDataSource<? extends Number> valuesData2 = XDDFDataSourcesFactory.fromArray(values2, valuesDataRange2, COLUMN_SPEAKERS);
        valuesData2.setFormatCode("General");


        XDDFBarChartData bar = (XDDFBarChartData) chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
        bar.setBarGrouping(BarGrouping.CLUSTERED);

        XDDFBarChartData.Series series1 = (XDDFBarChartData.Series) bar.addSeries(categoriesData, valuesData);
        series1.setTitle(series[0], chart.setSheetTitle(series[COLUMN_COUNTRIES - 1], COLUMN_COUNTRIES));

        XDDFBarChartData.Series series2 = (XDDFBarChartData.Series) bar.addSeries(categoriesData, valuesData2);
        series2.setTitle(series[1], chart.setSheetTitle(series[COLUMN_SPEAKERS - 1], COLUMN_SPEAKERS));

        bar.setVaryColors(true);
        bar.setBarDirection(BarDirection.COL);
        chart.plot(bar);

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.LEFT);
        legend.setOverlay(false);

        chart.setTitleText(chartTitle);
        chart.setTitleOverlay(false);
        chart.setAutoTitleDeleted(false);
    }

}
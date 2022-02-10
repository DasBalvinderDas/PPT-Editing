package com.example.demo;

import java.awt.Color;
import java.awt.Rectangle;
import java.awt.geom.Rectangle2D;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.sl.usermodel.TableCell.BorderEdge;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFDoughnutChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

public class Master {

	private static final int COLUMN_LANGUAGES = 0;
	private static final int COLUMN_COUNTRIES = 1;
	private static final int COLUMN_SPEAKERS = 2;

	public static void main(String[] args) throws FileNotFoundException, IOException {

		// Read existing slide
		XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("Assessment.pptx"));

		System.out.println(ppt.getSlides().size());

		// ppt.getSlides().get(0);

		// create a new slide show which copy few slides from Assessment slide
		XMLSlideShow ppt1 = new XMLSlideShow();

		ppt1.createSlide().importContent(ppt.getSlides().get(0));
		ppt1.createSlide().importContent(ppt.getSlides().get(1));
		ppt1.createSlide().importContent(ppt.getSlides().get(3));
		ppt1.createSlide().importContent(ppt.getSlides().get(4));

		// --------------------- Apply Donut Chart ----------------------------

		String chartTitle = "10 languages with most speakers as first language"; // first line is chart title
		String seriesText = "countries,speakers,language";
		String[] series = seriesText == null ? new String[0] : seriesText.split(",");

		// Category Axis Data
		List<String> listLanguages = new ArrayList<>(10);

		// Values
		List<Double> listCountries = new ArrayList<>(10);
		List<Double> listSpeakers = new ArrayList<>(10);

		// set model
//        String ln;
//        while((ln = modelReader.readLine()) != null) {
//            String[] vals = ln.split(",");
//            listCountries.add(Double.valueOf(vals[0]));
//            listSpeakers.add(Double.valueOf(vals[1]));
//            listLanguages.add(vals[2]);
//        }

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

		String[] categories = listLanguages.toArray(new String[0]);
		Double[] values1 = listCountries.toArray(new Double[0]);
		Double[] values2 = listSpeakers.toArray(new Double[0]);

		try {
			createSlideWithChart(ppt1, chartTitle, series, categories, values1, COLUMN_COUNTRIES);
			//createSlideWithChart(ppt1, chartTitle, series, categories, values2, COLUMN_SPEAKERS);
			// save the result
//			try (OutputStream out = new FileOutputStream("doughnut-chart-from-scratch.pptx")) {
//				ppt1.write(out);
//			}
		} catch (Exception e) {
			System.out.println(e);
		}

		// ---------------------------------------------------------------------

		

		FileOutputStream out = new FileOutputStream("merged.pptx");
		ppt1.write(out);
		out.close();

	}

	private static void createSlideWithChart(XMLSlideShow ppt, String chartTitle, String[] series, String[] categories,
			Double[] values, int valuesColumn) {
		XSLFSlide slide = ppt.createSlide();
		XSLFChart chart = ppt.createChart();
		Rectangle2D rect2D = new java.awt.Rectangle(fromCM(1.5), fromCM(4), fromCM(22), fromCM(14));
		slide.addChart(chart, rect2D);
		
		try {
			setDoughnutData(chart, chartTitle, series, categories, values, valuesColumn);
		}catch(Exception e ) {
			System.out.println(e);
		}
		
		
		//---------------------------------Create Table ----------------------------------------
		
		XSLFTable tbl = slide.createTable();
        tbl.setAnchor(new Rectangle(50, 50, 450, 300));

        int numColumns = 3;
        int numRows = 5;
        XSLFTableRow headerRow = tbl.addRow();
        headerRow.setHeight(50);
        // header
        for (int i = 0; i < numColumns; i++) {
            XSLFTableCell th = headerRow.addCell();
            XSLFTextParagraph p = th.addNewTextParagraph();
            p.setTextAlign(TextAlign.CENTER);
            XSLFTextRun r = p.addNewTextRun();
            r.setText("Header " + (i + 1));
            r.setBold(true);
            r.setFontColor(Color.white);
            th.setFillColor(new Color(79, 129, 189));
            th.setBorderWidth(BorderEdge.bottom, 2.0);
            th.setBorderColor(BorderEdge.bottom, Color.white);

            tbl.setColumnWidth(i, 150);  // all columns are equally sized
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
	}

	private static int fromCM(double cm) {
		return (int) (Math.rint(cm * Units.EMU_PER_CENTIMETER));
	}

	private static void setDoughnutData(XSLFChart chart, String chartTitle, String[] series, String[] categories,
			Double[] values, int valuesColumn) {
		final int numOfPoints = categories.length;
		final String categoryDataRange = chart
				.formatRange(new CellRangeAddress(1, numOfPoints, COLUMN_LANGUAGES, COLUMN_LANGUAGES));
		final String valuesDataRange = chart
				.formatRange(new CellRangeAddress(1, numOfPoints, valuesColumn, valuesColumn));
		final XDDFDataSource<?> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange,
				COLUMN_LANGUAGES);
		final XDDFNumericalDataSource<? extends Number> valuesData = XDDFDataSourcesFactory.fromArray(values,
				valuesDataRange, valuesColumn);
		valuesData.setFormatCode("General");

		XDDFDoughnutChartData data = (XDDFDoughnutChartData) chart.createData(ChartTypes.DOUGHNUT, null, null);
		XDDFDoughnutChartData.Series series1 = (XDDFDoughnutChartData.Series) data.addSeries(categoriesData,
				valuesData);
		series1.setTitle(series[0], chart.setSheetTitle(series[valuesColumn - 1], valuesColumn));

		data.setVaryColors(true);
		// data.setHoleSize(42);
		// data.setFirstSliceAngle(90);
		chart.plot(data);

		XDDFChartLegend legend = chart.getOrAddLegend();
		legend.setPosition(LegendPosition.LEFT);
		legend.setOverlay(false);

		chart.setTitleText(chartTitle);
		chart.setTitleOverlay(false);
		chart.setAutoTitleDeleted(false);
	}

}

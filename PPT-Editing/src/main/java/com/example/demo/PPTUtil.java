package com.example.demo;

import java.awt.Color;
import java.awt.Rectangle;
import java.io.File;
import java.io.IOException;

import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

public class PPTUtil {
	
	public static void topHeading(XSLFTextBox shape1,Rectangle position,Color color,String text) {
		XSLFTextParagraph p1 = shape1.addNewTextParagraph();
		XSLFTextRun r1 = p1.addNewTextRun();
		r1.setText(text);
		
		shape1.setAnchor(position);

		r1.setFontSize(35d);
		r1.setBold(true);
		r1.setFontColor(color);
	}

	public static void createParagraph(XSLFSlide slide,Rectangle position,String text,Double fontSize,Double spaceBefore,Double spaceAfter) {
		XSLFTextBox paraShape = slide.createTextBox();
		// initial height of the text box is 100 pt but
		
		XSLFTextParagraph paragraph = paraShape.addNewTextParagraph();
		// If spaceBefore >= 0, then space is a percentage of normal line height.
		// If spaceBefore < 0, the absolute value of linespacing is the spacing in points
		paragraph.setSpaceBefore(spaceBefore); // 20 pt from the previous paragraph
		paragraph.setSpaceAfter(spaceAfter); // 3 lines after the paragraph
		XSLFTextRun paraText = paragraph.addNewTextRun();
		paraText.setFontSize(fontSize);
		paraText.setText(text);
		
		paraShape.setAnchor(position);
	}
	
	public static void addBulletPoints(XSLFSlide slide,String bulletPoints, Rectangle position) {
		XSLFTextBox paraShape = slide.createTextBox();
		// initial height of the text box is 100 pt but
		
		XSLFTextParagraph paragraph = paraShape.addNewTextParagraph();
		paragraph.setIndentLevel(0);
		paragraph.setBullet(true);
		// If spaceBefore >= 0, then space is a percentage of normal line height.
		// If spaceBefore < 0, the absolute value of linespacing is the spacing in points
		XSLFTextRun paraText = paragraph.addNewTextRun();
		paraText.setText(bulletPoints);
		
		paraShape.setAnchor(position);
	}
	
	public static void footerImage(XMLSlideShow ppt, XSLFSlide slide1) throws IOException {
		File footerImg = new File("/Users/balvinder/office/eclipse-workspace/tooling/PPT-Editing/src/main/resources/static/img/hcl.png");
		XSLFPictureData footerPictureData = ppt.addPicture(footerImg, PictureType.PNG);

		/*XSLFPictureShape shape =*/
		XSLFPictureShape footerPictureShape = slide1.createPicture(footerPictureData);
		Rectangle footerAnchorPicture = new Rectangle(1600, 980, 180, 100);
		footerPictureShape.setAnchor(footerAnchorPicture);
	}
	
	public static void topHeaderImg(XMLSlideShow ppt, XSLFSlide slide1,Rectangle position) throws IOException {
		File img = new File("/Users/balvinder/office/eclipse-workspace/tooling/PPT-Editing/src/main/resources/static/img/header.png");
		XSLFPictureData pictureData = ppt.addPicture(img, PictureType.PNG);

		/*XSLFPictureShape shape =*/
		XSLFPictureShape pictureShape = slide1.createPicture(pictureData);
		Rectangle anchorPicture = position;
		pictureShape.setAnchor(anchorPicture);
	}
	
}

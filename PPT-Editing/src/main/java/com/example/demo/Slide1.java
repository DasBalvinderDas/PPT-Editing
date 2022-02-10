/*
 *  ====================================================================
 *    Licensed to the Apache Software Foundation (ASF) under one or more
 *    contributor license agreements.  See the NOTICE file distributed with
 *    this work for additional information regarding copyright ownership.
 *    The ASF licenses this file to You under the Apache License, Version 2.0
 *    (the "License"); you may not use this file except in compliance with
 *    the License.  You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 * ====================================================================
 */

package com.example.demo;

import java.awt.Color;
import java.awt.Rectangle;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.record.FooterRecord;
import org.apache.poi.sl.usermodel.PictureData.PictureType;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

/**
 * Basic paragraph and text formatting
 */
public final class Slide1 {
    private Slide1() {}

    public static void main(String[] args) throws IOException{
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide1 = ppt.createSlide();
            
            ppt.setPageSize(new java.awt.Dimension(1920, 1080));
            
            XSLFTextBox shape1 = slide1.createTextBox();
            // initial height of the text box is 100 pt but
            Rectangle anchor = new Rectangle(10, 100, 550, 100);
            shape1.setAnchor(anchor);

            XSLFTextParagraph p1 = shape1.addNewTextParagraph();
            XSLFTextRun r1 = p1.addNewTextRun();
            r1.setText("Our assessment and planning approach");
            r1.setFontSize(24d);
            r1.setFontColor(new Color(51,51,51));
            
            File img = new File("/Users/balvinder/office/eclipse-workspace/tooling/PPT-Editing/src/main/resources/static/img/header.png");
            XSLFPictureData pictureData = ppt.addPicture(img, PictureType.PNG);

            /*XSLFPictureShape shape =*/
            XSLFPictureShape pictureShape = slide1.createPicture(pictureData);
            Rectangle anchorPicture = new Rectangle(10, 200, 1270, 100);
            pictureShape.setAnchor(anchorPicture);


            XSLFTextBox shape2 = slide1.createTextBox();
            // initial height of the text box is 100 pt but
            Rectangle anchorShape2 = new Rectangle(10, 300, 550, 100);
            shape2.setAnchor(anchorShape2);
            
            XSLFTextParagraph p2 = shape2.addNewTextParagraph();
            // If spaceBefore >= 0, then space is a percentage of normal line height.
            // If spaceBefore < 0, the absolute value of linespacing is the spacing in points
            p2.setSpaceBefore(-20d); // 20 pt from the previous paragraph
            p2.setSpaceAfter(100d); // 3 lines after the paragraph
            XSLFTextRun r2 = p2.addNewTextRun();
            
            
            r2.setText("Our proven assessment and planning approach shown below was employed for your engagement and integrated the data collection, analytics, "
            		+ "and brokerage capabilities of the StratoZone® platform.\n"
            		+ "A critical component of any cloud assessment or planning service is to collect and rely on accurate and timely data. "
            		+ "As the first step of your assessment, the StratoProbe® data collection application is deployed to retrieve and analyze data from your environments.\n"
            		+ "");
            
            r2.setFontSize(20d);
            
            XSLFTextBox shape3 = slide1.createTextBox();
            // initial height of the text box is 100 pt but
            Rectangle anchor3 = new Rectangle(650, 300, 550, 200);
            shape3.setAnchor(anchor3);

            
            XSLFTextParagraph p3 = shape3.addNewTextParagraph();
            // If spaceBefore >= 0, then space is a percentage of normal line height.
            // If spaceBefore < 0, the absolute value of linespacing is the spacing in points
            p3.setSpaceBefore(-20d); // 20 pt from the previous paragraph
            p3.setSpaceAfter(100d); // 3 lines after the paragraph
            XSLFTextRun r3 = p3.addNewTextRun();
            
            r3.setText("Once discovery is completed, the StratoZone® platform is used to analyze data relating to your IT assets that are deemed to be in scope. "
            		+ "This analysis is a combination of proprietary technical methodologies combined with robust industry benchmark data, and real-time pricing "
            		+ "from the cloud-providers you selected for your assessment.\n"
            		+ "The findings, analysis, and recommendations are made available to you within your StratoZone® portal and summarized in this report.\n"
            		+ "");
            
            r3.setFontSize(20d);
            
            
            File footerImg = new File("/Users/balvinder/office/eclipse-workspace/tooling/PPT-Editing/src/main/resources/static/img/hcl.png");
            XSLFPictureData footerPictureData = ppt.addPicture(footerImg, PictureType.PNG);

            /*XSLFPictureShape shape =*/
            XSLFPictureShape footerPictureShape = slide1.createPicture(footerPictureData);
            Rectangle footerAnchorPicture = new Rectangle(1050, 700, 200, 100);
            footerPictureShape.setAnchor(footerAnchorPicture);
            
            
            XSLFTextBox footer = slide1.createTextBox();
            // initial height of the text box is 100 pt but
            Rectangle anchorFooter = new Rectangle(1000, 730, 400, 100);
            footer.setAnchor(anchorFooter);
            
            XSLFTextParagraph footerPara = footer.addNewTextParagraph();
            // If spaceBefore >= 0, then space is a percentage of normal line height.
            // If spaceBefore < 0, the absolute value of linespacing is the spacing in points
            footerPara.setSpaceBefore(-20d); // 20 pt from the previous paragraph
            footerPara.setSpaceAfter(100d); // 3 lines after the paragraph
            XSLFTextRun footerText = footerPara.addNewTextRun();
            
            footerText.setText("Copyright © 2022 HCL Technologies Limited | www.hcltech.com");
            
            footerText.setFontSize(10d);
            
            try (FileOutputStream out = new FileOutputStream("slide1.pptx")) {
                ppt.write(out);
            }
        }
    }
}
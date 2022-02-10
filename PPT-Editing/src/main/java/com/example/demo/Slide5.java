package com.example.demo;

import java.awt.Color;
import java.awt.Rectangle;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;

public final class Slide5 {
    private Slide5() {}

    public static void main(String[] args) throws IOException{
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide1 = ppt.createSlide();
            
            ppt.setPageSize(new java.awt.Dimension(1920, 1080));
            
            XSLFTextBox shape1 = slide1.createTextBox();
            // initial height of the text box is 100 pt but
            
            //Top Heading
            String headindText = "Next Steps";
            PPTUtil.topHeading(shape1,new Rectangle(10, 50, 550, 100),new Color(51,51,51),headindText);
            
            //Top Para
            String introParaText = "To continue the journey towards making public cloud an integral part of your IT infrastructure, we recommend the following next steps.";
            PPTUtil.createParagraph(slide1,new Rectangle(10, 100, 1900, 100),introParaText,30d,-20d,100d);

            //Top Header Image
            PPTUtil.topHeaderImg(ppt, slide1,new Rectangle(10, 200, 1900, 400));

            //First Para
            
            String firstParaText = "Architecture \nEstablishing Your Network, Security, and Identity ManagementArchitects your network, security, and identity management frameworks that will allow you to most efficiently consume public cloud.";
            PPTUtil.createParagraph(slide1,new Rectangle(10, 200, 400, 100),firstParaText,25d,-20d,100d);
            
            //Second Para
            
            String secondParaText = "Migration Planning \nEstablishing and Validating Your Migration PlanEstablishes and tests your migration plan to ensure that your assets can be migrated to your target cloud environment with minimal disruption to your business.";
            PPTUtil.createParagraph(slide1,new Rectangle(510, 200, 400, 100),secondParaText,25d,-20d,100d);

            //Third Para
            
            String thirdParaText = "Migration \nMigrate Your Assets to the CloudWe compare your assets to the best-match benchmark t-shirt sizes and selected the closest match from your cloud providers for each as-is asset.";
            PPTUtil.createParagraph(slide1,new Rectangle(1000, 200, 400, 100),thirdParaText,25d,-20d,100d);

            //Fourth Para
            
            String fourthParaText = "Application Services \nApplication ModernizationAnalyzes key applications for possible cloud-native optimization and modernization.";
            PPTUtil.createParagraph(slide1,new Rectangle(1500, 200, 400, 100),fourthParaText,25d,-20d,100d);

            //Footer
            PPTUtil.footerImage(ppt, slide1);
            
            String footerTextNote = "Copyright © 2022 HCL Technologies Limited | www.hcltech.com";
            PPTUtil.createParagraph(slide1,new Rectangle(1300, 1000, 700, 100),footerTextNote,20d,-20d,100d);
            
            try (FileOutputStream out = new FileOutputStream("slide5.pptx")) {
                ppt.write(out);
            }
        }
    }

}
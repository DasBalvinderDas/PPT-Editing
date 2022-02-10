package com.example.demo;

import java.awt.Color;
import java.awt.Rectangle;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;

public final class Slide2 {
    private Slide2() {}

    public static void main(String[] args) throws IOException{
        try (XMLSlideShow ppt = new XMLSlideShow()) {
            XSLFSlide slide1 = ppt.createSlide();
            
            ppt.setPageSize(new java.awt.Dimension(1920, 1080));
            
            XSLFTextBox shape1 = slide1.createTextBox();
            // initial height of the text box is 100 pt but
            
            //Top Heading
            String headindText = "What we did for your assessment";
            PPTUtil.topHeading(shape1,new Rectangle(10, 50, 550, 100),new Color(51,51,51),headindText);
            
            //Top Header Image
            PPTUtil.topHeaderImg(ppt, slide1,new Rectangle(130, 110, 1800, 100));

            //Left Para
            
            String leftParaText = "Your engagement included a basic assessment. "
			+ "The basic assessment is a starting point for any customer embarking on the journey to public or hybrid cloud."
			+ "One or more data collectors were deployed in your environment and your data was automatically "
			+ "aggregated, analyzed, and staged for additional planning functions. "
			+ "The phases and tasks of the basic assessment are shown below:";
            
            PPTUtil.createParagraph(slide1,new Rectangle(10, 200, 600, 100),leftParaText,30d,-20d,100d);
            
            //Middle Para
            
            String middleParaText = "Phase1 \nDiscovery, Inventory Analysis, and Cloud Readiness "
            		+ "The objective of this phase is to collect data from the target workloads and complete inventory analysis including basic cloud readiness."
            		+ "  The StratoProbe® discovery engine gathered workload, application and network information and processed the following analytics:";
            
            PPTUtil.createParagraph(slide1,new Rectangle(640, 250, 600, 100),middleParaText,25d,-20d,100d);
            
            String leftBullets = " Inventory Analysis \n Asset performance analysis \n Network dependency mapping \n Cloud-readiness scoring \n Application inventory analysis ";
            PPTUtil.addBulletPoints(slide1,leftBullets,new Rectangle(640, 640, 350, 100));
            
            //Right Para
            String rightParaText = "Phase2 \nBasic Cloud Fit and Financial Analysis "
            		+ "The objective of this phase is to further analyze your data to provide insights into cloud readiness, potential savings from cloud, "
            		+ "consumption strategies including IaaS and PaaS alternatives, and to review your projected spend in the selected cloud providers. "
            		+ "The expected output from this phase includes:";
            
            PPTUtil.createParagraph(slide1,new Rectangle(1300, 250, 600, 100),rightParaText,25d,-20d,100d);
            
            String rightBullets = " Best-fit vendor catalog-product match (IaaS) \n PaaS-fit analysis \n IaaS-fit analysis \n Cloud-spend estimates (by vendor catalog) \n TCO and ROI against benchmark baselines ";
            PPTUtil.addBulletPoints(slide1,rightBullets,new Rectangle(1300, 640, 350, 100));

            
            //Footer
            PPTUtil.footerImage(ppt, slide1);
            
            String footerTextNote = "Copyright © 2022 HCL Technologies Limited | www.hcltech.com";
            PPTUtil.createParagraph(slide1,new Rectangle(1300, 1000, 700, 100),footerTextNote,20d,-20d,100d);
            
            try (FileOutputStream out = new FileOutputStream("slide2.pptx")) {
                ppt.write(out);
            }
        }
    }

}
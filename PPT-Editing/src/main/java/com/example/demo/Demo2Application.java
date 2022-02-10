package com.example.demo;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class Demo2Application {

	public static void main(String[] args) {
		//SpringApplication.run(Demo2Application.class, args);
		
		//create a new empty slide show
		XMLSlideShow ppt = new XMLSlideShow();
		//add first slide
		XSLFSlide blankSlide = ppt.createSlide();
		
	}

}

package org.example;

import org.apache.poi.hslf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class change_Back_Ground_and_Text_Color {

    public static void main(String[] args) {
        String folderPath = "C:\\Users\\User\\OneDrive\\Máy tính\\Chầu";

        File folder = new File(folderPath);
        File[] files = folder.listFiles();

        if (files != null) {
            for (File file : files) {
                if (file.isFile() && (file.getName().endsWith(".ppt") || file.getName().endsWith(".pptx"))) {
                    changeBackgroundAndFontColor(file.getAbsolutePath());
                }
            }
        }
    }

    private static void changeBackgroundAndFontColor(String filePath) {
        try {
            FileInputStream fis = new FileInputStream(filePath);

            if (isOOXML(filePath)) {
                XMLSlideShow pptx = new XMLSlideShow(fis);

                for (XSLFSlide slide : pptx.getSlides()) {
                    // Thay đổi màu nền ở đây (mã RGB)
                    XSLFBackground fill = slide.getBackground();
                    //fill.setFillColor(new Color(0,0,255));
                    fill.setFillColor(new Color(0, 0, 102, 255));
                    for (XSLFShape shape : slide.getShapes()) {
                        if (shape instanceof XSLFTextShape) {
                            XSLFTextShape textShape = (XSLFTextShape) shape;

                            for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
                                for (XSLFTextRun textRun : paragraph.getTextRuns()) {
                                    textRun.setFontColor(new Color(255, 255, 255, 255));
                                    //textRun.setFontColor(new Color(247, 253, 0)); // Thay đổi màu chữ ở đây (mã RGB)
                                }
                            }
                        }
                    }
                }

                FileOutputStream out = new FileOutputStream(filePath);
                pptx.write(out);
                out.close();
            } else {
                HSLFSlideShow ppt = new HSLFSlideShow(fis);

                for (HSLFSlide slide : ppt.getSlides()) {
                    HSLFBackground background = slide.getBackground();
                    HSLFFill fill = background.getFill();
                    if(fill.getPictureData()!= null){
                        continue;
                    }
                    //fill.setForegroundColor(new Color(0,0,255));// Thay đổi màu nền ở đây (mã RGB)
                    fill.setForegroundColor(new Color(0, 0, 102, 255));
                    for (HSLFShape shape : slide.getShapes()) {
                        if (shape instanceof HSLFTextShape) {
                            HSLFTextShape textShape = (HSLFTextShape) shape;

                            for (HSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
                                for (HSLFTextRun textRun : paragraph.getTextRuns()) {
                                    textRun.setFontColor(new Color(255, 255, 255, 255));
                                    //textRun.setFontColor(new Color(247, 253, 0)); // Thay đổi màu chữ ở đây (mã RGB)
                                }
                            }
                        }
                    }
                }

                FileOutputStream out = new FileOutputStream(filePath);
                ppt.write(out);
                out.close();
            }

            System.out.println("Đã thay đổi màu nền và màu chữ của tệp tin: \n" + filePath);
            System.out.println();
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    private static boolean isOOXML(String filePath) throws IOException, InvalidFormatException {
        FileInputStream fis = new FileInputStream(filePath);
        try {
            XMLSlideShow pptx = new XMLSlideShow(fis);
            return true;
        } catch (Exception e) {
            return false;
        } finally {
            fis.close();
        }
    }
}
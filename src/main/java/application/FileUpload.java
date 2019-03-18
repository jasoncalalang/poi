package application;

import java.io.File;
import java.io.IOException;
import javax.servlet.ServletException;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.annotation.WebServlet;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.Part;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import java.util.Enumeration;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.URL;

import org.apache.poi.xslf.usermodel.*;
import java.awt.Color;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import java.awt.Rectangle;
import java.awt.geom.Rectangle2D;

import org.apache.poi.xslf.usermodel.XSLFTextShape;

@WebServlet("/UploadServlet")
@MultipartConfig(fileSizeThreshold = 1024 * 1024 * 2, // 2MB
        maxFileSize = 1024 * 1024 * 10, // 10MB
        maxRequestSize = 1024 * 1024 * 50) // 50MB
public class FileUpload extends HttpServlet {
    /**
     * handles file upload
     */
	
	HelloWorld hw = new HelloWorld();
	
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
    	
    	URL fileUrl =  Thread.currentThread().getContextClassLoader().getResource("/MOR-Template_all_copy.pptx");
    	System.out.println(fileUrl);

    	hw.setSelectedMonth(request.getParameter("month"));
    	hw.setPptxTemplate(fileUrl.toString());
    	XMLSlideShow powerpoint = processFile(request.getPart("filename").getInputStream());
        response.setContentType("application/vnd.openxmlformats-officedocument.presentationml.presentation");
        response.setHeader("Content-Disposition", "filename=\"presentation.pptx\"");
        powerpoint.write(response.getOutputStream());
        powerpoint.close();
        response.getOutputStream().close();

    }

    /**
     * Extracts file name from HTTP header content-disposition
     */
    private String extractFileName(Part part) {
        String contentDisp = part.getHeader("content-disposition");
        String[] items = contentDisp.split(";");
        for (String s : items) {
            if (s.trim().startsWith("filename")) {
                return s.substring(s.indexOf("=") + 2, s.length() - 1);
            }
        }
        return "";
    }

    private XMLSlideShow processFile(InputStream io) {

        XMLSlideShow powerpoint = null;
        try {
            Workbook workbook = WorkbookFactory.create(io);
            powerpoint = hw.process(workbook);
            workbook.close();

        } catch (Exception e) {
            e.printStackTrace();
        }

        return powerpoint;
    }
}
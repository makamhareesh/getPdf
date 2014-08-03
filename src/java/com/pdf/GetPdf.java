/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package com.pdf;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.Image;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.pdf.RandomAccessFileOrArray;
import com.itextpdf.text.pdf.codec.GifImage;
import com.itextpdf.text.pdf.codec.TiffImage;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;  
import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;  
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

/**
 *
 * @author hmakam
 */
public class GetPdf extends HttpServlet {

    /**
     * Processes requests for both HTTP <code>GET</code> and <code>POST</code>
     * methods.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        
        String url = request.getParameter("text");
        
        response.setContentType("application/pdf");
        ServletOutputStream out = response.getOutputStream();
        convertToPdf(url,out);
        out.flush();
        out.close();
    }
    
    private void convertToPdf(String url,ServletOutputStream out){
        
     try {
        String contentType = new URL(url).openConnection().getContentType();
        
      Document document = new Document();
      PdfWriter.getInstance(document,out);
          document.open();
        if(contentType.equalsIgnoreCase("application/vnd.ms-excel")){
            addXls(document, url, "xls");
            
        }else if(contentType.equalsIgnoreCase("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")){
            addXls(document, url, "xlsx");
        }else if (contentType.equalsIgnoreCase("application/msword")){
            docConvert(document, url, "doc");
        }else if(contentType.equalsIgnoreCase("application/vnd.openxmlformats-officedocument.wordprocessingml.document")){
            docConvert(document, url, "docx");
        }else if(contentType.equalsIgnoreCase("image/tiff")){
            addTif(document, url);
        }else if (contentType.equalsIgnoreCase("image/gif")){
            addGif(document,url);
        }else{
            Image image = Image.getInstance(url);
            image.scaleToFit(550,800);
            document.add(image);
         
        }
           document.close();
        }
      
    catch (MalformedURLException ex) {
            Logger.getLogger(GetPdf.class.getName()).log(Level.SEVERE, null, ex);
    }   catch (IOException ex) {
            Logger.getLogger(GetPdf.class.getName()).log(Level.SEVERE, null, ex);
        }catch (DocumentException ex) {
            Logger.getLogger(GetPdf.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        
    }
    
 public static void addGif(Document document, String path) throws IOException, DocumentException {
        GifImage img = new GifImage(path);
        int n = img.getFrameCount();
        for (int i = 1; i <= n; i++) {
            document.add(img.getImage(i));
        }
    }

public static void addTif(Document document, String path) throws DocumentException, IOException {
        RandomAccessFileOrArray ra = new RandomAccessFileOrArray(new URL(path));
        int n = TiffImage.getNumberOfPages(ra);
        Image img;
        for (int i = 1; i <= n; i++) {
            img = TiffImage.getTiffImage(ra, i);
            img.scaleToFit(550,800);
            document.add(img);
        }
    }   

public static void addXls(Document document,String url,String type) throws IOException, DocumentException{
    Iterator<Row> rowIterator;
    int colNo;
    if(type.equals("xls")){
    HSSFWorkbook excelWorkbook = new HSSFWorkbook(new URL(url).openStream());
    HSSFSheet my_worksheet = excelWorkbook.getSheetAt(0);
    rowIterator = my_worksheet.iterator();
    colNo = my_worksheet.getRow(0).getLastCellNum();
    }
    else{
    XSSFWorkbook excelWorkbook1 = new XSSFWorkbook(new URL(url).openStream());
     XSSFSheet my_worksheet = excelWorkbook1.getSheetAt(0); 
     rowIterator = my_worksheet.iterator();
     colNo = my_worksheet.getRow(0).getLastCellNum();
    }
    PdfPTable my_table = new PdfPTable(colNo);
    PdfPCell table_cell = null;
     while(rowIterator.hasNext()) {
                        Row row = rowIterator.next(); //Read Rows from Excel document       
                        Iterator<Cell> cellIterator = row.cellIterator();//Read every column for every row that is READ
                                while(cellIterator.hasNext()) {
                                        Cell cell = cellIterator.next(); //Fetch CELL
                                      if(cell.getCellType() == (Cell.CELL_TYPE_NUMERIC)){
                                          table_cell=new PdfPCell(new Phrase(new Double(cell.getNumericCellValue()).toString()));
                                          System.out.println(cell.getNumericCellValue());
                                          my_table.addCell(table_cell);
                                      }else if(cell.getCellType() == (Cell.CELL_TYPE_STRING)){
                                          table_cell=new PdfPCell(new Phrase(cell.getStringCellValue()));
                                          System.out.println(cell.getStringCellValue());
                                          my_table.addCell(table_cell);
                                      }else if(cell.getCellType() == (Cell.CELL_TYPE_FORMULA)){
                                          table_cell=new PdfPCell(new Phrase(cell.getCellFormula()));
                                          my_table.addCell(table_cell);
                                      }else if(cell.getCellType() == (Cell.CELL_TYPE_BLANK)){
                                          table_cell=new PdfPCell(new Phrase(""));
                                          my_table.addCell(table_cell);
                                      }else{
                                          table_cell=new PdfPCell(new Phrase(""));
                                          my_table.addCell(table_cell);
                                      }
                                }
                }
    document.add(my_table);
}

public static void docConvert(Document document,String url,String type) throws IOException, DocumentException{
      WordExtractor we;
     
      
    if(type.equals("doc")){
    HWPFDocument  wordDoc = new HWPFDocument(new URL(url).openStream());
     we = new WordExtractor(wordDoc);
      String[] paragraphs = we.getParagraphText();
      for (int i = 0; i < paragraphs.length; i++) {  
        paragraphs[i] = paragraphs[i].replaceAll("\\cM?\r?\n", "");  
      document.add(new Paragraph(paragraphs[i]));
      }
    }
    else{
    XWPFDocument  wordDoc = new XWPFDocument(new URL(url).openStream());
     List<IBodyElement> contents = wordDoc.getBodyElements();
     
     for(IBodyElement content:contents){
          if(content.getElementType()== BodyElementType.PARAGRAPH){
           List<XWPFParagraph> paras =   content.getBody().getParagraphs();
           for(XWPFParagraph para:paras){
               document.add(new Paragraph(para.getParagraphText()));
           }
              
          }else if(content.getElementType() == BodyElementType.TABLE){
               List<XWPFTable> tables = content.getBody().getTables();
             for(XWPFTable table:tables){
                 List<XWPFTableRow> rows = table.getRows();
                 for(XWPFTableRow row:rows){
                   List<XWPFTableCell> tablecells =  row.getTableCells();
                 }
           }   
          }
         
     }
    }
    
}

    // <editor-fold defaultstate="collapsed" desc="HttpServlet methods. Click on the + sign on the left to edit the code.">
    /**
     * Handles the HTTP <code>GET</code> method.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        processRequest(request, response);
    }

    /**
     * Handles the HTTP <code>POST</code> method.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
        processRequest(request, response);
    }

    /**
     * Returns a short description of the servlet.
     *
     * @return a String containing servlet description
     */
    @Override
    public String getServletInfo() {
        return "Short description";
    }// </editor-fold>



}
 
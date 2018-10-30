/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.rtassignmnet2;


import com.itextpdf.text.Document;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 *
 * @author lenovo
 */
public class Try {
    public static void main(String args[]) throws IOException, BiffException{
        ExcelToPdf();
    }
    public static void ExcelToPdf(){
        try{
            InputStream file1;
            FileOutputStream file2;
            file1 = new FileInputStream("C:\\Users\\lenovo\\Desktop\\chessResultsList.xls");
            file2 = new FileOutputStream("C:\\Users\\lenovo\\Desktop\\table.pdf");
            Workbook wb=Workbook.getWorkbook(file1);
            Sheet sheet=wb.getSheet(0);
            Document doc = new Document();
            doc.setPageSize(PageSize.A4);
            PdfWriter.getInstance(doc,file2);
            doc.open();
            int rowNum=sheet.getRows();
            int colNum=sheet.getColumns();
            String[][] table=new String[rowNum][colNum];
            PdfPTable pdftable = new PdfPTable(colNum);
            pdftable.setWidthPercentage(100);
            pdftable.setHorizontalAlignment(PdfPTable.ALIGN_LEFT);
            for(int j=0;j<rowNum;j++){
               for(int i=0;i<colNum;i++){
                   Cell cell=sheet.getCell(i,j);
                   table[j][i] =cell.getContents();
                   System.out.printf("%40s",table[j][i]);
                   pdftable.addCell(table[j][i]);
               }
               System.out.println("");
            }
            doc.add(pdftable);
            doc.close();
        }catch(Exception e){
            System.out.println(e);
        }
    }
}

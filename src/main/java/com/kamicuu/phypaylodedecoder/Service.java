/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.kamicuu.phypaylodedecoder;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.bind.DatatypeConverter;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Kamil
 */
public class Service {
    
     public void loadFromExcel(String fileLocation) {
         
         XSSFSheet sheet = null;
         XSSFWorkbook workbook = null;
         OutputStream fileOut = null;

         try {
             workbook = new XSSFWorkbook(new FileInputStream(new File(fileLocation))); 
             sheet = workbook.getSheetAt(0);
             System.out.println("Otwarto plik");
         } catch (FileNotFoundException ex) {
             Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
             System.out.println("Nie znaleziono pliku");
         } catch (IOException ex) {
             System.out.println("Otawrcie pliku nie powiodło sie (IOException)");
             Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
         }
    
         Iterator<Row> rowIterator = sheet.iterator();
         
          while (rowIterator.hasNext()){
             Row row = rowIterator.next();
            
             var cell = row.getCell(6);
             var lastCell = row.getLastCellNum();
             
             if(cell.getStringCellValue().equals("phyPayload")){
                 row.createCell(lastCell).setCellValue("devEui");
                 row.createCell(lastCell+1).setCellValue("joinEui");
                 System.out.println("Rozpoczynam przetwarzanie phyPayload");
             }else{
                List<String> data = getDevEuiFromPhyPalode(cell.getStringCellValue());
                
                row.createCell(lastCell).setCellValue(data.get(0));
                row.createCell(lastCell+1).setCellValue(data.get(1));
             
             }
             
          }
         try {
             fileOut = new FileOutputStream(fileLocation);
             workbook.write(fileOut);
             System.out.println("Przetwarzanie zakonczone powodzeniem");
         } catch (FileNotFoundException ex) {
             System.out.println("Problem ze znalezieniem pliku");
             Logger.getLogger(Service.class.getName()).log(Level.SEVERE, null, ex);
         } catch (IOException ex) {
             System.out.println("Zapis pliku nie powiodło sie (IOException)");
             Logger.getLogger(Service.class.getName()).log(Level.SEVERE, null, ex);
         }
         
    }
     
     public List<String> getDevEuiFromPhyPalode(String phyPaylode){
         
        List<String> output =  new ArrayList<>();
         
        byte[] devEui = new byte[8];
        byte[] joinEui = new byte[8];
        
        byte[] decodedBytes = Base64.getDecoder().decode(phyPaylode);
        ArrayUtils.reverse(decodedBytes);
        
        
        if(decodedBytes.length==23){
            
            for(int i=0; i<8; i++){
                devEui[i] = decodedBytes[6+i];
                joinEui[i] = decodedBytes[14+i];
            }

            output.add(DatatypeConverter.printHexBinary(devEui));
            output.add(DatatypeConverter.printHexBinary(joinEui));
                
        }else{
            output.add("Niepoprawana liczba bajtow w ramce!");
            output.add("Niepoprawana liczba bajtow w ramce!");
        }

     
        return output;
     }
}


package pl.siłownia;

import com.sun.prism.paint.Color;
import java.awt.Font;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Exel_Actions {
    
   static File plik_Z_Danymi; 
   static String nazwaPlikuExel;
   static FileOutputStream strumienZapisu; 
    FileInputStream strumienOdczytu;
   
   
    static XSSFWorkbook arkusz; 
    static XSSFSheet strona;
    XSSFRow aktualnyWierszStrony;
    XSSFCell aktualnaKomorkaWiersza;
   

    public Exel_Actions() {
                       
    }
   
    
    static void stworzNowyPlikExel(String tytulPliku)
    {
        nazwaPlikuExel=tytulPliku;
        plik_Z_Danymi = new File(tytulPliku);
        arkusz=new XSSFWorkbook();
        strona = arkusz.createSheet("1");
        
       try {
           strumienZapisu = new FileOutputStream(plik_Z_Danymi);
           arkusz.write(strumienZapisu);
           strumienZapisu.close();
           
       } catch (FileNotFoundException wyjątek1) {
         
           JOptionPane.showMessageDialog(null,"Plik o podanym tytule nie istnieje !!!","Błąd",JOptionPane.ERROR_MESSAGE);
           
       } catch (IOException ex) {
           
           Logger.getLogger(Exel_Actions.class.getName()).log(Level.SEVERE, null, ex);
       }
                 
    }
    
   void stworzTabeleDoPomiaruEfektowTreningu(List tablicaTytulowNaglowka)
   {
       
       try {
           strumienOdczytu=new FileInputStream(plik_Z_Danymi);
           arkusz=new XSSFWorkbook(strumienOdczytu);
           strona=arkusz.getSheet("1");
           
           stworzNaglowekTabeli(tablicaTytulowNaglowka);
                   
       } catch (FileNotFoundException ex) {
           Logger.getLogger(Exel_Actions.class.getName()).log(Level.SEVERE, null, ex);
       } catch (IOException ex) {
           Logger.getLogger(Exel_Actions.class.getName()).log(Level.SEVERE, null, ex);
       }
       
      
       
   }
 
    
   
  // void stworzNaglowekTabeli(String[] polaNaglowka)
   void stworzNaglowekTabeli(List tablicaTytulowNaglowka)
   {
       aktualnyWierszStrony=strona.createRow(0);
           
           XSSFCellStyle stylNagłowka = arkusz.createCellStyle();
           XSSFFont czcionkaNaglowka = arkusz.createFont();
           czcionkaNaglowka.setColor(HSSFColor.AQUA.index);
           stylNagłowka.setFont(czcionkaNaglowka);
           
           stylNagłowka.setFillForegroundColor(HSSFColor.GREEN.index);
           stylNagłowka.setFillPattern(FillPatternType.SOLID_FOREGROUND);
          
           //String tytul = (String)tablicaTytulowNaglowka.get(0);
           for(int i=0; i<12;i++)
           {
                aktualnaKomorkaWiersza=aktualnyWierszStrony.createCell(i);
                aktualnaKomorkaWiersza.setCellValue((String)tablicaTytulowNaglowka.get(i));              
                aktualnaKomorkaWiersza.setCellStyle(stylNagłowka);               
           }
                  
       try {
           
           strumienZapisu = new FileOutputStream(plik_Z_Danymi);
           arkusz.write(strumienZapisu);
           strumienZapisu.close();
           
       } catch (FileNotFoundException ex) {
           Logger.getLogger(Exel_Actions.class.getName()).log(Level.SEVERE, null, ex);
       } catch (IOException ex) {
           Logger.getLogger(Exel_Actions.class.getName()).log(Level.SEVERE, null, ex);
       }
                
   }
   
   
   
   
   
    
}

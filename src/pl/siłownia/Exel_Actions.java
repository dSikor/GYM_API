
package pl.siłownia;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Exel_Actions {
    
   static File plik_Z_Danymi; 
   static String nazwaPlikuExel;
   static FileOutputStream strumienZapisu;
   
    XSSFWorkbook arkusz; 
    XSSFSheet strona;
    XSSFRow aktualnyWierszStrony;
    XSSFCell aktualnaKomurkaWiersza;
   

    public Exel_Actions() {
                
        
        
        
        
    }
   
    
    static void stworzNowyPlikExel(String tytulPliku)
    {
        nazwaPlikuExel=tytulPliku;
        plik_Z_Danymi = new File(tytulPliku);
        
       try {
           strumienZapisu = new FileOutputStream(plik_Z_Danymi);
           
       } catch (FileNotFoundException wyjątek1) {
         
           JOptionPane.showMessageDialog(null,"Plik o podanym tytule nie istnieje !!!","Błąd",JOptionPane.ERROR_MESSAGE);
       }
        
        
        
        
        
    }
    

}

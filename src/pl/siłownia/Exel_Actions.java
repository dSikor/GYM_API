
package pl.siłownia;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
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
    
    
    
    
    
    
    
    
    
    
    

}

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pl.siłownia;

import java.awt.Dimension;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import javax.imageio.ImageIO;
import javax.swing.JPanel;

/**
 *
 * @author Damian
 */
public class JPanel_Gafika extends JPanel{
    
    private BufferedImage grafikaSilacz;

    public JPanel_Gafika(String nazwaPlikuJPG) {
      //  public JPanel_Gafika() {
        super();
       
        File imageFile = new File(nazwaPlikuJPG);
        
		try {
			grafikaSilacz = ImageIO.read(imageFile);
		} catch (IOException e) {
			System.err.println("Blad odczytu obrazka");
			e.printStackTrace();
		}
        
                
                Dimension dimension = new Dimension(grafikaSilacz.getWidth(), grafikaSilacz.getHeight());
		setPreferredSize(dimension);
    }
     
    @Override
	public void paintComponent(Graphics g) {
            
		Graphics2D g2d = (Graphics2D) g;
                //g2d.scale(1.0/7,1.0/7);
                g2d.drawImage(grafikaSilacz, 0, 0, 181,250, this);
		//g2d.drawImage(grafikaSilacz, 0, 0, this);
	}      
}

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;

public class PptConverter {
	public static void convertToImages(String inputFile) throws FileNotFoundException, IOException
	{
		//creating an empty presentation
	      File file=new File(inputFile);
	      HSLFSlideShow ppt = new HSLFSlideShow(new FileInputStream(file));
	      
	      //getting the dimensions and size of the slide 
	      Dimension pgsize = ppt.getPageSize();
	      List<HSLFSlide> slide = ppt.getSlides();
	      
	      for (int i = 0; i < slide.size(); i++) {
	         BufferedImage img = new BufferedImage(pgsize.width, pgsize.height,BufferedImage.TYPE_INT_RGB);
	         Graphics2D graphics = img.createGraphics();

	         //clear the drawing area
	         graphics.setPaint(Color.white);
	         graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));

	         //render
	         slide.get(i).draw(graphics);
	         
	         //creating an image file as output
		     FileOutputStream out = new FileOutputStream("ppt_image.png");
		     javax.imageio.ImageIO.write(img, "png", out);
		     ppt.write(out);
		      
		     System.out.println("Image successfully created");
		     out.close();	
		     
	      }
	      ppt.close();
	}
}

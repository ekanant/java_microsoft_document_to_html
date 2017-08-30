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
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

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
	
	public static void convertToOnePhoto(String inputFile) throws FileNotFoundException, IOException
	{
		File file=new File(inputFile);
		HSLFSlideShow ppt = new HSLFSlideShow(new FileInputStream(file));
		
		//getting the dimensions and size of the slide 
		Dimension pgsize = ppt.getPageSize();
		List<HSLFSlide> slide = ppt.getSlides();
		int gapBetweenSlide = 50;
		List<BufferedImage> slideRendered = new ArrayList<BufferedImage>();

		for (int i = 0; i < slide.size(); i++) {
			
			BufferedImage img = new BufferedImage(pgsize.width, 
					(pgsize.height + gapBetweenSlide) * slide.size(),
					BufferedImage.TYPE_INT_RGB);
			Graphics2D graphics = img.createGraphics();
			//clear the drawing area
			graphics.setPaint(Color.BLACK);
			graphics.fill(new Rectangle2D.Float(0, pgsize.height, pgsize.width, pgsize.height));

			//render
			slide.get(i).draw(graphics);
		     
			slideRendered.add(img);
		     
		}
		
		BufferedImage mergeImage = new BufferedImage(pgsize.width, 
				(pgsize.height + gapBetweenSlide) * slide.size(),
				BufferedImage.TYPE_INT_RGB);
		Graphics2D graphics = mergeImage.createGraphics();
		int offsetToDraw = 0;
		for(int i = 0 ; i < slideRendered.size() ; i++)
		{
			offsetToDraw = (int)(i * pgsize.getHeight());
			if(i > 0) offsetToDraw += gapBetweenSlide;
			graphics.drawImage(slideRendered.get(i), null, 0, offsetToDraw);
		}
		
		

		
		//creating an image file as output
		FileOutputStream out = new FileOutputStream("ppt_image.png");
		javax.imageio.ImageIO.write(mergeImage, "png", out);
		ppt.write(out);
		out.close();	
		
		
		ppt.close();
	}
}

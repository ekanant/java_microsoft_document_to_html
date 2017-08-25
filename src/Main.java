import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.io.OutputStream;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocumentCore;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.converter.WordToHtmlUtils;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.w3c.dom.Document;


public class Main {

	public static void main(String[] args) {
		String docPath = "xxx.doc";
		convertDocToHtml(docPath);
		
		
		String docxPath = "xxx.docx";
		convertDocxToHtml(docxPath);
	}
	public static String convertDocToHtml(String inputFile) 
	{
		String html = "";
		try
		{
			// TODO Auto-generated method stub
			HWPFDocumentCore wordDocument = WordToHtmlUtils.loadDoc(new FileInputStream(inputFile));

		    WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
		            DocumentBuilderFactory.newInstance().newDocumentBuilder()
		                    .newDocument());
		    wordToHtmlConverter.processDocument(wordDocument);
		    Document htmlDocument = wordToHtmlConverter.getDocument();
		    ByteArrayOutputStream out = new ByteArrayOutputStream();
		    DOMSource domSource = new DOMSource(htmlDocument);
		    StreamResult streamResult = new StreamResult(out);

		    TransformerFactory tf = TransformerFactory.newInstance();
		    Transformer serializer = tf.newTransformer();
		    serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
		    serializer.setOutputProperty(OutputKeys.INDENT, "yes");
		    serializer.setOutputProperty(OutputKeys.METHOD, "html");
		    serializer.transform(domSource, streamResult);
		    out.close();

		    html = new String(out.toByteArray());
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		return html;
	}
	
	static String convertDocxToHtml(String inputFile)
	{
		String html = "";
		try
		{
		//convert .docx to HTML string
	        InputStream in= new FileInputStream(new File(inputFile));
	        XWPFDocument document = new XWPFDocument(in);

	        XHTMLOptions options = XHTMLOptions.create().URIResolver(new FileURIResolver(new File("word/media")));

	        OutputStream out = new ByteArrayOutputStream();


	        XHTMLConverter.getInstance().convert(document, out, options);
	        html=out.toString();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		return html;
	}
}

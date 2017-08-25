import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;

public class Main {

	public static void main(String[] args) throws IOException, TransformerException, ParserConfigurationException {
		String appPath = new File(".").getCanonicalPath();
		String docPath = appPath + File.separator + "doc.doc";
		String docHtml = WordDocConverter.convertToHtml(docPath);
		PrintWriter writer = new PrintWriter(appPath + File.separator + "doc.html", "UTF-8");
		writer.println(docHtml);
		writer.close();

		String docxPath = new File(".").getCanonicalPath() + File.separator + "docx.docx";
		String docxHtml = WordDocxConverter.convertToHtml(docxPath);
		PrintWriter writer2 = new PrintWriter(appPath + File.separator + "docx.html", "UTF-8");
		writer2.println(docxHtml);
		writer2.close();

	}

}

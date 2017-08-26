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

		String docxPath = appPath + File.separator + "docx.docx";
		String docxHtml = WordDocxConverter.convertToHtml(docxPath);
		PrintWriter writer2 = new PrintWriter(appPath + File.separator + "docx.html", "UTF-8");
		writer2.println(docxHtml);
		writer2.close();
		
		String xlsPath = appPath + File.separator + "xls.xls";
		StringBuilder sb = new StringBuilder();
		ExcelConverter x = ExcelConverter.create(xlsPath, sb);
		x.setCompleteHTML(true);
        x.printPage();
        PrintWriter writer3 = new PrintWriter(appPath + File.separator + "xls.html", "UTF-8");
		writer3.println(sb.toString());
		writer3.close();
		
		String xlsxPath = appPath + File.separator + "xlsx.xlsx";
		StringBuilder sb2 = new StringBuilder();
		ExcelConverter x2 = ExcelConverter.create(xlsxPath, sb2);
		x2.setCompleteHTML(true);
        x2.printPage();
        PrintWriter writer4 = new PrintWriter(appPath + File.separator + "xlsx.html", "UTF-8");
		writer4.println(sb2.toString());
		writer4.close();
	}

}

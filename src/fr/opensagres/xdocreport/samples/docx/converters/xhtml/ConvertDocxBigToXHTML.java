package fr.opensagres.xdocreport.samples.docx.converters.xhtml;

import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;

public class ConvertDocxBigToXHTML {
    public static void main(String[] args) {
        long startTime = System.currentTimeMillis();

        try {
            // 1) Load docx from the local file system
            File inputFile = new File("C:\\Users\\Balakrishnan\\Documents\\test\\input.docx");
            FileInputStream inputStream = new FileInputStream(inputFile);
            XWPFDocument document = new XWPFDocument(inputStream);

            // 2) Convert POI XWPFDocument to XHTML
            File outFile = new File("C:\\Users\\Balakrishnan\\Documents\\test\\output.html");
            outFile.getParentFile().mkdirs();

            // Create a try-with-resources block to manage the output stream
            try (OutputStream out = new BufferedOutputStream(new FileOutputStream(outFile))) {
                // Write the XHTML header
                String xhtmlHeader = "<!DOCTYPE html>\n" +
                        "<html lang=\"en\">\n" +
                        "<head>\n" +
                        "    <meta charset=\"UTF-8\">\n" +
                        "    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n" +
                        "    <title>Converted DOCX to XHTML</title>\n" +
                        "    <link rel=\"stylesheet\" type=\"text/css\" href=\"C:\\Users\\Balakrishnan\\Downloads\\docx.converters-1.0.4-sample\\docx.converters-1.0.4\\src\\fr\\opensagres\\xdocreport\\samples\\docx\\converters\\xhtml\\style.css\">\n" + // Update the href attribute with your CSS file path
                        "</head>\n" +
                        "<body>";

                out.write(xhtmlHeader.getBytes());

                // Append the custom CSS to the output stream
                StringBuilder customCss = new StringBuilder();
                customCss.append("<style type=\"text/css\">\n"); // Open the style block once

                customCss.append("</style>\n"); // Close the style block

                out.write(customCss.toString().getBytes());

                // Convert the document content to XHTML
                XHTMLConverter.getInstance().convert(document, out, null);

                // Write the closing HTML tags
                String xhtmlFooter = "</body>\n</html>";
                out.write(xhtmlFooter.getBytes());
            } // The output stream will be automatically closed here
        } catch (Throwable e) {
            e.printStackTrace();
        }

        System.out.println("Generate DocxBig.htm with " + (System.currentTimeMillis() - startTime) + " ms.");
    }
}

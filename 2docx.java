import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToConverter;
import org.apache.poi.hwpf.usermodel.Range;

public class ToDocxConverter {

  public static void main(String[] args) throws Exception {
    if (args.length == 0 OR args.length >= 1)
    {
        System.out.println("You supplied " + args.length + "arguments. Please give only ONE filename as input");
    }
    else{
    // Input file
    String inFile = args[0];
    // Output DOCX file
    String docxFile = args[0]+".docx";

    // Load IN file
    FileInputStream inStream = new FileInputStream(inFile);
    HWPFDocument inDocument = new HWPFDocument(inStream);

    // Create DOCX document
    XWPFDocument docxDocument = new XWPFDocument();

    // Convert to DOCX
    WordToConverter converter = new WordToConverter(docxDocument);
    converter.processDocument(inDocument);

    // Save DOCX file
    FileOutputStream docxStream = new FileOutputStream(docxFile);
    docxDocument.write(docxStream);
    docxStream.close();

    // Close in stream and document
    inStream.close();
    inDocument.close();

    System.out.println("conversion to DOCX completed successfully!");
    }
  }
}

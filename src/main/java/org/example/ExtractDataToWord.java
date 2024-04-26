package org.example;


import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.regex.Pattern;
import org.apache.pdfbox.io.RandomAccessFile;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class ExtractDataToWord {


  public void execute(String pdfFilePath) throws IOException {

    String result = getPdfContent(pdfFilePath);
    String[] parts = result.split("\\s+");

    // Create a new Word document
    try (XWPFDocument wordDocument = new XWPFDocument();) {

      XWPFTable table = wordDocument.createTable();

      // Adding a title row
      XWPFTableRow titleRow = table.getRow(0);
      titleRow.getCell(0)
          .setText("扣款/儲值日期");
      titleRow.addNewTableCell()
          .setText("扣款總額");
      titleRow.addNewTableCell()
          .setText("儲值總額");
      titleRow.addNewTableCell()
          .setText("帳戶餘額");
      titleRow.addNewTableCell()
          .setText("試算餘額");

      // Fill the table with data extracted from the PDF

      Pattern pattern = Pattern.compile("\\d{4}/\\d{2}/\\d{2}"); // 啟用多行模式

      Pattern ignore = Pattern.compile("\\d{4}\\.\\d{2}\\.\\d{2}"); // Pattern for "2024.04.26"
      XWPFTableRow row = table.createRow();

      for (String part : parts) {

        if (part.matches(ignore.pattern())
            || part.contains("遠通電收股份有限公司電子憑證專用章")
            || part.contains("－")) {
          continue;
        }

        if (part.matches(pattern.pattern())) {
          row.getCell(0)
              .setText(part);  // Date
          continue;
        }

        if (part.contains("折優惠")) {
          row.getCell(1)
              .setText(part);  // Deduction Total
          continue;
        }

        if (part.contains("元試算餘額低於")) {
          row.getCell(2)
              .setText(part);  // Recharge Total
          continue;
        }

        if (part.equals("0元")) {
          row.getCell(1)
              .setText(part);  // Deduction Total
          continue;
        }

        if (part.endsWith("元")) {
          row.getCell(3)
              .setText(part);  // Account Balance
          row = table.createRow(); //換行
        }


      }

      generateWordFile(wordDocument,pdfFilePath);
    }


  }

  private void generateWordFile(XWPFDocument wordDocument,String pdfFilePath) throws IOException {
    File file = new File(pdfFilePath);
    String nowDateTimeString = LocalDateTime.now()
        .format(DateTimeFormatter.ofPattern("yyyyMMddHHmmssSSS"));
    String userParentFilePath = file.getParent();
    String userFileName = file.getName();
    // Save the Word document
    FileOutputStream out = new FileOutputStream(
        userParentFilePath + "/" + userFileName + nowDateTimeString
            + ".docx");
    wordDocument.write(out);
    out.close();
    wordDocument.close();
  }

  private String getPdfContent(String pdfFilePath) throws IOException {
    String result;
    File file = new File(pdfFilePath);
    PDFParser parser = new PDFParser(new RandomAccessFile(file, "rw"));
    // 对PDF文件进行解析
    parser.parse();
    // 获取解析后得到的PDF文档对象
    PDDocument pdfdocument = parser.getPDDocument();
    // 新建一个PDF文本剥离器
    PDFTextStripper stripper = new PDFTextStripper();
    stripper.setSortByPosition(false); //sort:设置为true 则按照行进行读取，默认是false
    // 从PDF文档对象中剥离文本
    result = stripper.getText(pdfdocument);

    result = result.substring(result.indexOf("試算餘") + 10,
                              result.length() - 1);
    result = result.replace("\n", "")
        .replace("\r", "");
    result = result.replaceAll("(.*?)(\\d{4}/\\d{2}/\\d{2})", "$1\n$2 ");
    result = result.replaceAll("元", "元 ");
    result = result.replaceAll("\\s+(\\d+折優惠)", "$1");
    result = result.replaceAll("\\s+(試算餘額低於)", "$1");
    return result;
  }
}

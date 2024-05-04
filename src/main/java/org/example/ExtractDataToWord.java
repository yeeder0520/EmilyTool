package org.example;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;
import org.apache.pdfbox.io.RandomAccessFile;
import org.apache.pdfbox.pdfparser.PDFParser;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class ExtractDataToWord {

  Map<String, List<String>> carNumberData = new HashMap<>();

  public void execute(File[] pdfFiles) throws IOException {
    // Initialize the resultDoc outside the loop
    initCarNumberMap();

    try (XWPFDocument resultDoc = new XWPFDocument()) {

      for (File pdfFile : pdfFiles) {
        getPdfContent(pdfFile.getAbsolutePath());
      }

      carNumberData.forEach((key, value) -> {
        XWPFTable table = null;
        try {

          if (!value.isEmpty()) {
            // 在第一行印出車牌號碼
            resultDoc.createParagraph()
                .createRun()
                .setText("車牌號碼: " + key);
            table = resultDoc.createTable();
            setupTable(table);
          }
          for (String content : value) {

            XWPFTableRow row = table.createRow();

            String[] parts = content.split("\\s+");

            fillTable(parts, row, table);
            resultDoc.createParagraph()
                .createRun()
                .addBreak();
          }

        } catch (Exception e) {
          e.printStackTrace();
        }
      });
      saveDocument(resultDoc, pdfFiles[0].getParent());
    }
  }

  private void initCarNumberMap() {
    carNumberData.put("AMD-3275", new ArrayList<>());
    carNumberData.put("AHT-9309", new ArrayList<>());
    carNumberData.put("BDC-7297", new ArrayList<>());
    carNumberData.put("0185-NB", new ArrayList<>());
  }

  private void setupTable(XWPFTable table) {
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
  }

  private void fillTable(String[] parts, XWPFTableRow row, XWPFTable table) {
    Pattern pattern = Pattern.compile("\\d{4}/\\d{2}/\\d{2}"); // 啟用多行模式
    Pattern ignore = Pattern.compile("\\d{4}\\.\\d{2}\\.\\d{2}"); // Pattern for "2024.04.26"

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
  }

  private void saveDocument(XWPFDocument doc, String parentDirectory)
      throws IOException {
    String timestamp = LocalDateTime.now()
        .format(DateTimeFormatter.ofPattern("yyyyMMddHHmmssSSS"));
    try (FileOutputStream out = new FileOutputStream(new File(parentDirectory,
                                                              "結果檔案"
                                                                  + timestamp
                                                                  + ".docx"))) {
      doc.write(out);
    }
    doc.close();
  }

  private void getPdfContent(String pdfFilePath) throws IOException {
    String result;
    File file = new File(pdfFilePath);
    PDFParser parser = new PDFParser(new RandomAccessFile(file, "rw"));
    parser.parse();

    String carNumber;

    try (PDDocument pdfDocument = parser.getPDDocument()) {
      PDFTextStripper stripper = new PDFTextStripper();
      result = stripper.getText(pdfDocument);

      //取得車牌號碼
      carNumber = result.substring(result.indexOf("自負一切法律責任") + 11, result.indexOf(" 餘額紀錄列印"));

      if (carNumberData.containsKey(carNumber)) { //取得車牌號碼
        result = result.substring(result.indexOf("試算餘") + 10,
                                  result.length() - 1);
        result = result.replace("\n", "")
            .replace("\r", "");
        result = result.replaceAll("(.*?)(\\d{4}/\\d{2}/\\d{2})", "$1\n$2 ");
        result = result.replaceAll("元", "元 ");
        result = result.replaceAll("\\s+(\\d+折優惠)", "$1");
        result = result.replaceAll("\\s+(試算餘額低於)", "$1");
        carNumberData.get(carNumber)
            .add(result);
      }
    }
  }
}

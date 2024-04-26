package org.example;

import java.io.File;
import java.io.IOException;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

public class Main {

  public static void main(String[] args) {
    // 創建框架
    JFrame frame = new JFrame("PDF File Chooser for Multiple Files");
    frame.setSize(400, 400);
    frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

    // 創建按鈕
    JButton button = new JButton("請點我，並選擇阿姨您要的PDF");
    button.addActionListener(e -> {
      JFileChooser fileChooser = new JFileChooser();
      fileChooser.setAcceptAllFileFilterUsed(false);
      fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
          "PDF Documents",
          "pdf"));
      fileChooser.setMultiSelectionEnabled(true);  // 啟用多選模式
      int option = fileChooser.showOpenDialog(frame);
      if (option == JFileChooser.APPROVE_OPTION) {
        File[] selectedFiles = fileChooser.getSelectedFiles();
        for (File file : selectedFiles) {
          // 處理每個選擇的檔案
          JOptionPane.showMessageDialog(frame,
                                        "上傳成功，檔案會產出在跟您選擇檔案一樣的位置唷！♥♥");
          // 這裡可以添加調用處理 PDF 檔案的方法
          try {
            handleFile(file);
          } catch (IOException ignored) {
          }
        }
      }
      JOptionPane.showMessageDialog(frame, "執行完畢");

    });

    // 添加按鈕到框架
    frame.getContentPane()
        .add(button);
    frame.setVisible(true);
  }

  private static void handleFile(File file) throws IOException {
    // 在這裡處理文件，例如讀取或處理 PDF
    ExtractDataToWord extractor = new ExtractDataToWord();
    extractor.execute(file.getAbsolutePath());
  }


}
package jp.ne.yonem;

import com.formdev.flatlaf.themes.FlatMacDarkLaf;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetAdapter;
import java.awt.dnd.DropTargetDropEvent;
import java.io.File;
import java.util.List;
import javax.swing.*;
import jp.ne.yonem.components.HistoryPathComboBox;
import jp.ne.yonem.components.TreeConsolePanel;
import jp.ne.yonem.util.ExcelUtil;
import jp.ne.yonem.util.TextTreeUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/** メインクラス */
public class FileTreeFrame extends JFrame {

  private static final Logger logger = LoggerFactory.getLogger(Logger.ROOT_LOGGER_NAME);

  public static void main(String[] args) {
    FlatMacDarkLaf.setup();
    new FileTreeFrame();
  }

  /** アプリケーションのタイトル */
  private final String APP_TITLE = "FileTree";

  /** アプリケーションの幅 */
  private final int APP_WIDTH = 500;

  /** アプリケーションの高さ */
  private final int APP_HEIGHT = 800;

  /** 正常終了時のタイトル */
  private static final String SUCCESS_TITLE = "正常終了";

  /** 正常終了時のメッセージ */
  private static final String SUCCESS_MESSAGE = "ファイルを出力しました";

  /** 異常終了時のタイトル */
  private final String FAILURE_TITLE = "異常終了";

  /** 異常終了時のメッセージ */
  private final String FAILURE_MESSAGE = "処理中に例外が発生しました。ログを確認してください";

  /** ルートディレクトリテキストボックスのラベル */
  private final JLabel lblFile = new JLabel("ルートディレクトリ");

  /** アプリケーションのコンソール */
  private final TreeConsolePanel consolePanel = new TreeConsolePanel(new Insets(10, 10, 10, 10));

  /** 実行ボタン */
  private final JButton btnSubmit = new JButton("出力");

  /** Excel出力モードチェックボックス */
  private final JCheckBox chkExcel = new JCheckBox("Excel", false);

  /** ディレクトリのみ出力チェックボックス */
  private final JCheckBox chkDirectoryOnly = new JCheckBox("ディレクトリのみ", false);

  /** 選択中フォルダコンボボックス */
  private final HistoryPathComboBox comboRootDirectory = new HistoryPathComboBox();

  /** プログレスバー */
  private final JProgressBar progressBar = new JProgressBar();

  public FileTreeFrame() {
    super();

    try {
      setTitle(APP_TITLE);
      setResizable(true);
      setMinimumSize(new Dimension(APP_WIDTH, APP_HEIGHT));
      setSize(APP_WIDTH, APP_HEIGHT);
      setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
      setLocationRelativeTo(null);
      var panel = new JPanel();
      panel.setLayout(new BorderLayout());
      add(panel);

      // NORTH
      var northPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
      northPanel.add(lblFile);

      comboRootDirectory.addActionListener(
          e -> {
            if (!comboRootDirectory.isUpdating()
                && "comboBoxChanged".equals(e.getActionCommand())) {
              onSubmit();
            }
          });
      northPanel.add(comboRootDirectory);
      northPanel.add(chkExcel);
      northPanel.add(chkDirectoryOnly);
      panel.add(northPanel, BorderLayout.NORTH);

      // CENTER
      var centerPanel = new JPanel(new BorderLayout());
      centerPanel.add(consolePanel, BorderLayout.CENTER);

      progressBar.setVisible(false);
      progressBar.setStringPainted(true);
      progressBar.setString("処理中...");
      centerPanel.add(progressBar, BorderLayout.SOUTH);
      panel.add(centerPanel, BorderLayout.CENTER);

      // SOUTH
      var southPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
      southPanel.add(btnSubmit);
      btnSubmit.addActionListener(e -> onSubmit());
      panel.add(southPanel, BorderLayout.SOUTH);

      var dt =
          new DropTarget(
              this,
              DnDConstants.ACTION_COPY,
              new DropTargetAdapter() {
                @Override
                public void drop(DropTargetDropEvent event) {
                  try {
                    event.acceptDrop(DnDConstants.ACTION_COPY);
                    var transferData =
                        event.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);

                    if (transferData instanceof List<?> files && !files.isEmpty()) {
                      var file = (File) files.getFirst();
                      comboRootDirectory.setSelectedItem(file.getAbsolutePath());
                    }
                  } catch (Exception e) {
                    logger.error("Drop failed", e);
                  }
                }
              });
      consolePanel.setDropTarget(dt);
      comboRootDirectory.setDropTarget(dt);
      comboRootDirectory.getEditor().getEditorComponent().setDropTarget(dt);
      panel.setDropTarget(dt);
      setVisible(true);

    } catch (Exception e) {
      logger.error(null, e);
      JOptionPane.showMessageDialog(
          null, FAILURE_MESSAGE, FAILURE_TITLE, JOptionPane.ERROR_MESSAGE);
    }
  }

  /** 出力ボタン押下時の処理 */
  private void onSubmit() {
    var path = comboRootDirectory.getSelectedPath();
    if (path.isEmpty()) return;

    btnSubmit.setEnabled(false);
    progressBar.setVisible(true);
    progressBar.setIndeterminate(true);

    new SwingWorker<String, Void>() {

      @Override
      protected String doInBackground() throws Exception {
        var rootDirectory = new File(path);

        if (chkExcel.isSelected()) {
          ExcelUtil.convertDir2Tree(rootDirectory, chkDirectoryOnly.isSelected());

          if (Desktop.isDesktopSupported()) {
            Desktop.getDesktop().open(rootDirectory);
          }
          return SUCCESS_MESSAGE;

        } else {
          return TextTreeUtil.convertDir2Text(rootDirectory, chkDirectoryOnly.isSelected());
        }
      }

      @Override
      protected void done() {
        try {
          var result = get();

          if (chkExcel.isSelected()) {
            consolePanel.setText("Start!!\n");
            consolePanel.append(result + "\nEnd!!");
            JOptionPane.showMessageDialog(
                null, result, SUCCESS_TITLE, JOptionPane.INFORMATION_MESSAGE);

          } else {
            consolePanel.setText(result);
          }
          comboRootDirectory.saveHistory(path);

        } catch (Exception e) {
          logger.error("処理失敗", e);
          consolePanel.setText("エラー: " + e.getMessage());

        } finally {
          btnSubmit.setEnabled(true);
          progressBar.setVisible(false);
          progressBar.setIndeterminate(false);
        }
      }
    }.execute();
  }
}

package jp.ne.yonem;

import com.formdev.flatlaf.themes.FlatMacDarkLaf;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.StringSelection;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetAdapter;
import java.awt.dnd.DropTargetDropEvent;
import java.io.File;
import java.util.List;
import java.util.Objects;
import javax.swing.*;
import jp.ne.yonem.components.HistoryPathComboBox;
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
  private final JTextArea taConsole = new JTextArea();

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

  /** コピーボタン */
  private final JButton btnCopy = new JButton();

  /** デフォルトマージン */
  private final Insets defaultInsets = new Insets(10, 10, 10, 10);

  public FileTreeFrame() {
    super();

    try {
      setTitle(APP_TITLE);
      setResizable(false);
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
      taConsole.setEditable(false);
      taConsole.setMargin(defaultInsets);
      var scrollPane = new JScrollPane(taConsole);

      btnCopy.setText("📋");
      btnCopy.setToolTipText("クリップボードにコピー");
      btnCopy.setFocusable(false);
      btnCopy.setCursor(new Cursor(Cursor.HAND_CURSOR));
      btnCopy.addActionListener(e -> copyToClipboard());
      btnCopy.putClientProperty("JButton.buttonType", "toolBarButton");

      var layeredPane = new JLayeredPane();
      layeredPane.setLayout(
          new LayoutManager() {
            @Override
            public void addLayoutComponent(String name, Component comp) {}

            @Override
            public void removeLayoutComponent(Component comp) {}

            @Override
            public Dimension preferredLayoutSize(Container parent) {
              return scrollPane.getPreferredSize();
            }

            @Override
            public Dimension minimumLayoutSize(Container parent) {
              return scrollPane.getMinimumSize();
            }

            @Override
            public void layoutContainer(Container parent) {
              scrollPane.setBounds(0, 0, parent.getWidth(), parent.getHeight());
              int btnWidth = 40;
              int btnHeight = 30;
              btnCopy.setBounds(parent.getWidth() - btnWidth - 25, 10, btnWidth, btnHeight);
            }
          });

      layeredPane.add(scrollPane, JLayeredPane.DEFAULT_LAYER);
      layeredPane.add(btnCopy, JLayeredPane.PALETTE_LAYER);
      centerPanel.add(layeredPane, BorderLayout.CENTER);

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
      taConsole.setDropTarget(dt);
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
            taConsole.setText("Start!!\n");
            taConsole.append(result + "\nEnd!!");
            JOptionPane.showMessageDialog(
                null, result, SUCCESS_TITLE, JOptionPane.INFORMATION_MESSAGE);

          } else {
            taConsole.setText(result);
          }
          comboRootDirectory.saveHistory(path);

        } catch (Exception e) {
          logger.error("処理失敗", e);
          taConsole.setText("エラー: " + e.getMessage());

        } finally {
          btnSubmit.setEnabled(true);
          progressBar.setVisible(false);
          progressBar.setIndeterminate(false);
        }
      }
    }.execute();
  }

  private Timer copyTimer;

  /** クリップボードへのコピー処理 */
  private void copyToClipboard() {
    var text = taConsole.getText();
    if (Objects.isNull(text) || text.isEmpty()) return;

    var selection = new StringSelection(text);
    Toolkit.getDefaultToolkit().getSystemClipboard().setContents(selection, null);
    if (Objects.nonNull(copyTimer) && copyTimer.isRunning()) copyTimer.stop();

    btnCopy.setText("✅");

    copyTimer =
        new Timer(
            1500,
            e -> {
              btnCopy.setText("📋");
              copyTimer.stop();
            });
    copyTimer.setRepeats(false);
    copyTimer.start();
  }
}

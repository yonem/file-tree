package jp.ne.yonem;

import com.formdev.flatlaf.themes.FlatMacDarkLaf;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetAdapter;
import java.awt.dnd.DropTargetDropEvent;
import java.io.File;
import java.nio.file.Files;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
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

  /** タブ表示 */
  private final JTabbedPane tabbedPane = new JTabbedPane();

  /** アプリケーションのコンソール */
  private final TreeConsolePanel consolePanel = new TreeConsolePanel(new Insets(10, 10, 10, 10));

  /** 統計情報表示テーブル */
  private final JTable statsTable = new JTable();

  private final DefaultTableModel tableModel =
      new DefaultTableModel(new Object[] {"拡張子", "ファイル数", "総行数(LOC)", "平均行数"}, 0);

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
      statsTable.setModel(tableModel);
      tabbedPane.addTab("ツリー表示", consolePanel);
      tabbedPane.addTab("拡張子統計", new JScrollPane(statsTable));
      centerPanel.add(tabbedPane, BorderLayout.CENTER);

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

    record SearchResult(String treeText, List<ExtensionStat> stats) {}

    new SwingWorker<SearchResult, Void>() {

      @Override
      protected SearchResult doInBackground() throws Exception {
        var root = new File(path);
        var isDirOnly = chkDirectoryOnly.isSelected();
        var treeText = "";

        if (chkExcel.isSelected()) {
          ExcelUtil.convertDir2Tree(root, isDirOnly);
          treeText = SUCCESS_MESSAGE;

          if (Desktop.isDesktopSupported()) {
            Desktop.getDesktop().open(root);
          }
        } else {
          treeText = TextTreeUtil.convertDir2Text(root, isDirOnly);
        }

        try (var stream = Files.walk(root.toPath())) {
          var statsMap =
              stream
                  .filter(Files::isRegularFile)
                  .collect(
                      Collectors.groupingBy(
                          p -> {
                            var name = p.getFileName().toString();
                            var dotIndex = name.lastIndexOf('.');
                            return dotIndex == -1
                                ? "(no extension)"
                                : name.substring(dotIndex).toLowerCase();
                          }));

          var statsList =
              statsMap.entrySet().stream()
                  .map(
                      entry -> {
                        var totalLines =
                            entry.getValue().stream()
                                .mapToLong(
                                    p -> {
                                      try (var lines = Files.lines(p)) {
                                        return lines.count();
                                      } catch (Exception e) {
                                        return 0;
                                      }
                                    })
                                .sum();
                        return new ExtensionStat(
                            entry.getKey(), entry.getValue().size(), totalLines);
                      })
                  .sorted(Comparator.comparingLong(ExtensionStat::count).reversed())
                  .toList();

          return new SearchResult(treeText, statsList);
        }
      }

      @Override
      protected void done() {
        try {
          var result = get();
          consolePanel.setText(result.treeText());

          tableModel.setRowCount(0);
          result
              .stats()
              .forEach(
                  s ->
                      tableModel.addRow(
                          new Object[] {
                            s.extension(), s.count(), s.totalLines(), s.getAverageLines()
                          }));

          if (chkExcel.isSelected()) {
            JOptionPane.showMessageDialog(
                null, result.treeText(), SUCCESS_TITLE, JOptionPane.INFORMATION_MESSAGE);
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

  /**
   * 拡張子ごとの統計情報
   *
   * @param extension 拡張子名
   * @param count ファイル数
   * @param totalLines 総行数
   */
  public record ExtensionStat(String extension, long count, long totalLines) {

    /**
     * 平均行数を取得する
     *
     * @return 平均行数
     */
    public long getAverageLines() {
      return count == 0 ? 0 : totalLines / count;
    }
  }
}

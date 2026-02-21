package jp.ne.yonem;

import static jp.ne.yonem.util.ExcelUtil.convertDir2Tree;

import com.formdev.flatlaf.themes.FlatMacDarkLaf;
import java.awt.*;
import java.io.File;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Objects;
import javax.swing.*;
import jp.ne.yonem.components.SelectedFileTextField;
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

  /** 選択中フォルダ */
  private final JTextField txtRootDirectory = new SelectedFileTextField();

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
      northPanel.add(txtRootDirectory);
      northPanel.add(chkExcel);
      northPanel.add(chkDirectoryOnly);
      panel.add(northPanel, BorderLayout.NORTH);

      // CENTER
      taConsole.setEditable(false);
      taConsole.setMargin(defaultInsets);
      panel.add(new JScrollPane(taConsole));

      // SOUTH
      var southPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
      southPanel.add(btnSubmit);
      btnSubmit.addActionListener(e -> onSubmit());
      panel.add(southPanel, BorderLayout.SOUTH);
      setVisible(true);

    } catch (Exception e) {
      logger.error(null, e);
      JOptionPane.showMessageDialog(
          null, FAILURE_MESSAGE, FAILURE_TITLE, JOptionPane.ERROR_MESSAGE);
    }
  }

  /** 出力ボタン押下時の処理 */
  private void onSubmit() {
    var path = txtRootDirectory.getText();
    if (path.isEmpty()) return;

    try {
      new SwingWorker<Void, Void>() {

        @Override
        protected Void doInBackground() {

          try {
            btnSubmit.setEnabled(false);
            var rootDirectory = new File(path);

            if (chkExcel.isSelected()) {
              taConsole.setText("Start!!\n");
              convertDir2Tree(rootDirectory);
              JOptionPane.showMessageDialog(
                  null, SUCCESS_MESSAGE, SUCCESS_TITLE, JOptionPane.INFORMATION_MESSAGE);
              taConsole.append(SUCCESS_MESSAGE);
              taConsole.append("\n");
              taConsole.append("End!!");
              return null;
            }
            outputConsole(rootDirectory, 0, "", false);

          } catch (Exception e) {
            taConsole.setText(e.getMessage());

          } finally {
            btnSubmit.setEnabled(true);
          }
          return null;
        }
      }.execute();

    } catch (Exception e) {
      JOptionPane.showMessageDialog(
          null, FAILURE_MESSAGE, FAILURE_TITLE, JOptionPane.ERROR_MESSAGE);
    }
  }

  /**
   * コンソールにディレクトリ構造を出力する
   *
   * @param file 対象ファイルまたはディレクトリ
   * @param indent インデント深さ
   * @param hierarchy 階層文字列
   * @param isEOL 最終行フラグ
   */
  private void outputConsole(File file, int indent, String hierarchy, boolean isEOL) {

    if (chkDirectoryOnly.isSelected() && file.isFile()) return;

    if (0 == indent) {
      taConsole.setText(null);
      taConsole.append(file.getName());
      taConsole.append("\n");
      hierarchy += "   ";

    } else {
      taConsole.append(hierarchy);
      taConsole.append(isEOL ? "└─ " : "├─ ");
      taConsole.append(file.getName());
      taConsole.append("\n");
    }
    if (file.isFile()) return;

    var lists = file.listFiles();
    if (Objects.isNull(lists)) return;

    var filteredLists =
        Arrays.stream(lists)
            .filter(f -> !chkDirectoryOnly.isSelected() || f.isDirectory())
            .sorted(Comparator.comparing(File::isDirectory).reversed().thenComparing(File::getName))
            .toArray(File[]::new);

    for (var i = 0; i < filteredLists.length; i++) {
      var next = filteredLists[i];
      var currentHierarchy = hierarchy;
      if (i == 0 && 0 < indent) currentHierarchy += isEOL ? "   " : "│  ";
      var isLast = i == filteredLists.length - 1;
      outputConsole(next, indent + 1, currentHierarchy, isLast);
    }
  }
}

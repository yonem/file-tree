package jp.ne.yonem;

import static java.util.prefs.Preferences.userNodeForPackage;

import com.formdev.flatlaf.themes.FlatMacDarkLaf;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.StringSelection;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetAdapter;
import java.awt.dnd.DropTargetDropEvent;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.prefs.Preferences;
import javax.swing.*;
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

  /** 選択中フォルダ */
  private final JComboBox<String> comboRootDirectory = new JComboBox<>();

  private static final String PREF_KEY_HISTORY = "path_history";
  private static final int MAX_HISTORY = 10;

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
      comboRootDirectory.setEditable(true);
      comboRootDirectory.setPreferredSize(new Dimension(230, 25));
      var editorComponent = comboRootDirectory.getEditor().getEditorComponent();

      editorComponent.addMouseListener(
          new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent e) {
              showFileChooser();
            }
          });

      editorComponent.addKeyListener(
          new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent e) {
              if (e.getKeyCode() == KeyEvent.VK_ENTER) {
                showFileChooser();
              }
            }
          });
      comboRootDirectory.addActionListener(
          e -> {
            if (!isUpdatingHistory && "comboBoxChanged".equals(e.getActionCommand())) {
              var item = comboRootDirectory.getSelectedItem();
              if (Objects.nonNull(item) && !item.toString().isEmpty()) onSubmit();
            }
          });

      loadHistory();
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
    var path = (String) comboRootDirectory.getEditor().getItem();
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
          saveHistory(path);

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

  /** 履歴をPreferencesから読み込む */
  private void loadHistory() {
    isUpdatingHistory = true;

    try {
      var prefs = Preferences.userNodeForPackage(FileTreeFrame.class);
      var historyRaw = prefs.get(PREF_KEY_HISTORY, "");

      comboRootDirectory.removeAllItems();
      if (!historyRaw.isEmpty()) {
        for (var path : historyRaw.split(",")) {
          comboRootDirectory.addItem(path);
        }
      }
      comboRootDirectory.setSelectedIndex(-1);

    } finally {
      isUpdatingHistory = false;
    }
  }

  private boolean isUpdatingHistory = false;

  /** 成功したパスを履歴に保存する */
  private void saveHistory(String newPath) {
    if (newPath.isEmpty() || isUpdatingHistory) return;
    isUpdatingHistory = true;

    try {
      var prefs = userNodeForPackage(FileTreeFrame.class);
      var historyList = new ArrayList<String>();

      for (int i = 0; i < comboRootDirectory.getItemCount(); i++) {
        historyList.add(comboRootDirectory.getItemAt(i));
      }
      historyList.remove(newPath);
      historyList.addFirst(newPath);

      if (historyList.size() > MAX_HISTORY)
        historyList = new ArrayList<>(historyList.subList(0, MAX_HISTORY));

      comboRootDirectory.removeAllItems();
      var sb = new StringBuilder();

      for (var path : historyList) {
        comboRootDirectory.addItem(path);
        if (!sb.isEmpty()) sb.append(",");
        sb.append(path);
      }
      prefs.put(PREF_KEY_HISTORY, sb.toString());
      comboRootDirectory.setSelectedItem(newPath);

    } finally {
      isUpdatingHistory = false;
    }
  }

  /** ディレクトリ選択ダイアログを表示し、選択されたパスをコンボボックスにセットする */
  private void showFileChooser() {
    var currentPath = (String) comboRootDirectory.getEditor().getItem();
    var chooser = new JFileChooser(currentPath);

    chooser.setMultiSelectionEnabled(false);
    chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
    chooser.setAcceptAllFileFilterUsed(false);

    var selected = chooser.showOpenDialog(this);

    if (selected == JFileChooser.APPROVE_OPTION) {
      var path = chooser.getSelectedFile().getAbsolutePath();
      comboRootDirectory.getEditor().setItem(path);
      onSubmit();
    }
  }
}

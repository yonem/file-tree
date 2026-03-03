package jp.ne.yonem.components;

import java.awt.*;
import java.awt.datatransfer.StringSelection;
import java.util.Objects;
import javax.swing.*;
import javax.swing.text.*;

public class TreeConsolePanel extends JPanel {

  private final JTextPane consolePane = new JTextPane();
  private final JButton btnCopy = new JButton("📋");
  private Timer copyTimer;

  public TreeConsolePanel(Insets insets) {
    setLayout(new BorderLayout());

    consolePane.setEditable(false);
    consolePane.setBackground(new Color(30, 30, 30));
    consolePane.setForeground(new Color(220, 220, 220));
    consolePane.setFont(new Font("Monospaced", Font.PLAIN, 13));
    consolePane.setMargin(insets);

    var scrollPane = new JScrollPane(consolePane);
    scrollPane.setBorder(BorderFactory.createEmptyBorder());

    btnCopy.setFocusable(false);
    btnCopy.setCursor(new Cursor(Cursor.HAND_CURSOR));
    btnCopy.putClientProperty("JButton.buttonType", "toolBarButton");
    btnCopy.addActionListener(e -> copyToClipboard());

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
            btnCopy.setBounds(parent.getWidth() - 65, 10, 40, 30);
          }
        });

    layeredPane.add(scrollPane, JLayeredPane.DEFAULT_LAYER);
    layeredPane.add(btnCopy, JLayeredPane.PALETTE_LAYER);
    add(layeredPane, BorderLayout.CENTER);
  }

  /** テキストのクリアとセット */
  public void setText(String text) {
    consolePane.setText("");
    appendStyledText(text);
  }

  /** スタイル付きテキストの追加 */
  public void append(String text) {
    appendStyledText(text);
  }

  /** ディレクトリとファイルを判別して色分けするロジック */
  private void appendStyledText(String text) {
    var doc = consolePane.getStyledDocument();

    var dirStyle = consolePane.addStyle("Directory", null);
    StyleConstants.setForeground(dirStyle, new Color(86, 156, 214));
    StyleConstants.setBold(dirStyle, true);

    var fileStyle = consolePane.addStyle("File", null);
    StyleConstants.setForeground(fileStyle, new Color(156, 220, 254));

    var defaultStyle = consolePane.addStyle("Default", null);
    StyleConstants.setForeground(defaultStyle, new Color(200, 200, 200));

    try {
      for (String line : text.split("\n")) {

        if (line.contains("📁") || line.endsWith("/")) {
          doc.insertString(doc.getLength(), line + "\n", dirStyle);

        } else if (line.contains("└─") || line.contains("├─")) {
          doc.insertString(doc.getLength(), line + "\n", fileStyle);

        } else {
          doc.insertString(doc.getLength(), line + "\n", defaultStyle);
        }
      }

    } catch (BadLocationException e) {
      e.printStackTrace();
    }
  }

  public String getText() {
    return consolePane.getText();
  }

  @Override
  public void setDropTarget(java.awt.dnd.DropTarget dt) {
    consolePane.setDropTarget(dt);
  }

  private void copyToClipboard() {
    var text = getText();
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

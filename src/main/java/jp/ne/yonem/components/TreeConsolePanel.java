package jp.ne.yonem.components;

import java.awt.*;
import java.awt.datatransfer.StringSelection;
import java.awt.dnd.DropTarget;
import java.util.Objects;
import javax.swing.*;

/** コピーボタンが右上に重なるコンソールパネル */
public class TreeConsolePanel extends JPanel {

  private final JTextArea taConsole = new JTextArea();
  private final JButton btnCopy = new JButton("📋");
  private Timer copyTimer;

  public TreeConsolePanel(Insets insets) {
    setLayout(new BorderLayout());

    taConsole.setEditable(false);
    taConsole.setMargin(insets);
    var scrollPane = new JScrollPane(taConsole);

    btnCopy.setToolTipText("クリップボードにコピー");
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

  public void setText(String text) {
    taConsole.setText(text);
  }

  public void append(String text) {
    taConsole.append(text);
  }

  public void setDropTarget(DropTarget dt) {
    taConsole.setDropTarget(dt);
  }

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

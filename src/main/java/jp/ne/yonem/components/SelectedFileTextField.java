package jp.ne.yonem.components;

import java.awt.*;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import javax.swing.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class SelectedFileTextField extends JTextField {

  private static final Logger logger = LoggerFactory.getLogger(Logger.ROOT_LOGGER_NAME);

  public SelectedFileTextField() {
    init();
  }

  /** テキストフィールドの初期化 */
  private void init() {
    this.setEditable(false);
    this.setColumns(20);
    this.addMouseListener(onSelectedFileClick());
    this.addKeyListener(onSelectedFileEnter());
  }

  /**
   * 選択中ファイルフォーカス時にエンターキーを押下した時の処理
   *
   * @return キーイベントリスナー
   */
  private KeyListener onSelectedFileEnter() {
    return new KeyListener() {

      @Override
      public void keyTyped(KeyEvent e) {
        if (e.getKeyChar() == KeyEvent.VK_ENTER) onSelectedFile();
      }

      @Override
      public void keyPressed(KeyEvent e) {}

      @Override
      public void keyReleased(KeyEvent e) {}
    };
  }

  /**
   * 選択中ファイルをクリックした時の処理
   *
   * @return マウスイベントリスナー
   */
  private MouseListener onSelectedFileClick() {
    return new MouseListener() {

      @Override
      public void mouseClicked(MouseEvent e) {
        onSelectedFile();
      }

      @Override
      public void mousePressed(MouseEvent e) {}

      @Override
      public void mouseReleased(MouseEvent e) {}

      @Override
      public void mouseEntered(MouseEvent e) {}

      @Override
      public void mouseExited(MouseEvent e) {}
    };
  }

  /** ディレクトリ選択処理 */
  private void onSelectedFile() {
    var chooser = new JFileChooser(this.getText());
    chooser.setMultiSelectionEnabled(false);
    chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
    chooser.setAcceptAllFileFilterUsed(false);
    var selected = chooser.showOpenDialog(null);
    if (selected != JFileChooser.APPROVE_OPTION) return;
    var path = chooser.getSelectedFile().getAbsolutePath();
    logger.debug("選択ディレクトリ：{}", path);
    this.setText(path);
  }
}

package jp.ne.yonem.components;

import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.util.ArrayList;
import java.util.Objects;
import java.util.prefs.Preferences;
import javax.swing.*;

/** 履歴管理機能とディレクトリ選択機能を備えたコンボボックス */
public class HistoryPathComboBox extends JComboBox<String> {

  private static final String PREF_KEY_HISTORY = "path_history";
  private static final int MAX_HISTORY = 10;
  private boolean isUpdating = false;

  public HistoryPathComboBox() {
    setEditable(true);
    setPreferredSize(new Dimension(230, 25));
    initEvents();
    loadHistory();
  }

  private void initEvents() {
    var editor = getEditor().getEditorComponent();

    editor.addMouseListener(
        new MouseAdapter() {
          @Override
          public void mouseClicked(MouseEvent e) {
            showFileChooser();
          }
        });

    editor.addKeyListener(
        new KeyAdapter() {
          @Override
          public void keyPressed(KeyEvent e) {
            if (e.getKeyCode() == KeyEvent.VK_ENTER) {
              showFileChooser();
            }
          }
        });
  }

  /** 履歴をPreferencesから読み込む */
  private void loadHistory() {
    isUpdating = true;
    try {
      var prefs = Preferences.userNodeForPackage(HistoryPathComboBox.class);
      var historyRaw = prefs.get(PREF_KEY_HISTORY, "");

      removeAllItems();

      if (!historyRaw.isEmpty()) {
        for (var path : historyRaw.split(",")) {
          addItem(path);
        }
      }
      setSelectedIndex(-1);

    } finally {
      isUpdating = false;
    }
  }

  /** 成功したパスを履歴に保存する */
  public void saveHistory(String newPath) {
    if (Objects.isNull(newPath) || newPath.isEmpty() || isUpdating) return;
    isUpdating = true;

    try {
      var prefs = Preferences.userNodeForPackage(HistoryPathComboBox.class);
      var historyList = new ArrayList<String>();

      for (int i = 0; i < getItemCount(); i++) {
        historyList.add(getItemAt(i));
      }
      historyList.remove(newPath);
      historyList.addFirst(newPath);

      if (historyList.size() > MAX_HISTORY) {
        historyList = new ArrayList<>(historyList.subList(0, MAX_HISTORY));
      }

      removeAllItems();
      var sb = new StringBuilder();

      for (var path : historyList) {
        addItem(path);
        if (!sb.isEmpty()) sb.append(",");
        sb.append(path);
      }
      prefs.put(PREF_KEY_HISTORY, sb.toString());
      setSelectedItem(newPath);

    } finally {
      isUpdating = false;
    }
  }

  private void showFileChooser() {
    var currentPath = (String) getEditor().getItem();
    var chooser = new JFileChooser(currentPath);
    chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

    if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
      getEditor().setItem(chooser.getSelectedFile().getAbsolutePath());
      actionPerformed(new ActionEvent(this, 0, "comboBoxChanged"));
    }
  }

  public String getSelectedPath() {
    return (String) getEditor().getItem();
  }

  public boolean isUpdating() {
    return isUpdating;
  }
}

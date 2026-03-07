package jp.ne.yonem.components;

import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.text.NumberFormat;
import java.util.List;
import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import jp.ne.yonem.FileTreeFrame.ExtensionStat;

/** 拡張子ごとの統計情報を表示するテーブルコンポーネント */
public class ExtensionStatTable extends JPanel {

  private final JTable table = new JTable();

  private final DefaultTableModel model =
      new DefaultTableModel(new Object[] {"拡張子", "ファイル数", "総行数(LOC)", "平均行数"}, 0) {

        @Override
        public Class<?> getColumnClass(int columnIndex) {
          return switch (columnIndex) {
            case 0 -> String.class;
            default -> Long.class;
          };
        }

        @Override
        public boolean isCellEditable(int row, int column) {
          return false;
        }
      };

  /** コンストラクタ */
  public ExtensionStatTable() {
    setLayout(new BorderLayout());
    table.setModel(model);
    table.setAutoCreateRowSorter(true);
    table.getTableHeader().setReorderingAllowed(false);
    table
        .getTableHeader()
        .addMouseListener(
            new MouseAdapter() {
              @Override
              public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 2) {
                  var column = table.columnAtPoint(e.getPoint());
                  var cursorType = table.getTableHeader().getCursor().getType();

                  if (column != -1 && cursorType == Cursor.E_RESIZE_CURSOR) {
                    adjustColumnWidth(column);
                  }
                }
              }
            });
    var numberRenderer = new NumberCellRenderer();

    for (var i = 1; i <= 3; i++) {
      table.getColumnModel().getColumn(i).setCellRenderer(numberRenderer);
    }
    add(new JScrollPane(table), BorderLayout.CENTER);
  }

  /**
   * 指定されたカラムの幅をコンテンツに合わせて自動調整する
   *
   * @param columnIndex カラムインデックス
   */
  private void adjustColumnWidth(int columnIndex) {
    var column = table.getColumnModel().getColumn(columnIndex);
    var headerRenderer = table.getTableHeader().getDefaultRenderer();
    var headerComp =
        headerRenderer.getTableCellRendererComponent(
            table, column.getHeaderValue(), false, false, 0, columnIndex);
    var width = headerComp.getPreferredSize().width;

    for (var row = 0; row < table.getRowCount(); row++) {
      var renderer = table.getCellRenderer(row, columnIndex);
      var comp = table.prepareRenderer(renderer, row, columnIndex);
      width = Math.max(comp.getPreferredSize().width + 10, width);
    }
    column.setPreferredWidth(width);
  }

  /** 統計情報をテーブルに反映する */
  public void updateStats(List<ExtensionStat> stats) {
    model.setRowCount(0);
    stats.forEach(
        s ->
            model.addRow(
                new Object[] {s.extension(), s.count(), s.totalLines(), s.getAverageLines()}));

    // データ更新時に全カラムを自動調整（任意）
    for (var i = 0; i < table.getColumnCount(); i++) {
      adjustColumnWidth(i);
    }
  }

  /** 数値を右揃えかつカンマ区切りで表示するためのレンダラー */
  private static class NumberCellRenderer extends DefaultTableCellRenderer {
    private final NumberFormat formatter = NumberFormat.getIntegerInstance();

    public NumberCellRenderer() {
      setHorizontalAlignment(SwingConstants.RIGHT);
    }

    @Override
    public Component getTableCellRendererComponent(
        JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
      var formattedValue = (value instanceof Number n) ? formatter.format(n) : value;
      return super.getTableCellRendererComponent(
          table, formattedValue, isSelected, hasFocus, row, column);
    }
  }
}

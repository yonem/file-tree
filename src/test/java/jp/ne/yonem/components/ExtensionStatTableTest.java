package jp.ne.yonem.components;

import static org.junit.jupiter.api.Assertions.*;

import java.awt.Container;
import java.awt.Cursor;
import java.awt.event.MouseEvent;
import java.util.List;
import java.util.Optional;
import javax.swing.JTable;
import jp.ne.yonem.FileTreeFrame.ExtensionStat;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.junit.jupiter.MockitoExtension;

/** ExtensionStatTableのテスト */
@ExtendWith(MockitoExtension.class)
class ExtensionStatTableTest {

  @InjectMocks private ExtensionStatTable sut;

  @BeforeEach
  void setUp() {
    sut = new ExtensionStatTable();
  }

  @Nested
  @DisplayName("正常系テストケース")
  class PositiveTests {

    @Test
    @DisplayName("正常系: 統計情報がテーブルに正しく反映されること")
    void test01() {
      var stats = List.of(new ExtensionStat(".java", 10, 1000));
      assertDoesNotThrow(() -> sut.updateStats(stats));
    }

    @Test
    @DisplayName("正常系: 各列のクラス型が正しく定義されていること")
    void test02() {
      var table = findComponent(sut, JTable.class).orElseThrow();
      var model = table.getModel();

      assertEquals(String.class, model.getColumnClass(0));
      assertEquals(Long.class, model.getColumnClass(1));
      assertEquals(Long.class, model.getColumnClass(2));
      assertEquals(Long.class, model.getColumnClass(3));
    }

    @Test
    @DisplayName("正常系: すべてのセルが編集不可に設定されていること")
    void test03() {
      var table = findComponent(sut, JTable.class).orElseThrow();
      var model = table.getModel();

      assertFalse(model.isCellEditable(0, 0));
      assertFalse(model.isCellEditable(0, 1));
      assertFalse(model.isCellEditable(0, 2));
      assertFalse(model.isCellEditable(0, 3));
    }

    @Test
    @DisplayName("正常系: ヘッダーのダブルクリック時に列幅の自動調整が行われること")
    void test04() {
      var stats = List.of(new ExtensionStat(".java", 1, 100));
      sut.updateStats(stats);

      var table = findComponent(sut, JTable.class).orElseThrow();
      var header = table.getTableHeader();
      header.setCursor(Cursor.getPredefinedCursor(Cursor.E_RESIZE_CURSOR));

      var event =
          new MouseEvent(
              header, MouseEvent.MOUSE_CLICKED, System.currentTimeMillis(), 0, 5, 5, 2, false);

      assertDoesNotThrow(
          () -> {
            for (var listener : header.getMouseListeners()) {
              listener.mouseClicked(event);
            }
          });
    }
  }

  @Nested
  @DisplayName("異常系テストケース")
  class NegativeTests {

    @Test
    @DisplayName("異常系: nullの統計リストを渡した際に例外がスローされること")
    void test01() {
      assertThrows(NullPointerException.class, () -> sut.updateStats(null));
    }
  }

  /**
   * コンテナ内から指定された型のコンポーネントを検索する
   *
   * @param container 検索対象のコンテナ
   * @param clazz 検索するクラス型
   * @return 見つかったコンポーネント（見つからない場合は空のOptional）
   */
  private <T> Optional<T> findComponent(Container container, Class<T> clazz) {
    for (var comp : container.getComponents()) {
      if (clazz.isInstance(comp)) {
        return Optional.of(clazz.cast(comp));
      }

      if (comp instanceof Container child) {
        var found = findComponent(child, clazz);

        if (found.isPresent()) {
          return found;
        }
      }
    }
    return Optional.empty();
  }
}

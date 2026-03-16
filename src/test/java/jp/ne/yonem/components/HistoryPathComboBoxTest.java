package jp.ne.yonem.components;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

import java.awt.Component;
import java.awt.event.KeyEvent;
import java.awt.event.MouseEvent;
import java.io.File;
import java.util.Optional;
import java.util.prefs.Preferences;
import javax.swing.JFileChooser;
import javax.swing.JTextField;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

/** HistoryPathComboBoxのテスト */
@ExtendWith(MockitoExtension.class)
class HistoryPathComboBoxTest {

  @Mock private Preferences prefs;

  @InjectMocks private HistoryPathComboBox sut;

  @BeforeEach
  void setUp() {
    try (var mockedPrefs = mockStatic(Preferences.class)) {
      mockedPrefs
          .when(() -> Preferences.userNodeForPackage(HistoryPathComboBox.class))
          .thenReturn(prefs);
      when(prefs.get(anyString(), anyString())).thenReturn("");
      sut = new HistoryPathComboBox();
    }
  }

  @Nested
  @DisplayName("正常系テストケース")
  class PositiveTests {

    @Test
    @DisplayName("正常系: マウスクリック時にファイルチューザーが表示され選択結果が反映されること")
    void test01() {
      var editor = cast(sut.getEditor().getEditorComponent(), Component.class).orElseThrow();
      var event =
          new MouseEvent(
              editor, MouseEvent.MOUSE_CLICKED, System.currentTimeMillis(), 0, 0, 0, 1, false);
      var selectedFile = mock(File.class);
      when(selectedFile.getAbsolutePath()).thenReturn("/mouse/path");

      try (var mockedChooser =
          mockConstruction(
              JFileChooser.class,
              (mock, context) -> {
                when(mock.showOpenDialog(any())).thenReturn(JFileChooser.APPROVE_OPTION);
                when(mock.getSelectedFile()).thenReturn(selectedFile);
              })) {
        assertDoesNotThrow(
            () -> {
              for (var listener : editor.getMouseListeners()) {
                listener.mouseClicked(event);
              }
            });
        assertEquals("/mouse/path", sut.getSelectedPath());
        assertFalse(sut.isUpdating());
      }
    }

    @Test
    @DisplayName("正常系: 履歴が最大数を超えた場合に古い履歴が削除されること")
    void test02() {
      try (var mockedPrefs = mockStatic(Preferences.class)) {
        mockedPrefs
            .when(() -> Preferences.userNodeForPackage(HistoryPathComboBox.class))
            .thenReturn(prefs);

        for (var i = 1; i <= 11; i++) {
          sut.saveHistory("/path/" + i);
        }

        assertEquals(10, sut.getItemCount());
        assertEquals("/path/11", sut.getItemAt(0));
        assertEquals("/path/2", sut.getItemAt(9));
        assertFalse(sut.isUpdating());
      }
    }

    @Test
    @DisplayName("正常系: エンターキー押下時にファイルチューザーで選択したパスが反映されること")
    void test03() {
      var editor = cast(sut.getEditor().getEditorComponent(), JTextField.class).orElseThrow();
      var keyEvent =
          new KeyEvent(
              editor, KeyEvent.KEY_PRESSED, System.currentTimeMillis(), 0, KeyEvent.VK_ENTER, ' ');
      var selectedFile = mock(File.class);
      when(selectedFile.getAbsolutePath()).thenReturn("/key/path");

      try (var mockedChooser =
          mockConstruction(
              JFileChooser.class,
              (mock, context) -> {
                when(mock.showOpenDialog(any())).thenReturn(JFileChooser.APPROVE_OPTION);
                when(mock.getSelectedFile()).thenReturn(selectedFile);
              })) {
        assertDoesNotThrow(
            () -> {
              for (var listener : editor.getKeyListeners()) {
                listener.keyPressed(keyEvent);
              }
            });
        assertEquals("/key/path", sut.getSelectedPath());
        assertFalse(sut.isUpdating());
      }
    }
  }

  @Nested
  @DisplayName("異常系テストケース")
  class NegativeTests {

    @Test
    @DisplayName("異常系: isUpdatingが真の間にsaveHistoryを呼び出しても処理が無視されること")
    void test01() {
      try (var mockedPrefs = mockStatic(Preferences.class)) {
        mockedPrefs
            .when(() -> Preferences.userNodeForPackage(HistoryPathComboBox.class))
            .thenReturn(prefs);

        var testSut = new HistoryPathComboBox();
        var initialPath = "/initial";
        testSut.saveHistory(initialPath);

        doAnswer(
                inv -> {
                  testSut.saveHistory("/nested/path");
                  return null;
                })
            .when(prefs)
            .put(anyString(), contains(initialPath));

        testSut.saveHistory(initialPath);

        assertEquals(1, testSut.getItemCount());
        assertEquals(initialPath, testSut.getItemAt(0));
      }
    }

    @Test
    @DisplayName("異常系: 保存するパスが空文字の場合は処理がスキップされること")
    void test02() {
      sut.saveHistory("");
      assertEquals(0, sut.getItemCount());
      assertFalse(sut.isUpdating());
    }
  }

  private <T> Optional<T> cast(Object component, Class<T> clazz) {
    return Optional.ofNullable(component).filter(clazz::isInstance).map(clazz::cast);
  }
}

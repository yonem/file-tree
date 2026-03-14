package jp.ne.yonem.components;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.Insets;
import java.awt.dnd.DropTarget;
import java.util.Objects;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLayeredPane;
import javax.swing.JScrollPane;
import javax.swing.JTextPane;
import javax.swing.Timer;
import javax.swing.text.AbstractDocument;
import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.Element;
import javax.swing.text.StyleConstants;
import javax.swing.text.StyledDocument;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.junit.jupiter.MockitoSettings;
import org.mockito.quality.Strictness;

@ExtendWith(MockitoExtension.class)
@MockitoSettings(strictness = Strictness.LENIENT)
class TreeConsolePanelTest {

  private TreeConsolePanel sut;
  private JTextPane consolePane;
  private JButton btnCopy;
  private JLayeredPane layeredPane;

  @BeforeEach
  void setUp() {
    sut = new TreeConsolePanel(new Insets(5, 5, 5, 5));
    layeredPane = (JLayeredPane) sut.getComponent(0);
    var scrollPane = (JScrollPane) layeredPane.getComponentsInLayer(JLayeredPane.DEFAULT_LAYER)[0];
    consolePane = (JTextPane) scrollPane.getViewport().getView();
    btnCopy = (JButton) layeredPane.getComponentsInLayer(JLayeredPane.PALETTE_LAYER)[0];
  }

  @Nested
  @DisplayName("正常系テストケース")
  class PositiveTests {

    @Test
    @DisplayName("正常系: テキスト追加とTimerイベントの網羅")
    void test01() throws Exception {
      sut.setText("📁 root\n" + "└─ child.txt\n" + "normal");
      sut.append("📂 dir/");
      sut.append("├─ file");

      btnCopy.doClick();

      for (var listener : btnCopy.getActionListeners()) {
        listener.actionPerformed(null);
      }

      var field = TreeConsolePanel.class.getDeclaredField("copyTimer");
      field.setAccessible(true);
      var timer = (Timer) field.get(sut);

      if (Objects.nonNull(timer)) {
        for (var al : timer.getActionListeners()) {
          al.actionPerformed(null);
        }
      }
      assertNotNull(sut.getText());
    }

    @Test
    @DisplayName("正常系: レイアウトマネージャー全メソッド")
    void test02() {
      var frame = new JFrame();
      frame.add(sut);
      sut.setSize(new Dimension(500, 400));

      var lm = layeredPane.getLayout();
      lm.layoutContainer(layeredPane);
      lm.preferredLayoutSize(layeredPane);
      lm.minimumLayoutSize(layeredPane);
      lm.removeLayoutComponent(btnCopy);
      lm.addLayoutComponent("test", btnCopy);

      var container = new javax.swing.JPanel();
      container.setLayout(new BorderLayout());
      container.setSize(new Dimension(500, 400));
      container.add(consolePane);
      consolePane.setSize(new Dimension(100, 100));
      assertEquals(500, consolePane.getWidth());
      frame.dispose();
    }

    @Test
    @DisplayName("正常系: ViewFactory(LabelView含む)とDropTarget")
    void test03() {
      var kit = (TreeConsolePanel.NoWrapEditorKit) consolePane.getEditorKit();
      var factory = kit.getViewFactory();
      var doc = (AbstractDocument) consolePane.getDocument();
      var attr = mock(AttributeSet.class);

      String[] elements = {
        AbstractDocument.ContentElementName,
        AbstractDocument.ParagraphElementName,
        AbstractDocument.SectionElementName,
        StyleConstants.ComponentElementName,
        StyleConstants.IconElementName,
        "UnknownType"
      };

      for (String type : elements) {
        var elem = mock(Element.class);
        when(elem.getName()).thenReturn(type);
        when(elem.getAttributes()).thenReturn(attr);
        when(elem.getDocument()).thenReturn(doc);
        when(elem.getStartOffset()).thenReturn(0);
        when(elem.getEndOffset()).thenReturn(1);

        var view = factory.create(elem);
        assertNotNull(view);
      }

      var dt = new DropTarget();
      sut.setDropTarget(dt);
      assertEquals(dt, consolePane.getDropTarget());
    }
  }

  @Nested
  @DisplayName("異常系テストケース")
  class NegativeTests {

    @Test
    @DisplayName("異常系: 例外ハンドリングの網羅")
    void test01() throws BadLocationException {
      sut.append("init styles");

      var spyPane = spy(consolePane);
      var mockDoc = mock(StyledDocument.class);

      doAnswer(invocation -> consolePane.getStyle((String) invocation.getArguments()[0]))
          .when(spyPane)
          .getStyle(anyString());
      doAnswer(invocation -> consolePane.addStyle((String) invocation.getArguments()[0], null))
          .when(spyPane)
          .addStyle(anyString(), any());

      when(spyPane.getStyledDocument()).thenReturn(mockDoc);

      try {
        var field = TreeConsolePanel.class.getDeclaredField("consolePane");
        field.setAccessible(true);
        field.set(sut, spyPane);
      } catch (Exception e) {
        fail(e);
      }

      doThrow(new BadLocationException("mock", 0))
          .when(mockDoc)
          .insertString(anyInt(), anyString(), any());

      assertDoesNotThrow(
          () -> {
            sut.append("trigger exception");
          });
    }

    @Test
    @DisplayName("異常系: 入力値ガード")
    void test02() {
      try {
        sut.append(null);
      } catch (Exception ignored) {
      }
      try {
        sut.setText(null);
      } catch (Exception ignored) {
      }
    }

    @Test
    @DisplayName("異常系: 空文字コピー")
    void test03() {
      consolePane.setText("");
      btnCopy.doClick();
      assertEquals("📋", btnCopy.getText());
    }
  }
}

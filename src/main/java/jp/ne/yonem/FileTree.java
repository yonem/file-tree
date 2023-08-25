package jp.ne.yonem;

import javax.swing.*;

public class FileTree {

    private static final String SUCCESS_TITLE = "正常終了";
    private static final String SUCCESS_MESSAGE = "ファイルを出力しました";
    private static final String FAILURE_TITLE = "異常終了";
    private static final String FAILURE_MESSAGE = "ファイルの出力に失敗しました";

    public static void main(String[] args) {

        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            var chooser = new JFileChooser();
            chooser.setMultiSelectionEnabled(false);
            chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            var selected = chooser.showOpenDialog(null);

            if (selected != JFileChooser.APPROVE_OPTION) {
                return;
            }
            System.out.printf("選択ディレクトリ：%s%n", chooser.getSelectedFile().getAbsolutePath());

            JOptionPane.showMessageDialog(
                    null,
                    SUCCESS_MESSAGE,
                    SUCCESS_TITLE,
                    JOptionPane.INFORMATION_MESSAGE
            );

        } catch (Exception e) {
            JOptionPane.showMessageDialog(
                    null,
                    FAILURE_MESSAGE,
                    FAILURE_TITLE,
                    JOptionPane.ERROR_MESSAGE
            );
        }
    }
}

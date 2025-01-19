import Audit.MainFile.InterFace;
import com.formdev.flatlaf.fonts.roboto.FlatRobotoFont;
import com.formdev.flatlaf.themes.FlatMacDarkLaf;

import javax.swing.*;
import java.awt.*;

public class AutomationFile extends JFrame {

    public AutomationFile()
    {
        init();
    }

    private void init ()
    {
        setTitle("File Search and Copy");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(new Dimension(1200, 700));
        setLocationRelativeTo(null);
        setContentPane(new InterFace());
    }

    public static void main (String [] args)
    {
        FlatRobotoFont.install();

        FlatMacDarkLaf.registerCustomDefaultsSource("ResourcesFiles.themes");
        UIManager.put("defaultFont", new Font(FlatRobotoFont.FAMILY, Font.PLAIN, 13));
        FlatMacDarkLaf.setup();
        EventQueue.invokeLater(() -> new AutomationFile().setVisible(true));
    }
}

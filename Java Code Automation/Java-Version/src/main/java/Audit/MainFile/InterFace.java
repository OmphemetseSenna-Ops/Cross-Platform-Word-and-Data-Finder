package Audit.MainFile;

import com.formdev.flatlaf.FlatClientProperties;
import net.miginfocom.swing.MigLayout;
import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import javax.swing.*;
import java.awt.Font;
import java.awt.Color;
import javax.swing.border.Border;
import javax.swing.border.CompoundBorder;
import javax.swing.border.LineBorder;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.*;
import java.nio.file.Files;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class InterFace extends JPanel {

    private JTextField FolderField;
    private JTextField KeywordsField;
    private JComboBox<String> fileTypeComboBox;
    private JButton BrowseButton;
    private JButton SearchButton;
    private JButton HelpButton;
    private JButton ThemeButton;

    public InterFace() {
        init();
    }

    private void init() {
        setLayout(new MigLayout("fill, insets 20", "[center]", "[center]"));
        FolderField = new JTextField();
        KeywordsField = new JTextField();
        fileTypeComboBox = new JComboBox<>(new String[]{"Text Files", "Word Documents", "Excel Spreadsheets", "PDF Files"});
        SearchButton = new JButton("SEARCH FILES");
        BrowseButton = new JButton("Browse Folder");
        HelpButton = new JButton("Help");
        ThemeButton = new JButton("Theme");

        JPanel panel = new JPanel(new MigLayout("wrap, fillx, insets 35 45 30 45", "fill, 250:320"));

        panel.putClientProperty(FlatClientProperties.STYLE, "" +
                "arc: 20;" +
                "[Light]background: darken(@background, 3%);" +
                "[dark]background: lighten(@background, 3%)");

        JLabel lbTitle = new JLabel("Welcome!");
        lbTitle.putClientProperty(FlatClientProperties.STYLE, "" + "font:bold +10");

        JLabel description = new JLabel("Searching for a text in unknown file or looking to sort files?");
        description.putClientProperty(FlatClientProperties.STYLE, "" +
                "[light]foreground: lighten(@foreground, 30%);" +
                "[dark]foreground: darken(@foreground, 30%)");

        panel.add(lbTitle);
        panel.add(description);
        panel.add(new JLabel("Select Folder"), "gapy 8");
        panel.add(FolderField);
        panel.add(BrowseButton);
        panel.add(new JLabel("Enter Keywords"), "gapy 8");
        panel.add(KeywordsField);
        panel.add(new JLabel("Select File type"), "gapy 8");
        panel.add(fileTypeComboBox);
        panel.add(new JLabel(""));
        panel.add(new JLabel(" ------------------------------------------------------------------------------ "));
        panel.add(SearchButton, "height 40::, span");

        // Add the Help and Theme buttons to the top right corner
        JPanel topPanel = new JPanel(new MigLayout("insets 0", "[right]", ""));
        topPanel.add(HelpButton, "split 2, right");
        topPanel.add(ThemeButton, "right");
        panel.add(topPanel, "dock north");

        // Set initial border for SearchButton
        Border initialBorder = new CompoundBorder(
                new LineBorder(Color.decode("#5856D6"), 2),
                new LineBorder(Color.decode("#5856D6"), 1)
        );
        SearchButton.putClientProperty(FlatClientProperties.STYLE, "" +
                "arc: 15;" +
                "borderWidth: 2;" +
                "borderColor: #5856D6;");

        // Add mouse listener to change border on hover
        SearchButton.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseEntered(MouseEvent e) {
                Border hoverBorder = new CompoundBorder(
                        new LineBorder(Color.ORANGE, 2),
                        new LineBorder(Color.ORANGE, 1)
                );
                SearchButton.setBorder(hoverBorder);
            }

            @Override
            public void mouseExited(MouseEvent e) {
                SearchButton.setBorder(initialBorder);
            }
        });

        // Add action listener to BrowseButton
        BrowseButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                browseFolder();
            }
        });

        // Add action listener to SearchButton
        SearchButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                startSearch();
            }
        });

        // Add action listener to HelpButton
        HelpButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                showHelp();
            }
        });

        // Add action listener to ThemeButton
        ThemeButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                changeTheme();
            }
        });

        add(panel);
    }

    // Method to browse folder
    private void browseFolder() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            FolderField.setText(selectedFile.getAbsolutePath());
        }
    }

    // Method to start the search
    private void startSearch() {
        String folderPath = FolderField.getText();
        String keyword = KeywordsField.getText();
        String fileType = (String) fileTypeComboBox.getSelectedItem();

        if (folderPath.isEmpty() || keyword.isEmpty() || fileType.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Please fill in all fields.", "Input Error", JOptionPane.WARNING_MESSAGE);
            return;
        }

        int numFilesFound = searchAndCopyFiles(folderPath, keyword, fileType);

        if (numFilesFound > 0) {
            JOptionPane.showMessageDialog(this, "Number of files containing '" + keyword + "': " + numFilesFound, "Search Complete", JOptionPane.INFORMATION_MESSAGE);
        } else {
            JOptionPane.showMessageDialog(this, "No files containing '" + keyword + "' found.", "Search Complete", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    // Method to search and copy files
    private int searchAndCopyFiles(String sourceFolder, String keyword, String fileType) {
        File folder = new File(sourceFolder);
        File[] files = folder.listFiles((dir, name) -> {
            switch (fileType) {
                case "Text Files":
                    return name.endsWith(".txt");
                case "Word Documents":
                    return name.endsWith(".docx");
                case "Excel Spreadsheets":
                    return name.endsWith(".xlsx");
                case "PDF Files":
                    return name.endsWith(".pdf");
                default:
                    return false;
            }
        });

        if (files == null) {
            return 0;
        }

        Pattern pattern = Pattern.compile("\\b" + Pattern.quote(keyword) + "\\b");
        int count = 0;
        File destinationFolder = new File(sourceFolder, keyword);

        for (File file : files) {
            try {
                String content = "";
                if (fileType.equals("Text Files")) {
                    content = new String(Files.readAllBytes(file.toPath()));
                } else if (fileType.equals("Word Documents")) {
                    content = readWordDocument(file);
                } else if (fileType.equals("Excel Spreadsheets")) {
                    content = readExcelFile(file);
                }

                Matcher matcher = pattern.matcher(content);
                if (matcher.find()) {
                    if (!destinationFolder.exists()) {
                        destinationFolder.mkdir();
                    }
                    FileUtils.copyFileToDirectory(file, destinationFolder);
                    count++;
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return count;
    }

    // Method to read content from Word document
    private String readWordDocument(File file) throws IOException {
        StringBuilder content = new StringBuilder();
        try (FileInputStream fis = new FileInputStream(file);
             XWPFDocument document = new XWPFDocument(fis)) {
            for (XWPFParagraph para : document.getParagraphs()) {
                content.append(para.getText()).append(" ");
            }
        }
        return content.toString();
    }

    // Method to read content from Excel file
    private String readExcelFile(File file) throws IOException {
        StringBuilder content = new StringBuilder();
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        content.append(cell.toString()).append(" ");
                    }
                    content.append("\n");
                }
            }
        }
        return content.toString();
    }

    // Method to show help information
    private void showHelp() {
        JOptionPane.showMessageDialog(this, "How to Use the File Search and Copy Tool:\n\n" +
                "1. Select the folder containing the files by clicking 'Browse'.\n" +
                "2. Enter the keyword to search for in the files.\n" +
                "3. Choose the type of files to search from the dropdown menu.\n" +
                "4. Click 'SEARCH FILES' to start the search.\n" +
                "5. The tool will search for the whole word in the files and copy matching files to a new folder named after the keyword.\n\n" +
                "Note: The tool currently supports Text Files, Word Documents, Excel Spreadsheets, and PDF Files.", "Help", JOptionPane.INFORMATION_MESSAGE);
    }

    // Method to change theme
    private void changeTheme() {
        // Implementation to change theme
        JOptionPane.showMessageDialog(this, "Theme change functionality goes here.", "Change Theme", JOptionPane.INFORMATION_MESSAGE);
    }

    public static void main(String[] args) {
        JFrame frame = new JFrame();
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setTitle("File Searching Tool");
        frame.add(new InterFace());
        frame.pack();
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }
}

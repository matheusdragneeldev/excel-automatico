package org.example;

import com.formdev.flatlaf.FlatLightLaf;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import javax.swing.*;
import javax.swing.text.*;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.List;
import java.util.concurrent.ExecutionException;

public class FormularioGUI extends JFrame {

    private final JTextField linhaTextField;
    private final JButton preencherButton;
    private final JTextPane logPane;
    private final JProgressBar progressBar;

    private WebDriver driver;
    private final DataFormatter formatter = new DataFormatter();
    private volatile boolean inicializacaoConcluida = false;

    public FormularioGUI() {
        setTitle("Excel Automático");
        setSize(550, 420);
        setLocationRelativeTo(null);
        setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);

        JPanel topPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 10));
        topPanel.add(new JLabel("Linha do Excel:"));
        linhaTextField = new JTextField(5);
        topPanel.add(linhaTextField);

        preencherButton = new JButton("Preencher");
        preencherButton.setEnabled(false);
        topPanel.add(preencherButton);

        logPane = new JTextPane();
        logPane.setEditable(false);
        JScrollPane scrollPane = new JScrollPane(logPane);

        progressBar = new JProgressBar();
        progressBar.setIndeterminate(true);
        progressBar.setString("Inicializando...");
        progressBar.setStringPainted(true);

        setLayout(new BorderLayout(10, 10));
        add(topPanel, BorderLayout.NORTH);
        add(scrollPane, BorderLayout.CENTER);
        add(progressBar, BorderLayout.SOUTH);
        ((JPanel) getContentPane()).setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        preencherButton.addActionListener((e) -> executarPreenchimento());

        addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowClosing(java.awt.event.WindowEvent windowEvent) {
                if (!inicializacaoConcluida) {
                    JOptionPane.showMessageDialog(FormularioGUI.this,
                            "Aguarde a inicialização terminar.",
                            "Inicializando", JOptionPane.INFORMATION_MESSAGE);
                    return;
                }
                preencherButton.setEnabled(false);
                linhaTextField.setEnabled(false);
                appendLog("Encerrando recursos...\n", Color.GRAY);
                new SwingWorker<Void, Void>() {
                    @Override
                    protected Void doInBackground() {
                        if (driver != null) driver.quit();
                        return null;
                    }
                    @Override
                    protected void done() {
                        System.exit(0);
                    }
                }.execute();
            }
        });

        inicializarAutomacao();
    }

    private void inicializarAutomacao() {
        appendLog("Inicializando automação...\n", Color.BLUE);
        new SwingWorker<Void, String>() {
            @Override
            protected Void doInBackground() throws Exception {
                publish("Iniciando WebDriver...");
                WebDriverManager.chromedriver().setup();
                ChromeOptions options = new ChromeOptions();
                options.addArguments("--no-sandbox", "--disable-extensions", "--disable-popup-blocking", "--remote-allow-origins=*");
                driver = new ChromeDriver(options);
                driver.manage().window().maximize();
                driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
                publish("WebDriver pronto.");

                publish("Procurando arquivo Excel...");
                String caminhoExcel = PreenchedorDeFormulario.encontrarPrimeiroArquivoXlsx();
                if (caminhoExcel == null) {
                    throw new IOException("Nenhum arquivo .xlsx encontrado na pasta.");
                }
                publish("Arquivo encontrado: " + new File(caminhoExcel).getName());
                return null;
            }
            @Override
            protected void process(List<String> chunks) {
                for (String msg : chunks) appendLog(msg + "\n", Color.DARK_GRAY);
            }
            @Override
            protected void done() {
                try {
                    get();
                    appendLog("\nPronto! Insira o número da linha e clique em Preencher.\n", new Color(0, 128, 0));
                    preencherButton.setEnabled(true);
                    progressBar.setVisible(false);
                } catch (Exception e) {
                    Throwable cause = e.getCause() != null ? e.getCause() : e;
                    appendLog("\nERRO: " + cause.getMessage() + "\n", Color.RED);
                    JOptionPane.showMessageDialog(FormularioGUI.this, cause.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
                } finally {
                    inicializacaoConcluida = true;
                }
            }
        }.execute();
    }

    private void executarPreenchimento() {
        String textoLinha = linhaTextField.getText().trim();
        int linha;
        try {
            linha = Integer.parseInt(textoLinha);
            if (linha <= 0) throw new NumberFormatException();
        } catch (NumberFormatException ex) {
            JOptionPane.showMessageDialog(this, "Número inválido!", "Erro", JOptionPane.ERROR_MESSAGE);
            return;
        }
        preencherButton.setEnabled(false);
        appendLog("Iniciando preenchimento (linha " + linha + ")...\n", Color.BLUE);

        new SwingWorker<String, Void>() {
            @Override
            protected String doInBackground() {
                String caminhoExcel = PreenchedorDeFormulario.encontrarPrimeiroArquivoXlsx();
                if (caminhoExcel == null) return "Arquivo Excel não encontrado.";
                try (FileInputStream fis = new FileInputStream(caminhoExcel);
                     XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
                    XSSFSheet sheet = workbook.getSheetAt(0);
                    Row row = sheet.getRow(linha - 1);
                    if (row == null) return "Linha vazia ou inexistente.";
                    PreenchedorDeFormulario.preencherFormulario(driver, row, formatter);
                    return "Formulário da linha " + linha + " preenchido!";
                } catch (IOException e) {
                    return "Erro ao ler Excel: " + e.getMessage();
                } catch (Exception e) {
                    return "Erro: " + e.getMessage();
                }
            }
            @Override
            protected void done() {
                try {
                    appendLog(get() + "\n", new Color(0, 128, 0));
                } catch (InterruptedException | ExecutionException e) {
                    appendLog("Erro inesperado: " + e.getMessage() + "\n", Color.RED);
                } finally {
                    preencherButton.setEnabled(true);
                }
            }
        }.execute();
    }

    private void appendLog(String msg, Color color) {
        StyledDocument doc = logPane.getStyledDocument();
        Style style = logPane.addStyle("Style", null);
        StyleConstants.setForeground(style, color);
        try {
            doc.insertString(doc.getLength(), msg, style);
        } catch (BadLocationException ignored) {}
    }

    public static void main(String[] args) {
        FlatLightLaf.setup(); // Tema moderno
        SwingUtilities.invokeLater(() -> new FormularioGUI().setVisible(true));
    }
}

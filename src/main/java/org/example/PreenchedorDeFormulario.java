package org.example;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.IOException;
import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.Duration;
import java.util.stream.Stream;

public class PreenchedorDeFormulario {

    private static final String FORM_URL = "https://form.jotform.com/252195590102048";

    // Preenche os campos de um formulário web com dados de uma linha da planilha
    public static void preencherFormulario(WebDriver driver, Row row, DataFormatter formatter) {
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));
        try {
            driver.get(FORM_URL);

            // Para cada campo, tenta preencher, capturando erros para não travar
            preencherCampoSeguro(wait, By.name("q8_insiraUma"), formatter.formatCellValue(row.getCell(1)));
            preencherCampoSeguro(wait, By.name("q9_insiraUma9"), formatter.formatCellValue(row.getCell(2)));
            preencherCampoSeguro(wait, By.name("q10_insiraUma10"), formatter.formatCellValue(row.getCell(3)));
            preencherCampoSeguro(wait, By.name("q11_insiraUma11"), formatter.formatCellValue(row.getCell(4)));
            preencherCampoSeguro(wait, By.name("q12_insiraUma12"), formatter.formatCellValue(row.getCell(5)));
            preencherCampoSeguro(wait, By.name("q16_insiraUma16"), formatter.formatCellValue(row.getCell(6)));
            preencherCampoSeguro(wait, By.name("q14_insiraUma14"), formatter.formatCellValue(row.getCell(7)));
            preencherCampoSeguro(wait, By.name("q15_insiraUma15"), formatter.formatCellValue(row.getCell(8)));

        } catch (Exception e) {
            System.err.println("Erro ao acessar a página ou preencher o formulário: " + e.getMessage());
        }
    }

    // Preenche um único campo na página de forma segura.
    private static void preencherCampoSeguro(WebDriverWait wait, By locator, String texto) {
        if (texto != null && !texto.trim().isEmpty()) {
            try {
                WebElement campo = wait.until(ExpectedConditions.visibilityOfElementLocated(locator));
                campo.clear();
                campo.sendKeys(texto);
            } catch (Exception e) {
                System.err.println("Erro ao preencher campo " + locator + ": " + e.getMessage());
            }
        }
    }

    // Encontra o caminho do primeiro arquivo .xlsx no mesmo diretório do JAR/execução.
    public static String encontrarPrimeiroArquivoXlsx() {
        try {
            Path diretorioJar = Paths.get(PreenchedorDeFormulario.class.getProtectionDomain()
                    .getCodeSource().getLocation().toURI()).getParent();
            try (Stream<Path> arquivos = Files.list(diretorioJar)) {
                return arquivos
                        .filter(Files::isRegularFile)
                        .filter(path -> path.getFileName().toString().toLowerCase().endsWith(".xlsx"))
                        .findFirst()
                        .map(Path::toString)
                        .orElse(null);
            }
        } catch (URISyntaxException | IOException e) {
            e.printStackTrace();
            return null;
        }
    }
}
package org.contraApp;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.*;
import java.nio.file.Paths;
import java.text.Normalizer;
import java.time.Duration;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.*;

class RelatorioEnvio {
    String nome, telefone, status, detalheErro;
    RelatorioEnvio(String nome, String telefone, String status, String detalheErro) {
        this.nome = nome;
        this.telefone = telefone;
        this.status = status;
        this.detalheErro = detalheErro;
    }
}

public class Main {

    private static final String PASTA_CONTRACHEQUES = "contracheques";

    public static void main(String[] args) throws InterruptedException {
        WebDriverManager.chromedriver().setup();
        WebDriver driver = iniciarChromeComPerfil();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
        driver.get("https://web.whatsapp.com");

        // Espera automática por login, QR ou já logado
        System.out.println("🚀 Aguardando autenticação no WhatsApp Web...");
        boolean autenticado = false;
        long deadline = System.currentTimeMillis() + 180000; // até 3 minutos
        while (System.currentTimeMillis() < deadline && !autenticado) {
            try {
                // Vários seletores: campo de conversa, sidebar, campo de pesquisa
                if (
                        driver.findElements(By.xpath("//div[@contenteditable='true' and @role='textbox']")).size() > 0 ||
                                driver.findElements(By.xpath("//div[contains(@class,'_ak1l')]")).size() > 0 ||
                                driver.findElements(By.xpath("//div[@tabindex='-1' and @role='region']")).size() > 0 ||
                                driver.findElements(By.xpath("//div[@title='Caixa de texto de pesquisa']")).size() > 0 ||
                                driver.findElements(By.xpath("//div[@title='Search input textbox']")).size() > 0
                ) {
                    autenticado = true;
                } else {
                    Thread.sleep(1200);
                }
            } catch (Exception e) {
                Thread.sleep(1200);
            }
        }

        if (!autenticado) {
            System.err.println("❌ WhatsApp Web não autenticado a tempo. Encerrando programa.");
            driver.quit();
            return;
        }
        System.out.println("✅ WhatsApp Web autenticado! Iniciando os envios...");

        enviarContracheques(driver, wait);
        driver.quit();
    }

    private static WebDriver iniciarChromeComPerfil() {
        String perfilPath = Paths.get(System.getProperty("user.home"), "chrome-profile-envio").toString();
        new File(perfilPath).mkdirs();
        ChromeOptions options = new ChromeOptions();
        options.addArguments("user-data-dir=" + perfilPath);
        return new ChromeDriver(options);
    }

    private static void enviarContracheques(WebDriver driver, WebDriverWait wait) {
        LocalDate data = LocalDate.now().minusMonths(1);
        String nomeMes = data.getMonth().getDisplayName(TextStyle.FULL, new Locale("pt", "BR"));
        int ano = data.getYear();
        List<RelatorioEnvio> relatorio = new ArrayList<>();

        try (InputStream is = new FileInputStream("funcionarios.xlsx");
             Workbook workbook = new XSSFWorkbook(is)) {
            Iterator<Row> rows = workbook.getSheetAt(0).iterator();
            if (rows.hasNext()) rows.next(); // cabeçalho

            while (rows.hasNext()) {
                Row row = rows.next();
                String nome = getCellValue(row.getCell(0));
                String telefone = getCellValue(row.getCell(1)).replaceAll("[^0-9]", "");

                if (nome.isBlank() || telefone.length() < 10) {
                    relatorio.add(new RelatorioEnvio(nome, telefone, "ERRO", "Nome vazio ou telefone inválido"));
                    continue;
                }
                String nomeArquivo = normalizar(nome) + ".pdf";
                File arquivo = Paths.get(PASTA_CONTRACHEQUES, nomeArquivo).toFile();
                if (!arquivo.exists()) {
                    relatorio.add(new RelatorioEnvio(nome, telefone, "ERRO", "Arquivo não encontrado"));
                    continue;
                }
                String resultado = enviarParaContato(driver, wait, nome, telefone, arquivo, nomeMes, ano);
                relatorio.add(new RelatorioEnvio(nome, telefone, resultado.startsWith("OK") ? "ENVIADO" : "ERRO", resultado));
            }
        } catch (Exception e) {
            System.err.println("Falha ao processar planilha: " + e.getMessage());
        }
        salvarRelatorio(relatorio);
    }

    private static String getDesktopPath() {
        String home = System.getProperty("user.home");
        String[] caminhos = {
                home + File.separator + "Desktop",
                home + File.separator + "Área de Trabalho",
                home + File.separator + "OneDrive" + File.separator + "Desktop",
                home + File.separator + "OneDrive" + File.separator + "Área de Trabalho"
        };
        for (String caminho : caminhos) {
            File dir = new File(caminho);
            if (dir.exists() && dir.isDirectory()) return dir.getAbsolutePath();
        }
        return null;
    }

    private static void salvarRelatorio(List<RelatorioEnvio> relatorio) {
        String desktopPath = getDesktopPath();
        String nomeArquivo = (desktopPath != null ? desktopPath : "") + (desktopPath != null ? File.separator : "") + "relatorio_envio.csv";
        try (PrintWriter writer = new PrintWriter(new File(nomeArquivo))) {
            writer.println("Nome,Telefone,Status");
            for (RelatorioEnvio item : relatorio)
                writer.printf("%s,%s,%s%n",
                        item.nome == null ? "" : item.nome.replace(",", ";"),
                        item.telefone == null ? "" : item.telefone,
                        item.status == null ? "" : item.status);
        } catch (Exception e) {
            try (PrintWriter writer = new PrintWriter(new File("relatorio_envio.csv"))) {
                writer.println("Nome,Telefone,Status");
                for (RelatorioEnvio item : relatorio)
                    writer.printf("%s,%s,%s%n",
                            item.nome == null ? "" : item.nome.replace(",", ";"),
                            item.telefone == null ? "" : item.telefone,
                            item.status == null ? "" : item.status);
            } catch (Exception ex) {
                System.err.println("Erro ao salvar relatório: " + ex.getMessage());
            }
        }
    }

    private static String enviarParaContato(WebDriver driver, WebDriverWait wait, String nome, String telefone, File arquivoPDF, String mes, int ano) throws InterruptedException {
        driver.get("https://web.whatsapp.com/send?phone=55" + telefone);
        try {
            wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(@class,'copyable-text')]")));
            Thread.sleep(2000);

            boolean anexoEnviado = false;
            try {
                By botaoAnexar = By.xpath("//span[@data-icon='plus-rounded' or @data-icon='clip']");
                wait.until(ExpectedConditions.elementToBeClickable(botaoAnexar)).click();
                Thread.sleep(1000);
                WebElement input = driver.findElement(By.xpath("//input[@type='file']"));
                input.sendKeys(arquivoPDF.getAbsolutePath());
                Thread.sleep(3000);
                WebElement botaoEnviar = wait.until(ExpectedConditions.elementToBeClickable(
                        By.xpath("//div[@role='button' and @aria-label='Enviar']")));
                botaoEnviar.click();
                anexoEnviado = true;
                wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@contenteditable='true' and @aria-label='Digite uma mensagem']")));
                Thread.sleep(1000);
            } catch (Exception ignore) {}

            String mensagem = "Olá " + nome + ", este é o seu contracheque de " + mes + " de " + ano + ". Qualquer dúvida, fale com o RH.";
            boolean mensagemEnviada = false; int tentativas = 0;
            while (!mensagemEnviada && tentativas < 5) {
                try {
                    WebElement caixaTexto = null;
                    try { caixaTexto = wait.until(ExpectedConditions.elementToBeClickable(
                            By.xpath("//div[@contenteditable='true' and @aria-label='Digite uma mensagem']"))); }
                    catch (Exception e1) {
                        try { caixaTexto = wait.until(ExpectedConditions.elementToBeClickable(
                                By.xpath("//div[@contenteditable='true' and @title='Mensagem']"))); }
                        catch (Exception e2) { caixaTexto = wait.until(ExpectedConditions.elementToBeClickable(
                                By.xpath("//div[@contenteditable='true']"))); }
                    }
                    caixaTexto.click();
                    Thread.sleep(500);
                    caixaTexto.sendKeys(Keys.chord(Keys.CONTROL, "a"), Keys.BACK_SPACE);
                    Thread.sleep(300);
                    caixaTexto.sendKeys(mensagem);
                    Thread.sleep(800);
                    caixaTexto.sendKeys(Keys.ENTER); Thread.sleep(200);
                    new Actions(driver).moveToElement(caixaTexto).click().sendKeys(Keys.ENTER).perform(); Thread.sleep(200);
                    caixaTexto.click(); Thread.sleep(100); caixaTexto.sendKeys(Keys.ENTER);
                    mensagemEnviada = true;
                } catch (Exception eMsg) { tentativas++; Thread.sleep(2000); }
            }
            if (mensagemEnviada) return anexoEnviado ? "OK - Mensagem e anexo enviados" : "OK - Apenas mensagem enviada";
            else return "Falha ao enviar mensagem";
        } catch (Exception e) {
            return "Erro ao enviar: " + e.getMessage();
        }
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        return cell.getCellType() == CellType.STRING ? cell.getStringCellValue().trim()
                : String.valueOf((long) cell.getNumericCellValue()).trim();
    }
    private static String normalizar(String texto) {
        return Normalizer.normalize(texto, Normalizer.Form.NFD)
                .replaceAll("\\p{InCombiningDiacriticalMarks}+", "")
                .toLowerCase()
                .trim()
                .replaceAll("[^a-z0-9]", " ");
    }
}
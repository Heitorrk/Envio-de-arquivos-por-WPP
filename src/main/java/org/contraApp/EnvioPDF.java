package org.contraApp;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.*;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.text.Normalizer;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.*;
import java.util.NoSuchElementException;

class RelatorioEnvio {
    String nome, telefone, status, detalhe;
    RelatorioEnvio(String nome, String telefone, String status, String detalhe) {
        this.nome = nome; this.telefone = telefone; this.status = status; this.detalhe = detalhe;
    }
}

public class EnvioPDF {

    private static final String PASTA_CONTRACHEQUES = "contracheques";
    private static final String PLANILHA = "funcionarios.xlsx";
    private static final String ARQUIVO_MENSAGEM = "mensagem.txt"; // {nome},{mes},{ano}
    private static final DataFormatter FORMATTER = new DataFormatter(new Locale("pt","BR"));

    public static void main(String[] args) throws Exception {
        // Pré-etapa: padroniza nomes de PDFs (se você usa o RenomeadorPDF)
        try {
            System.out.println("🔧 Pré-etapa: renomear PDFs (planilha='" + PLANILHA + "', pasta='" + PASTA_CONTRACHEQUES + "')");
            RenomeadorPDF.renomearContracheques();
        } catch (Throwable t) {
            System.out.println("ℹ️ Renomeador não executado (classe ausente ou opcional). Seguindo.");
        }

        WebDriver driver = iniciarChrome();
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(60));

        driver.get("https://web.whatsapp.com");
        System.out.println("🚀 Aguardando autenticação no WhatsApp Web...");
        if (!aguardarAutenticacao(driver)) {
            System.err.println("❌ Não autenticou a tempo. Encerrando.");
            driver.quit();
            return;
        }
        System.out.println("✅ Autenticado! Iniciando…");

        enviarContracheques(driver, wait);
        driver.quit();
    }

    private static WebDriver iniciarChrome() {
        Path base = getAppDataDir();
        Path profile = base.resolve("chrome-profile-envio");
        profile.toFile().mkdirs();

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--user-data-dir=" + profile.toAbsolutePath());
        options.addArguments("--no-default-browser-check", "--disable-notifications",
                "--disable-infobars", "--lang=pt-BR");
        options.addArguments("--disable-dev-shm-usage", "--no-sandbox");
        options.addArguments("--disable-gpu");
        options.addArguments("--remote-debugging-port=0");
        options.addArguments("--disable-extensions");
        // Para agendamento em servidor sem UI, você pode ativar o headless:
        // options.addArguments("--headless=new");
        return new ChromeDriver(options);
    }

    private static Path getAppDataDir() {
        String os = System.getProperty("os.name").toLowerCase();
        String home = System.getProperty("user.home");
        if (os.contains("win")) return Paths.get(System.getenv("LOCALAPPDATA"), "Evox", "ContraZap");
        if (os.contains("mac")) return Paths.get(home, "Library", "Application Support", "Evox", "ContraZap");
        return Paths.get(home, ".config", "evox", "contrazap");
    }

    private static boolean aguardarAutenticacao(WebDriver driver) throws InterruptedException {
        long deadline = System.currentTimeMillis() + 240000; // 4 min
        while (System.currentTimeMillis() < deadline) {
            try {
                boolean ok = !driver.findElements(By.cssSelector("div[contenteditable='true'][data-tab]")).isEmpty()
                        || !driver.findElements(By.cssSelector("footer")).isEmpty();
                if (ok) return true;
                Thread.sleep(400);
            } catch (Exception ignored) { Thread.sleep(400); }
        }
        return false;
    }

    private static void enviarContracheques(WebDriver driver, WebDriverWait wait) {
        // MÊS DE COMPETÊNCIA usado na mensagem {mes}/{ano}
        LocalDate competencia = LocalDate.now().minusMonths(1);
        String mesNome = competencia.getMonth().getDisplayName(TextStyle.FULL, new Locale("pt","BR"));
        int ano = competencia.getYear();
        String mensagemPadrao = lerMensagemPadrao();

        List<RelatorioEnvio> relatorio = new ArrayList<>();

        try (InputStream is = new FileInputStream(PLANILHA);
             Workbook wb = new XSSFWorkbook(is)) {

            Iterator<Row> rows = wb.getSheetAt(0).iterator();
            if (rows.hasNext()) rows.next(); // cabeçalho
            int numLinha = 2;

            while (rows.hasNext()) {
                Row r = rows.next();
                String nome = getCell(r.getCell(0));                 // A -> nome
                String telefone = getCell(r.getCell(1)).replaceAll("[^0-9]",""); // B
                String nomeArquivoC = getCell(r.getCell(2));         // C -> nome do arquivo (base)

                if (!linhaValida(nome, telefone)) {
                    String det = "Linha inválida: nome='" + nome + "' tel='" + telefone + "'";
                    System.err.println("↳ linha " + numLinha + " ignorada: " + det);
                    relatorio.add(new RelatorioEnvio(nome, telefone, "IGNORADO", det));
                    numLinha++; continue;
                }

                String base = nomeArquivoC.isBlank() ? nome : nomeArquivoC;
                base = sanitizeBaseNomeArquivo(base);

                // 1) tentativa direta
                File arquivoPDF = Paths.get(PASTA_CONTRACHEQUES, base + ".pdf").toFile();

                // 2) fallback: localizar por aproximação
                if (!arquivoPDF.exists()) {
                    File aprox = localizarPdfAproximado(base);
                    if (aprox != null) {
                        System.out.println("🔎 Usando arquivo encontrado por aproximação: " + aprox.getName());
                        arquivoPDF = aprox;
                    } else {
                        String det = "Arquivo não encontrado: " + Paths.get(PASTA_CONTRACHEQUES, base + ".pdf");
                        System.err.println("ERRO linha " + numLinha + ": " + det);
                        relatorio.add(new RelatorioEnvio(nome, telefone, "ERRO", det));
                        numLinha++; continue;
                    }
                }

                String resultado = enviarWhats(driver, wait, nome, telefone, arquivoPDF, mesNome, ano, mensagemPadrao);
                if ("OK".equals(resultado)) {
                    relatorio.add(new RelatorioEnvio(nome, telefone, "ENVIADO", "OK - anexo + mensagem"));
                } else {
                    System.err.println("ERRO linha " + numLinha + ": " + resultado);
                    relatorio.add(new RelatorioEnvio(nome, telefone, "ERRO", resultado));
                }
                numLinha++;
                try { Thread.sleep(400); } catch (InterruptedException ignored) {}
            }
        } catch (FileNotFoundException e) {
            System.err.println("❌ Planilha '" + PLANILHA + "' não encontrada.");
        } catch (Exception e) {
            System.err.println("Falha ao processar planilha: " + e.getMessage());
        }

        // Salva relatório e ARQUIVA PDFs, esvaziando a pasta de trabalho
        salvarRelatorioEArquivar(relatorio);
    }

    private static String enviarWhats(WebDriver driver, WebDriverWait wait, String nome, String telefone, File pdf,
                                      String mes, int ano, String mensagemPadrao) {
        try {
            System.out.println("📱 Enviando para: " + nome + " (" + telefone + ")");
            String url = "https://web.whatsapp.com/send?phone=55" + telefone;
            driver.get(url);

            System.out.println("⏳ Aguardando chat carregar...");
            wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("footer")));
            Thread.sleep(700);

            List<WebElement> invalidNumber = driver.findElements(By.xpath("//*[contains(text(), 'Número de telefone') or contains(text(), 'Phone number')]"));
            if (!invalidNumber.isEmpty()) {
                return "Número de telefone inválido ou não cadastrado no WhatsApp";
            }

            // ============ Anexar documento ============
            System.out.println("📎 Anexando arquivo...");
            WebElement btnAnexar = null;
            List<String> seletoresAnexar = Arrays.asList(
                    "div[title='Anexar']", "span[data-icon='plus']", "span[data-icon='attach-menu-plus']",
                    "span[data-icon='clip']", "div[aria-label='Anexar']", "button[aria-label='Anexar']",
                    "*[data-testid='clip']", "div[role='button'] span[data-icon='plus']",
                    "div[role='button'] span[data-icon='clip']",
                    "div[title='Attach']", "span[data-icon='attach']", "span[data-icon='paperclip']",
                    "div[aria-label='Attach']", "*[data-testid='attach']",
                    "*[title*='nexar']", "*[title*='ttach']", "*[aria-label*='nexar']",
                    "*[aria-label*='ttach']", "div[role='button'][title*='clip']",
                    "button[title*='clip']", "span[class*='clip']", "div[class*='attach']"
            );
            for (String s : seletoresAnexar) {
                List<WebElement> els = driver.findElements(By.cssSelector(s));
                for (WebElement el : els) { if (el.isDisplayed() && el.isEnabled()) { btnAnexar = el; break; } }
                if (btnAnexar != null) break;
            }
            if (btnAnexar == null) return "Botão de anexar não encontrado. Verifique se o chat carregou.";
            btnAnexar.click();
            Thread.sleep(700);

            WebElement btnDocumento = null;
            List<String> seletoresDocumento = Arrays.asList(
                    "li[role='button'] span[data-icon='document']",
                    "li span[data-icon='document-filled']",
                    "li span[data-icon='document']",
                    "*[title*='Documento']", "*[aria-label*='Documento']",
                    "div[role='button'] span[data-icon='document']",
                    "*[data-testid='attach-document']",
                    "li span[data-icon='doc']", "li span[data-icon='pdf']",
                    "li[title*='Document']", "li[aria-label*='Document']"
            );
            for (String s : seletoresDocumento) {
                List<WebElement> els = driver.findElements(By.cssSelector(s));
                for (WebElement el : els) { if (el.isDisplayed() && el.isEnabled()) { btnDocumento = el; break; } }
                if (btnDocumento != null) break;
            }
            if (btnDocumento != null) { btnDocumento.click(); Thread.sleep(400); }

            System.out.println("📁 Selecionando arquivo...");
            WebElement inputFile = null;
            Thread.sleep(400);

            List<String> seletoresInput = Arrays.asList(
                    "input[type='file']",
                    "input[type='file']:not([accept*='image']):not([accept*='video'])",
                    "input[type='file'][accept*='*/*']",
                    "*[data-testid='media-attach-input']",
                    "form input[type='file']"
            );
            List<WebElement> todosInputs = new ArrayList<>();
            for (String s : seletoresInput) {
                todosInputs.addAll(driver.findElements(By.cssSelector(s)));
            }
            LinkedHashSet<WebElement> set = new LinkedHashSet<>(todosInputs);
            List<WebElement> inputs = new ArrayList<>(set);

            for (WebElement in : inputs) {
                String accept = in.getAttribute("accept");
                if (accept != null && (accept.contains("*/*") || accept.contains("application/"))) { inputFile = in; break; }
                if (accept == null || (!accept.toLowerCase().contains("image") && !accept.toLowerCase().contains("video"))) { inputFile = in; break; }
            }
            if (inputFile == null && !inputs.isEmpty()) inputFile = inputs.get(inputs.size() - 1);
            if (inputFile == null) inputFile = wait.until(ExpectedConditions.presenceOfElementLocated(By.cssSelector("input[type='file']")));

            System.out.println("📤 Enviando arquivo: " + pdf.getName());
            inputFile.sendKeys(pdf.getAbsolutePath());
            Thread.sleep(400);

            System.out.println("⏳ Aguardando preview do documento...");
            Thread.sleep(400);

            WebElement btnEnviar = null;
            List<String> seletoresEnviar = Arrays.asList(
                    "span[data-icon='send']", "*[data-testid='send']", "*[aria-label*='Enviar']",
                    "*[title*='Enviar']", "button span[data-icon='send']",
                    "div[role='button'] span[data-icon='send']",
                    "*[data-testid='compose-btn-send']",
                    "span[data-icon='send-light']", "button[aria-label*='Send']",
                    "button[title*='Send']", "button[data-testid='send-button']"
            );
            for (int tentativa = 0; tentativa < 15 && btnEnviar == null; tentativa++) {
                for (String s : seletoresEnviar) {
                    for (WebElement el : driver.findElements(By.cssSelector(s))) {
                        if (el.isDisplayed() && el.isEnabled()) { btnEnviar = el; break; }
                    }
                    if (btnEnviar != null) break;
                }
                if (btnEnviar == null) { Thread.sleep(400); }
            }
            if (btnEnviar != null) {
                try { btnEnviar.click(); }
                catch (Exception e) { ((JavascriptExecutor) driver).executeScript("arguments[0].click();", btnEnviar); }
                Thread.sleep(400);
            } else {
                new Actions(driver).sendKeys(Keys.ENTER).perform();
                Thread.sleep(400);
            }

            // ============ Mensagem ============
            System.out.println("💬 Enviando mensagem...");
            String mensagem = mensagemPadrao.replace("{nome}", nome).replace("{mes}", mes).replace("{ano}", String.valueOf(ano));

            WebElement caixaTexto = null;
            List<String> seletoresCaixa = Arrays.asList(
                    "div[contenteditable='true'][data-tab='10']",
                    "div[contenteditable='true'][data-tab]",
                    "div[contenteditable='true'][role='textbox']",
                    "*[data-testid='conversation-compose-box-input']",
                    "div[contenteditable='true']"
            );
            for (String s : seletoresCaixa) {
                for (WebElement el : driver.findElements(By.cssSelector(s))) {
                    if (el.isDisplayed() && el.isEnabled()) { caixaTexto = el; break; }
                }
                if (caixaTexto != null) break;
            }
            if (caixaTexto == null) caixaTexto = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("div[contenteditable='true']")));

            caixaTexto.click();
            Thread.sleep(400);
            caixaTexto.sendKeys(Keys.chord(Keys.CONTROL, "a"));
            caixaTexto.sendKeys(Keys.BACK_SPACE);
            Thread.sleep(400);

            String[] linhas = mensagem.split("\n");
            for (int i = 0; i < linhas.length; i++) {
                caixaTexto.sendKeys(linhas[i]);
                if (i < linhas.length - 1) caixaTexto.sendKeys(Keys.SHIFT, Keys.ENTER);
            }
            Thread.sleep(400);
            caixaTexto.sendKeys(Keys.ENTER);

            System.out.println("✅ Enviado com sucesso para " + nome);
            Thread.sleep(400);
            return "OK";

        } catch (TimeoutException te) {
            System.err.println("⏰ Timeout para " + nome + ": " + te.getMessage());
            return "Timeout aguardando elementos do WhatsApp";
        } catch (NoSuchElementException nse) {
            System.err.println("🔍 Elemento não encontrado para " + nome + ": " + nse.getMessage());
            return "Elemento não encontrado: " + nse.getMessage();
        } catch (Exception e) {
            System.err.println("❌ Erro para " + nome + ": " + e.getMessage());
            e.printStackTrace();
            return "Erro: " + e.getMessage();
        }
    }

    private static boolean linhaValida(String nome, String telefone) {
        if (nome == null || nome.trim().isEmpty()) return false;
        if (telefone == null) return false;
        return telefone.replaceAll("\\D", "").length() >= 10;
    }

    private static String lerMensagemPadrao() {
        File f = new File(ARQUIVO_MENSAGEM);
        String fallback = "Olá {nome}, este é o seu contracheque de {mes} de {ano}.";
        if (!f.exists()) return fallback;
        try (BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(f), StandardCharsets.UTF_8))) {
            StringBuilder sb = new StringBuilder(); String linha;
            while ((linha = br.readLine()) != null) sb.append(linha).append(System.lineSeparator());
            return sb.toString().trim().isEmpty() ? fallback : sb.toString().trim();
        } catch (IOException e) {
            System.err.println("❌ Não foi possível ler " + ARQUIVO_MENSAGEM + ": " + e.getMessage());
            return fallback;
        }
    }

    // ======= NOVO: salva relatório e ARQUIVA PDFs para a pasta do mês =======
    private static void salvarRelatorioEArquivar(List<RelatorioEnvio> relatorio) {
        Locale ptBR = new Locale("pt","BR");

        // MÊS DO ENVIO (quando o robô rodou). Para "mês de competência", troque para: YearMonth.now().minusMonths(1)
        YearMonth mesEnvio = YearMonth.now();

        // Pasta do Desktop → ContraX_Relatorios → "AAAA-MM nome-do-mês"
        String desktop = getDesktopPath();
        File baseRel = new File(desktop, "ContraX_Relatorios");
        String pastaMes = String.format(
                "%s %s",
                mesEnvio.format(DateTimeFormatter.ofPattern("yyyy-MM")),
                mesEnvio.getMonth().getDisplayName(TextStyle.FULL, ptBR)
        );
        File pastaDestino = new File(baseRel, pastaMes);
        if (!pastaDestino.exists() && !pastaDestino.mkdirs()) {
            System.err.println("⚠️ Não foi possível criar a pasta: " + pastaDestino.getAbsolutePath());
        }

        // Nome do CSV com timestamp
        File destinoCSV = new File(pastaDestino,
                "relatorio_envio_" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss")) + ".csv");

        // Contagem
        int enviados = 0, erros = 0, ignorados = 0;
        for (RelatorioEnvio r : relatorio) {
            switch (r.status.toUpperCase(Locale.ROOT)) {
                case "ENVIADO" -> enviados++;
                case "ERRO"    -> erros++;
                default        -> ignorados++;
            }
        }

        // Escreve CSV (com BOM p/ Excel)
        try (PrintWriter w = new PrintWriter(new OutputStreamWriter(new FileOutputStream(destinoCSV), StandardCharsets.UTF_8))) {
            w.write('\uFEFF');
            w.println("DataHora,Nome,Telefone,Status,Detalhe");
            String agora = LocalDateTime.now().format(DateTimeFormatter.ISO_LOCAL_DATE_TIME);
            for (RelatorioEnvio r : relatorio) {
                w.printf("%s,%s,%s,%s,%s%n", agora, csv(r.nome), csv(r.telefone), csv(r.status), csv(r.detalhe));
            }
            w.printf("%nTOTAL,,,%d enviados,%d erros,%d ignorados (Total: %d)%n",
                    enviados, erros, ignorados, enviados + erros + ignorados);
            System.out.println("📝 Relatório salvo em: " + destinoCSV.getAbsolutePath());
        } catch (Exception e) {
            System.err.println("Erro ao salvar relatório: " + e.getMessage());
        }

        // === Arquiva todos os PDFs do mês na mesma pasta do relatório ===
        arquivarPdfsDoMesPara(pastaDestino);

        System.out.println("📦 Arquivamento concluído. A pasta '" + PASTA_CONTRACHEQUES + "' está pronta para o próximo mês.");
    }

    private static void arquivarPdfsDoMesPara(File pastaDestino) {
        File dir = new File(PASTA_CONTRACHEQUES);
        File[] pdfs = dir.listFiles((d, n) -> n.toLowerCase().endsWith(".pdf"));

        if (pdfs == null) {
            System.out.println("ℹ️ Pasta '" + PASTA_CONTRACHEQUES + "' inexistente ou vazia. Nada para arquivar.");
            return;
        }
        int movidos = 0, falhas = 0;

        for (File f : pdfs) {
            try {
                Path alvo = pastaDestino.toPath().resolve(f.getName());
                alvo = resolveColisao(alvo);
                Files.move(f.toPath(), alvo, StandardCopyOption.REPLACE_EXISTING);
                movidos++;
            } catch (Exception e) {
                System.err.println("⚠️ Falha ao mover '" + f.getName() + "': " + e.getMessage());
                falhas++;
            }
        }

        // Tenta remover diretórios vazios dentro de contracheques (se houver)
        try {
            File[] restos = dir.listFiles();
            if (restos != null) {
                for (File r : restos) {
                    if (r.isDirectory()) {
                        try { Files.deleteIfExists(r.toPath()); } catch (Exception ignored) {}
                    }
                }
            }
        } catch (Exception ignored) {}

        System.out.println("📁 PDFs arquivados: " + movidos + " | falhas: " + falhas);
    }

    private static Path resolveColisao(Path alvo) {
        if (!Files.exists(alvo)) return alvo;
        String nome = alvo.getFileName().toString();
        String base = nome.replaceAll("(?i)\\.pdf$", "");
        String sufixo = "_" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        return alvo.getParent().resolve(base + sufixo + ".pdf");
    }

    private static String getDesktopPath() {
        String home = System.getProperty("user.home");
        List<String> candidatos = new ArrayList<>();
        candidatos.add(home + File.separator + "OneDrive" + File.separator + "Desktop");
        candidatos.add(home + File.separator + "OneDrive" + File.separator + "Área de Trabalho");
        candidatos.add(home + File.separator + "Desktop");
        candidatos.add(home + File.separator + "Área de Trabalho");
        candidatos.add("C:" + File.separator + "Users" + File.separator + "Public" + File.separator + "Desktop");
        for (String p : candidatos) {
            File d = new File(p);
            if (d.exists() && d.isDirectory()) return d.getAbsolutePath();
        }
        File fallback = new File("C:" + File.separator + "evox" + File.separator + "relatorios");
        if (!fallback.exists()) fallback.mkdirs();
        return fallback.getAbsolutePath();
    }

    private static String csv(String s) { return s == null ? "" : s.replace(",", ";"); }
    private static String getCell(Cell c) { return c == null ? "" : FORMATTER.formatCellValue(c).trim(); }

    private static String sanitizeBaseNomeArquivo(String s) {
        String out = s.trim();
        out = out.replaceAll("(?i)\\.pdf$", "");
        out = out.replaceAll("(?i)\\s+pdf$", "");
        return out;
    }

    // ======= BUSCA APROXIMADA NO DIRETÓRIO =======
    private static File localizarPdfAproximado(String base) {
        String alvo = normalizar(base).replaceAll("(?i)\\.pdf$", "");
        File dir = new File(PASTA_CONTRACHEQUES);
        File[] pdfs = dir.listFiles((d, n) -> n.toLowerCase().endsWith(".pdf"));
        if (pdfs == null || pdfs.length == 0) return null;

        // 1) normalizado igual
        for (File f : pdfs) {
            String norm = normalizar(f.getName().replaceAll("(?i)\\.pdf$", ""));
            if (norm.equals(alvo)) return f;
        }
        // 2) começa com / contém
        for (File f : pdfs) {
            String norm = normalizar(f.getName().replaceAll("(?i)\\.pdf$", ""));
            if (norm.startsWith(alvo) || norm.contains(alvo)) return f;
        }
        // 3) todos tokens do alvo presentes no nome
        List<String> tokAlvo = Arrays.asList(alvo.split("\\s+"));
        for (File f : pdfs) {
            String norm = normalizar(f.getName().replaceAll("(?i)\\.pdf$", ""));
            List<String> tokNome = Arrays.asList(norm.split("\\s+"));
            if (tokNome.containsAll(tokAlvo)) return f;
        }
        return null;
    }

    private static String normalizar(String texto) {
        return Normalizer.normalize(texto, Normalizer.Form.NFD)
                .replaceAll("\\p{InCombiningDiacriticalMarks}+", "")
                .toLowerCase()
                .trim()
                .replaceAll("[^a-z0-9]+", " ")
                .replaceAll("\\s{2,}", " ")
                .trim();
    }
}

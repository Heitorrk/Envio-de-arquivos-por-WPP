package org.contraApp;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.text.Normalizer;
import java.util.*;

public class RenomeadorPDF {

    private static final String PASTA_CONTRACHEQUES = "contracheques";
    private static final String PLANILHA = "funcionarios.xlsx";
    private static final String MAPA_ARQUIVOS = "map.csv"; // arquivo opcional para "forçar" casamentos
    private static final DataFormatter FORMATTER = new DataFormatter(new Locale("pt","BR"));

    private static final Set<String> STOP = new HashSet<>(Arrays.asList("de","da","do","das","dos","e","d","di"));

    private static class Linha {
        final String normA;      // A normalizado (nome real)
        final String normC;      // C normalizado (nome de arquivo alvo)
        final String originalA;  // A original (log)
        final String originalC;  // C original (log)
        Linha(String normA, String normC, String originalA, String originalC) {
            this.normA = normA; this.normC = normC; this.originalA = originalA; this.originalC = originalC;
        }
    }

    /** Executa antes do envio */
    public static void renomearContracheques() {
        System.out.println("[Renomeador] Iniciando…");
        System.out.println("[Renomeador] CWD = " + new File(".").getAbsolutePath());
        System.out.println("[Renomeador] Planilha = " + new File(PLANILHA).getAbsolutePath());
        System.out.println("[Renomeador] Pasta   = " + new File(PASTA_CONTRACHEQUES).getAbsolutePath());

        List<Linha> linhas = carregarAeC();
        System.out.println("[Renomeador] Linhas válidas (A e C preenchidas) = " + linhas.size());
        if (linhas.isEmpty()) {
            System.err.println("❌ Planilha vazia/ausente ou coluna C sem valores.");
            return;
        }

        Map<String, String> overrides = carregarMap();
        System.out.println("[Renomeador] Overrides (map.csv) carregados = " + overrides.size());

        File pasta = new File(PASTA_CONTRACHEQUES);
        if (!pasta.exists() || !pasta.isDirectory()) {
            System.err.println("❌ Pasta não encontrada: " + PASTA_CONTRACHEQUES);
            return;
        }

        File[] pdfs = pasta.listFiles((dir,name) -> name.toLowerCase().endsWith(".pdf"));
        int qtd = (pdfs == null) ? 0 : pdfs.length;
        System.out.println("[Renomeador] PDFs encontrados = " + qtd);
        if (qtd == 0) {
            System.err.println("⚠ Nenhum PDF em " + PASTA_CONTRACHEQUES);
            return;
        }

        for (File pdf : pdfs) {
            try (PDDocument doc = Loader.loadPDF(pdf)) {
                // 1) Tenta extrair pelo conteúdo
                String texto = new PDFTextStripper().getText(doc);
                String nomeExtraido = extrairNome(texto);
                Linha match = null;

                if (nomeExtraido != null && !nomeExtraido.isBlank()) {
                    String candNormA = normalizar(nomeExtraido);
                    match = acharPorA(candNormA, linhas); // tenta casar com a coluna A
                }

                // 2) FALLBACK: casar pelo próprio NOME DO ARQUIVO (ex.: "ramon.pdf")
                if (match == null) {
                    String baseArquivo = pdf.getName().replaceAll("(?i)\\.pdf$", "");
                    String candNormArquivo = normalizar(baseArquivo);
                    // primeiro contra A
                    match = acharPorA(candNormArquivo, linhas);
                    // se não achou, tenta contra C
                    if (match == null) match = acharPorC(candNormArquivo, linhas);
                }

                // 3) MAPA MANUAL (map.csv)
                if (match == null) {
                    String baseArquivo = pdf.getName().replaceAll("(?i)\\.pdf$", "");
                    String key = normalizar(baseArquivo);
                    if (overrides.containsKey(key)) {
                        String alvoCNorm = overrides.get(key); // já normalizado
                        Linha chosen = null;
                        for (Linha l : linhas) {
                            if (l.normC.equals(alvoCNorm)) { chosen = l; break; }
                        }
                        if (chosen != null) {
                            match = chosen;
                            System.out.println("🧭 (map.csv) " + pdf.getName() + " → coluna C \"" + chosen.originalC + "\"");
                        } else {
                            // Sem linha correspondente na planilha: renomeia mesmo assim pro alvo do mapa
                            String novoNomeArquivo = ensurePdfExtension(alvoCNorm);
                            Path destino = Paths.get(pasta.getAbsolutePath(), novoNomeArquivo);
                            try {
                                Files.move(pdf.toPath(), destino, StandardCopyOption.REPLACE_EXISTING);
                                System.out.println("✔ (map.csv sem planilha) " + pdf.getName() + " → " + novoNomeArquivo);
                            } catch (IOException ioe) {
                                System.err.println("❌ Falha ao renomear (map.csv) " + pdf.getName() + " → " + novoNomeArquivo + " (" + ioe.getMessage() + ")");
                            }
                            continue;
                        }
                    }
                }

                if (match == null) {
                    System.err.println("⚠ Não consegui casar com planilha: " + pdf.getName());
                    continue;
                }

                // Nome final: coluna C normalizada + ".pdf" (é o formato que o EnvioPDF espera)
                String novoNomeArquivo = ensurePdfExtension(match.normC);
                Path destino = Paths.get(pasta.getAbsolutePath(), novoNomeArquivo);

                if (!pdf.toPath().equals(destino)) {
                    try {
                        Files.move(pdf.toPath(), destino, StandardCopyOption.REPLACE_EXISTING);
                        System.out.println("✔ " + pdf.getName() + " → " + novoNomeArquivo +
                                "  (A: \"" + match.originalA + "\"  →  C: \"" + match.originalC + "\")");
                    } catch (IOException ioe) {
                        System.err.println("❌ Falha ao renomear " + pdf.getName() + " → " +
                                novoNomeArquivo + " (" + ioe.getMessage() + ")");
                    }
                } else {
                    System.out.println("= Já está com o nome esperado: " + pdf.getName());
                }
            } catch (IOException e) {
                System.err.println("Erro ao processar " + pdf.getName() + ": " + e.getMessage());
            }
        }
        System.out.println("[Renomeador] Finalizado.");
    }

    /** Lê colunas A (nome real) e C (nome do arquivo) */
    private static List<Linha> carregarAeC() {
        List<Linha> out = new ArrayList<>();
        try (InputStream is = new FileInputStream(PLANILHA);
             Workbook wb = new XSSFWorkbook(is)) {
            Iterator<Row> rows = wb.getSheetAt(0).iterator();
            if (rows.hasNext()) rows.next(); // cabeçalho
            while (rows.hasNext()) {
                Row r = rows.next();
                String a = FORMATTER.formatCellValue(r.getCell(0)).trim(); // A
                String c = FORMATTER.formatCellValue(r.getCell(2)).trim(); // C
                if (!a.isBlank() && !c.isBlank()) {
                    String baseC = sanitizeBaseNomeArquivo(c);
                    out.add(new Linha(normalizar(a), normalizar(baseC), a, c));
                }
            }
        } catch (FileNotFoundException e) {
            System.err.println("❌ Planilha '" + PLANILHA + "' não encontrada.");
        } catch (Exception e) {
            System.err.println("❌ Falha ao ler planilha: " + e.getMessage());
        }
        return out;
    }

    private static Map<String, String> carregarMap() {
        Map<String, String> map = new HashMap<>();
        File f = new File(MAPA_ARQUIVOS);
        if (!f.exists()) return map;
        try (BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(f), java.nio.charset.StandardCharsets.UTF_8))) {
            String line;
            while ((line = br.readLine()) != null) {
                line = line.trim();
                if (line.isEmpty() || line.startsWith("#")) continue;
                String[] parts = line.split(",", 2);
                if (parts.length < 2) continue;
                String arquivo = parts[0].trim();
                String alvoC   = parts[1].trim();
                if (!arquivo.isEmpty() && !alvoC.isEmpty()) {
                    String key = normalizar(arquivo.replaceAll("(?i)\\.pdf$", ""));
                    String val = normalizar(alvoC.replaceAll("(?i)\\.pdf$", ""));
                    map.put(key, val);
                }
            }
        } catch (Exception e) {
            System.err.println("[Renomeador] ⚠ Falha ao ler map.csv: " + e.getMessage());
        }
        return map;
    }

    // ===== MATCH ROBUSTO (A e C) =====
    private static Linha acharPorA(String candNormA, List<Linha> linhas) {
        List<String> candTok = sigTokens(candNormA);
        Linha melhor = null; int melhorScore = -1; int melhorLev = Integer.MAX_VALUE;

        for (Linha l : linhas) {
            List<String> tok = sigTokens(l.normA);
            int score = scoreTokens(candTok, tok);
            int lev = dist(candNormA, l.normA);
            if (score > melhorScore || (score == melhorScore && lev < melhorLev)) {
                melhor = l; melhorScore = score; melhorLev = lev;
            }
        }
        if (melhor == null) return null;
        if (melhorScore >= 80) return melhor;
        if (melhorScore >= 60) return melhor;
        int maxLen = Math.max(candNormA.length(), melhor.normA.length());
        return (melhorLev <= Math.ceil(maxLen * 0.30)) ? melhor : null;
    }

    private static Linha acharPorC(String candNorm, List<Linha> linhas) {
        List<String> candTok = sigTokens(candNorm);
        Linha melhor = null; int melhorScore = -1; int melhorLev = Integer.MAX_VALUE;

        for (Linha l : linhas) {
            List<String> tok = sigTokens(l.normC);
            int score = scoreTokens(candTok, tok);
            int lev = dist(candNorm, l.normC);
            if (score > melhorScore || (score == melhorScore && lev < melhorLev)) {
                melhor = l; melhorScore = score; melhorLev = lev;
            }
        }
        if (melhor == null) return null;
        if (melhorScore >= 80) return melhor;
        if (melhorScore >= 60) return melhor;
        int maxLen = Math.max(candNorm.length(), melhor.normC.length());
        return (melhorLev <= Math.ceil(maxLen * 0.30)) ? melhor : null;
    }

    private static List<String> sigTokens(String norm) {
        List<String> out = new ArrayList<>();
        for (String t : norm.split("\\s+")) {
            if (t.isBlank()) continue;
            if (STOP.contains(t)) continue;
            out.add(t);
        }
        return out;
    }
    private static int scoreTokens(List<String> a, List<String> b) {
        if (a.isEmpty() || b.isEmpty()) return 0;
        if (a.size() >= 2 && b.size() >= 2 && a.get(0).equals(b.get(0)) && a.get(1).equals(b.get(1))) return 100;
        int overlap = overlapCount(a, b);
        if (a.get(0).equals(b.get(0)) && overlap >= Math.min(3, Math.min(a.size(), b.size()))) return 80;
        if (a.get(0).equals(b.get(0))) return 60;
        if (overlap >= 3) return 60;
        if (overlap == 2) return 50;
        if (overlap == 1) return 30;
        return 0;
    }
    private static int overlapCount(List<String> a, List<String> b) {
        Set<String> set = new HashSet<>(b);
        int c = 0;
        for (String t : a) if (set.contains(t)) c++;
        return c;
    }
    private static int dist(String s1,String s2){
        int[][] dp=new int[s1.length()+1][s2.length()+1];
        for(int i=0;i<=s1.length();i++) dp[i][0]=i;
        for(int j=0;j<=s2.length();j++) dp[0][j]=j;
        for(int i=1;i<=s1.length();i++){
            for(int j=1;j<=s2.length();j++){
                int cost=(s1.charAt(i-1)==s2.charAt(j-1))?0:1;
                dp[i][j]=Math.min(Math.min(dp[i-1][j]+1, dp[i][j-1]+1), dp[i-1][j-1]+cost);
            }
        }
        return dp[s1.length()][s2.length()];
    }

    // ===== EXTRAÇÃO DO NOME NO PDF (quando houver texto) =====
    private static String extrairNome(String texto) {
        if (texto == null || texto.isBlank()) return null;
        String[] linhas = texto.split("\\R");
        for (int i = 0; i < linhas.length; i++) {
            String linha = linhas[i].trim();
            String lower = linha.toLowerCase();
            if (lower.matches("^(empregado|nome)\\b.*")) {
                String candidato = linha.replaceAll("(?i)^(empregado|nome)\\b\\s*[:\\-–]?\\s*", "").trim();
                if (candidato.isBlank() || candidato.equalsIgnoreCase("empregado") || candidato.equalsIgnoreCase("nome")) {
                    if (i + 1 < linhas.length) candidato = linhas[i + 1].trim();
                }
                candidato = candidato.replaceFirst("^\\d+\\s*", "").trim(); // remove matrícula
                if (!candidato.isBlank()) return candidato;
            }
        }
        return null;
    }

    // ===== NORMALIZAÇÃO / EXTENSÃO =====
    private static String sanitizeBaseNomeArquivo(String s) {
        String out = s.trim();
        out = out.replaceAll("(?i)\\.pdf$", "");
        out = out.replaceAll("(?i)\\s+pdf$", "");
        return out;
    }
    private static String ensurePdfExtension(String baseNormalized) {
        String b = baseNormalized.replaceAll("(?i)_pdf$", "");
        return b + ".pdf";
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
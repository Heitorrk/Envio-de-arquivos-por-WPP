package org.contraApp;

import java.io.File;
import java.time.Duration;
import java.time.Instant;

public class Main {
    public static void main(String[] args) {
        Instant inicio = Instant.now();
        try {
            System.out.println("▶️ contraApp iniciado");
            System.out.println("CWD: " + new File(".").getAbsolutePath());
            System.out.println("Java: " + System.getProperty("java.version"));
            System.out.println("Charset: " + System.getProperty("file.encoding"));

            // 1) RENOMEIA PDFs com base na planilha/pasta padrão
            System.out.println("🔄 Iniciando renomeação de PDFs...");
            RenomeadorPDF.renomearContracheques();

            // 2) ENVIA pelo WhatsApp Web
            System.out.println("📤 Iniciando envio de contracheques...");
            EnvioPDF.main(args); // (não chamar Main.main(args) aqui!)
        } catch (Throwable t) {
            System.err.println("❌ Erro inesperado: " + t.getMessage());
            t.printStackTrace();
            System.exit(1);
        } finally {
            long ms = Duration.between(inicio, Instant.now()).toMillis();
            System.out.println("✅ Finalizado em " + ms + " ms");
        }
    }
}
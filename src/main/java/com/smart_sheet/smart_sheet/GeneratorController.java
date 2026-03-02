package com.smart_sheet.smart_sheet;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

@RestController
@CrossOrigin(origins = "*") // Permite que tu frontend Angular se conecte sin problemas de CORS
@RequestMapping("/smart-sheet-reporter/api/beta/generator")
public class GeneratorController {

    @Autowired
    private DocumentGenerator documentGenerator;

    @PostMapping(value = "/generate", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<byte[]> generateDocuments(
            @RequestParam("excel") MultipartFile excelFile,
            @RequestParam("template") MultipartFile templateFile) {
        try {

            byte[] zipContent = documentGenerator.generateZipBundle(
                    excelFile.getInputStream(),
                    templateFile.getInputStream()
            );

            return ResponseEntity.ok()
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"procesados.zip\"")
                    .contentType(MediaType.parseMediaType("application/zip"))
                    .body(zipContent);

        } catch (Exception e) {
            // 1. IMPRIME EL ERROR EN LA CONSOLA (¡Fundamental para depurar!)
            e.printStackTrace();

            // 2. DEVUELVE UN ARCHIVO DE TEXTO CON EL ERROR EN VEZ DE UN ZIP VACÍO
            String mensajeError = "Ocurrió un error en el servidor:\n" + e.getMessage();
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"error_log.txt\"")
                    .contentType(MediaType.TEXT_PLAIN)
                    .body(mensajeError.getBytes());
        }
    }
}
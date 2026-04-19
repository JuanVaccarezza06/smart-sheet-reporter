package com.smart_sheet.smart_sheet;

import fr.opensagres.xdocreport.document.IXDocReport;
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry;
import fr.opensagres.xdocreport.template.IContext;
import fr.opensagres.xdocreport.template.TemplateEngineKind;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.*;

@Service
public class DocumentGenerator {

    public byte[] generateZipBundle(InputStream excelIs, InputStream templateIs) throws Exception {

        List<DataRow> dataRows = ExcelUtils.readExcel(excelIs);

        // 1. Cargamos la plantilla UNA SOLA VEZ
        IXDocReport report = XDocReportRegistry.getRegistry().loadReport(templateIs, TemplateEngineKind.Velocity);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();

        try (ZipOutputStream zos = new ZipOutputStream(baos)) {

            // NUEVO: Mapa para rastrear cuántas veces ha aparecido un nombre
            Map<String, Integer> nameTracker = new HashMap<>();

            for (DataRow row : dataRows) {

                // NUEVO: Detectar y omitir "filas fantasma" (todas sus celdas están vacías)
                boolean isRowEmpty = row.getData().values().stream()
                        .allMatch(value -> value == null || value.trim().isEmpty());

                if (isRowEmpty) {
                    continue; // Ignora esta fila y pasa a la siguiente
                }

                IContext context = report.createContext();
                row.getData().forEach(context::put);

                ByteArrayOutputStream out = new ByteArrayOutputStream();
                report.process(context, out);

                // NUEVO: Lógica robusta para nombres duplicados
                String rawName = row.getValueByName("Nombre"); // Asegúrate de usar el método correcto de tu DataRow
                String fileName;

                if (rawName != null && !rawName.trim().isEmpty()) {
                    String baseName = rawName.trim();

                    // Verificamos si el nombre ya existe en nuestro rastreador
                    if (nameTracker.containsKey(baseName)) {
                        int count = nameTracker.get(baseName) + 1;
                        nameTracker.put(baseName, count);
                        // Le agregamos el sufijo al nombre duplicado
                        fileName = baseName + "_" + count + ".docx";
                    } else {
                        // Es la primera vez que vemos este nombre
                        nameTracker.put(baseName, 0);
                        fileName = baseName + ".docx";
                    }
                } else {
                    // Fallback si la celda estaba completamente vacía
                    fileName = "Documento_Sin_Nombre_" + System.currentTimeMillis() + ".docx";
                }

                ZipEntry entry = new ZipEntry(fileName);
                zos.putNextEntry(entry);
                zos.write(out.toByteArray());
                zos.closeEntry();
            }
        }
        return baos.toByteArray();
    }

    // NOTA: Puedes borrar el método privado processTemplate(), ya integré la lógica arriba para hacerlo más limpio.

    private byte[] processTemplate(DataRow row, InputStream templateStream) throws Exception {
        // Cargamos la plantilla con el motor Velocity
        IXDocReport report = XDocReportRegistry.getRegistry().loadReport(templateStream, TemplateEngineKind.Velocity);
        IContext context = report.createContext();

        // Mapeamos cada llave del Excel a una variable del Word
        row.getData().forEach(context::put);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        report.process(context, out);
        return out.toByteArray();
    }
}
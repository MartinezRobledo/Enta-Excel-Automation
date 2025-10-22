package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;

import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.annotations.rules.SessionObject;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;
import java.net.URI;
import java.nio.file.*;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@BotCommand
@CommandPkg(
        label = "Paste Image To Excel",
        name = "pasteImageToExcel",
        description = "Pega una imagen en la hoja de Excel en una celda específica, ajustando tamaño",
        icon = "excel.svg"
)
public class PasteImageToExcel {

    // MsoTriState (Excel): msoTrue = -1, msoFalse = 0
    private static final int MSO_TRUE = -1;
    private static final int MSO_FALSE = 0;

    @Execute
    public void action(
            @Idx(index = "1", type = AttributeType.SESSION)
            @Pkg(label = "Workbook Session") @NotEmpty @SessionObject ExcelSession excelSession,

            @Idx(index = "2", type = AttributeType.TEXT)
            @Pkg(label = "Sheet Name") @NotEmpty String sheetName,

            @Idx(index = "3", type = AttributeType.FILE)
            @Pkg(label = "Image File Path") @NotEmpty String imagePath,

            @Idx(index = "4", type = AttributeType.TEXT)
            @Pkg(label = "Start Cell (e.g., B2)", default_value = "A1", default_value_type = DataType.STRING)
            @NotEmpty String startCell,

            @Idx(index = "5", type = AttributeType.NUMBER)
            @Pkg(label = "Width in cells", default_value = "3", default_value_type = DataType.NUMBER)
            @NotEmpty Double widthCells,

            @Idx(index = "6", type = AttributeType.NUMBER)
            @Pkg(label = "Height in cells", default_value = "4", default_value_type = DataType.NUMBER)
            @NotEmpty Double heightCells
    ) {
        try {
            // 1) Sanitizar y resolver ruta
            String resolvedPath = resolveImagePath(imagePath);

            // 2) Validar existencia (antes de ir a Excel)
            File f = new File(resolvedPath);
            if (!f.exists() || !f.isFile()) {
                throw new BotCommandException("No se encontró el archivo de imagen: " + resolvedPath);
            }

            // 3) Objetos Excel
            Session session = ExcelObjects.requireSession(excelSession);
            Dispatch wb     = ExcelObjects.requireWorkbook(session, excelSession);
            Dispatch sheet  = Dispatch.call(Dispatch.get(wb, "Sheets").toDispatch(), "Item", sheetName).toDispatch();

            // Activar hoja (evita ambigüedades en algunas instalaciones)
            Dispatch.call(sheet, "Activate");

            // 4) Rango base
            Dispatch start = Dispatch.call(sheet, "Range", startCell).toDispatch();

            // Calcular rango de celdas destino (ancho x alto), p. ej., B2:D5
            int w = Math.max(1, widthCells.intValue());
            int h = Math.max(1, heightCells.intValue());

            // Celda final: Offset(h-1, w-1)
            Dispatch end = Dispatch.call(start, "Offset", h - 1, w - 1).toDispatch();

            // Rango total (para tomar ancho/alto exactos en pixeles)
            Dispatch target = Dispatch.invoke(
                    sheet, "Range", Dispatch.Get, new Object[]{ start, end }, new int[1]
            ).toDispatch();

            double left     = Dispatch.get(start,  "Left").getDouble();
            double top      = Dispatch.get(start,  "Top").getDouble();
            double imgWidth = Dispatch.get(target, "Width").getDouble();
            double imgHeight= Dispatch.get(target, "Height").getDouble();

            // 5) Insertar imagen
            Dispatch shapes = Dispatch.get(sheet, "Shapes").toDispatch();
            Dispatch pic = null;
            Exception addPicError = null;

            try {
                // AddPicture(Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height)
                pic = Dispatch.call(shapes, "AddPicture",
                        new Variant(resolvedPath),
                        new Variant(MSO_FALSE),  // LinkToFile = false
                        new Variant(MSO_TRUE),   // SaveWithDocument = true
                        new Variant(left),
                        new Variant(top),
                        new Variant(imgWidth),
                        new Variant(imgHeight)
                ).toDispatch();
            } catch (Exception e) {
                addPicError = e;
            }

            // 6) Fallback si AddPicture falló (p. ej. ciertos tipos/instalaciones)
            if (pic == null) {
                try {
                    Dispatch pictures = Dispatch.get(sheet, "Pictures").toDispatch();
                    // Pictures.Insert devuelve un Picture; luego posicionamos y redimensionamos
                    pic = Dispatch.call(pictures, "Insert", resolvedPath).toDispatch();

                    Dispatch.put(pic, "Left",  new Variant(left));
                    Dispatch.put(pic, "Top",   new Variant(top));
                    Dispatch.put(pic, "Width", new Variant(imgWidth));
                    Dispatch.put(pic, "Height",new Variant(imgHeight));
                } catch (Exception e2) {
                    String detail = (addPicError != null ? addPicError.getMessage() : "n/a")
                            + " | fallback: " + e2.getMessage();
                    throw new BotCommandException(
                            "No se pudo insertar la imagen. Ruta: " + resolvedPath + ". Detalle: " + detail, e2
                    );
                }
            }

        } catch (BotCommandException e) {
            throw e;
        } catch (Exception e) {
            throw new BotCommandException("PasteImageToExcel failed: " + e.getMessage(), e);
        }
    }

    // --- Helpers de ruta ---

    private static String resolveImagePath(String raw) throws Exception {
        if (raw == null) throw new BotCommandException("Image path vacío.");

        String p = raw.trim();

        // quitar comillas envolventes
        if ((p.startsWith("\"") && p.endsWith("\"")) || (p.startsWith("'") && p.endsWith("'"))) {
            p = p.substring(1, p.length() - 1).trim();
        }

        // si viene como file:///
        if (p.startsWith("file:/")) {
            URI uri = URI.create(p);
            Path path = Paths.get(uri);
            p = path.toFile().getAbsolutePath();
        }

        // expandir variables de entorno Windows %VAR%
        p = expandEnvVars(p);

        // normalizar y absolutizar
        Path path = Paths.get(p);
        if (!path.isAbsolute()) {
            path = path.toAbsolutePath();
        }
        path = path.normalize();

        // reemplazar / por \ (Excel suele preferir)
        String normalized = path.toString().replace('/', '\\');

        // Si es Long Path, copiar a %TEMP% con nombre corto para evitar líos
        if (normalized.length() > 250) {
            Path tempDir = Paths.get(System.getProperty("java.io.tmpdir"));
            String fileName = Paths.get(normalized).getFileName().toString();
            Path tempTarget = tempDir.resolve("img_" + System.currentTimeMillis() + "_" + fileName);
            Files.copy(Paths.get(normalized), tempTarget, StandardCopyOption.REPLACE_EXISTING);
            normalized = tempTarget.toString();
        }

        return normalized;
    }

    private static String expandEnvVars(String input) {
        if (input == null) return null;
        Pattern p = Pattern.compile("%([^%]+)%"); // %USERPROFILE% etc.
        Matcher m = p.matcher(input);
        StringBuffer sb = new StringBuffer();
        Map<String,String> env = System.getenv();
        while (m.find()) {
            String var = m.group(1);
            String val = env.getOrDefault(var, m.group());
            // escapar \ para appendReplacement
            val = val.replace("\\", "\\\\");
            m.appendReplacement(sb, val);
        }
        m.appendTail(sb);
        return sb.toString();
    }
}
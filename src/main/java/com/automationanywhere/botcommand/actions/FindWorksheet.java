package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.utilities.ExcelObjects;
import com.automationanywhere.botcommand.utilities.ExcelSession;
import com.automationanywhere.botcommand.utilities.Session;

import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.*;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

import com.automationanywhere.botcommand.data.impl.StringValue; // <-- IMPORTANTE
import com.automationanywhere.botcommand.data.Value;            // opcional si lo querés usar

import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.util.regex.Pattern;

@BotCommand
@CommandPkg(
        label = "Find Worksheet",
        name = "findWorksheet",
        description = "Busca y devuelve el nombre de la primera hoja que coincide con un patrón (exact, includes o wildcard). Si no encuentra, retorna vacío.",
        icon = "excel.svg",
        return_label = "Worksheet Name",
        return_description = "Nombre de la primera hoja encontrada o vacío si no hay coincidencias",
        return_type = DataType.STRING
)
public class FindWorksheet {

    @Execute
    public StringValue action(     // <-- ahora devuelve StringValue
                                   @Idx(index = "1", type = AttributeType.SESSION)
                                   @Pkg(label = "Workbook Session")
                                   @NotEmpty @SessionObject ExcelSession excelSession,

                                   @Idx(index = "2", type = AttributeType.SELECT, options = {
                                           @Idx.Option(index = "2.1", pkg = @Pkg(label = "Exact",    value = "exact")),
                                           @Idx.Option(index = "2.2", pkg = @Pkg(label = "Includes", value = "includes")),
                                           @Idx.Option(index = "2.3", pkg = @Pkg(label = "Wildcard", value = "wildcard"))
                                   })
                                   @Pkg(label = "Search Mode", default_value = "includes", default_value_type = DataType.STRING)
                                   @SelectModes String searchMode,

                                   @Idx(index = "3", type = AttributeType.TEXT)
                                   @Pkg(label = "Search Text")
                                   @NotEmpty String searchText,

                                   @Idx(index = "4", type = AttributeType.CHECKBOX)
                                   @Pkg(label = "Case Sensitive?", default_value_type = DataType.BOOLEAN, default_value = "false")
                                   Boolean caseSensitive
    ) {
        try {
            String result = run(excelSession, searchMode, searchText, caseSensitive);
            // Tu requerimiento: si no encuentra => vacío (no null)
            if (result == null) result = "";
            return new StringValue(result);
        } catch (Exception first) {
            // Fallback por si el hilo no tenía COM inicializado
            try {
                ComThread.InitSTA();
                String result = run(excelSession, searchMode, searchText, caseSensitive);
                if (result == null) result = "";
                return new StringValue(result);
            } catch (Exception second) {
                throw (second instanceof BotCommandException)
                        ? (BotCommandException) second
                        : new BotCommandException("FindWorksheet failed: " + second.getMessage(), second);
            } finally {
                try { ComThread.Release(); } catch (Exception ignore) {}
            }
        }
    }

    private String run(
            ExcelSession excelSession, String searchMode, String searchText, Boolean caseSensitive
    ) {
        if (searchText == null || searchText.trim().isEmpty()) {
            return "";
        }

        String mode = (searchMode == null) ? "includes" : searchMode.trim().toLowerCase();
        if (!("exact".equals(mode) || "includes".equals(mode) || "wildcard".equals(mode))) {
            throw new BotCommandException("Search Mode inválido. Usa: exact | includes | wildcard.");
        }

        Session session = ExcelObjects.requireSession(excelSession);
        Dispatch wb = ExcelObjects.requireWorkbook(session, excelSession);

        Dispatch sheets = Dispatch.get(wb, "Worksheets").toDispatch();
        int count = Dispatch.get(sheets, "Count").getInt();

        final boolean cs = Boolean.TRUE.equals(caseSensitive);
        String needle = searchText.trim();

        Pattern wildcardPattern = null;
        if ("wildcard".equals(mode)) {
            wildcardPattern = buildWildcardRegex(needle, cs);
        }

        for (int i = 1; i <= count; i++) {
            Dispatch sheet = Dispatch.call(sheets, "Item", new Variant(i)).toDispatch();
            String name = Dispatch.get(sheet, "Name").getString();
            if (matches(name, needle, mode, cs, wildcardPattern)) {
                return name; // primera coincidencia
            }
        }
        return "";
    }

    private boolean matches(String name, String needle, String mode, boolean caseSensitive, Pattern wildcardPattern) {
        if ("wildcard".equals(mode)) {
            // Para wildcard usamos el regex con flags ⇒ no necesitamos normalizar acá
            return wildcardPattern.matcher(name).matches();
        } else {
            if (!caseSensitive) {
                name = name.toLowerCase();
                needle = needle.toLowerCase();
            }
            switch (mode) {
                case "exact":    return name.equals(needle);
                case "includes": return name.contains(needle);
                default:         return false;
            }
        }
    }

    /**
     * Convierte patrón wildcard a regex:
     *  - '*' => '.*'
     *  - '?' => '.'
     *  El resto se escapa.
     */
    private Pattern buildWildcardRegex(String pattern, boolean caseSensitive) {
        StringBuilder sb = new StringBuilder("^");
        for (int i = 0; i < pattern.length(); i++) {
            char ch = pattern.charAt(i);
            switch (ch) {
                case '*': sb.append(".*"); break;
                case '?': sb.append(".");  break;
                default:  sb.append(java.util.regex.Pattern.quote(String.valueOf(ch)));
            }
        }
        sb.append("$");
        int flags = caseSensitive ? 0 : Pattern.CASE_INSENSITIVE;
        return Pattern.compile(sb.toString(), flags);
    }
}
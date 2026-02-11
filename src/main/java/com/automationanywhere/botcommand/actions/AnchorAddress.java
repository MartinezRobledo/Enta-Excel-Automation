package com.automationanywhere.botcommand.actions;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.AttributeType;
import com.automationanywhere.commandsdk.model.DataType;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

@BotCommand
@CommandPkg(
        name = "anchorAddress",
        label = "Anclar referencia ($)",
        description = "Devuelve una celda o rango con $ en columna y/o fila según las opciones",
        icon = "excel.svg",
        node_label = "Anclar «{{reference}}» (Cols={{fixCols}}, Filas={{fixRows}})",
        return_label = "Referencia anclada",
        return_type = DataType.STRING,
        return_required = true
)
public class AnchorAddress {

    // Soporta prefijo opcional de hoja/libro:  'Hoja 1'!A1
    // Capturas: [1]=sheetQuoted, [2]=sheetPlain, [3]=$col?, [4]=COL, [5]=$row?, [6]=row
    private static final Pattern CELL_RE = Pattern.compile(
            "^(?:(?:'([^']+)')|([^'!]+))!?\\s*(\\$?)([A-Za-z]+)(\\$?)(\\d+)$"
    );

    @Execute
    public Value<String> action(
            @Idx(index = "1", type = AttributeType.TEXT)
            @Pkg(label = "Celda o Rango (ej. A1 o A1:C6)")
            @NotEmpty String reference,

            @Idx(index = "2", type = AttributeType.CHECKBOX)
            @Pkg(label = "Fijar columnas", default_value = "false", default_value_type = DataType.BOOLEAN)
            Boolean fixCols,

            @Idx(index = "3", type = AttributeType.CHECKBOX)
            @Pkg(label = "Fijar filas", default_value = "false", default_value_type = DataType.BOOLEAN)
            Boolean fixRows
    ) {
        try {
            String ref = reference.trim();
            if (ref.isEmpty()) {
                throw new BotCommandException("La referencia no puede estar vacía.");
            }

            // ¿Es rango?
            int sep = ref.indexOf(':');
            if (sep >= 0) {
                String left = ref.substring(0, sep).trim();
                String right = ref.substring(sep + 1).trim();

                ParsedCell a = parseCell(left);
                ParsedCell b = parseCellWithOptionalSheet(right, a.sheet); // hereda hoja si right no la trae
                String outLeft  = build(a, fixCols, fixRows);
                String outRight = build(b, fixCols, fixRows);

                // Si a trae hoja, la conservamos sólo en el primer lado (estilo Excel)
                if (a.sheet != null && !a.sheet.isEmpty()) {
                    outRight = stripLeadingSheet(outRight);
                }
                return new StringValue(outLeft + ":" + outRight);
            } else {
                ParsedCell c = parseCell(ref);
                return new StringValue(build(c, fixCols, fixRows));
            }

        } catch (BotCommandException e) {
            throw e;
        } catch (Exception e) {
            throw new BotCommandException("AnchorAddress falló: " + e.getMessage(), e);
        }
    }

    // ---- Helpers ----

    private static ParsedCell parseCell(String s) {
        String t = s.trim();
        Matcher m = CELL_RE.matcher(t);
        if (!m.matches()) {
            throw new BotCommandException("Referencia inválida: «" + s + "». Use formato como A1 o 'Hoja 1'!C3.");
        }
        String sheet = (m.group(1) != null) ? m.group(1)
                : (m.group(2) != null) ? m.group(2) : null;
        String colDollar = m.group(3);
        String col = m.group(4).toUpperCase();
        String rowDollar = m.group(5);
        String row = m.group(6);

        return new ParsedCell(sheet, colDollar != null && !colDollar.isEmpty(), col,
                rowDollar != null && !rowDollar.isEmpty(), row);
    }

    // Si el lado derecho no incluye hoja, hereda la del lado izquierdo
    private static ParsedCell parseCellWithOptionalSheet(String s, String inheritSheet) {
        try {
            return parseCell(s);
        } catch (BotCommandException ex) {
            // Intentar parsear sin hoja y agregarla
            Matcher m = Pattern.compile("^(\\$?)([A-Za-z]+)(\\$?)(\\d+)$").matcher(s.trim());
            if (!m.matches()) throw ex;
            String colDollar = m.group(1);
            String col = m.group(2).toUpperCase();
            String rowDollar = m.group(3);
            String row = m.group(4);
            return new ParsedCell(inheritSheet,
                    colDollar != null && !colDollar.isEmpty(), col,
                    rowDollar != null && !rowDollar.isEmpty(), row);
        }
    }

    private static String build(ParsedCell c, Boolean fixCols, Boolean fixRows) {
        boolean colAbs = Boolean.TRUE.equals(fixCols);
        boolean rowAbs = Boolean.TRUE.equals(fixRows);

        String col = (colAbs ? "$" : "") + c.col;
        String row = (rowAbs ? "$" : "") + c.row;

        String body = col + row;
        if (c.sheet == null || c.sheet.isEmpty()) {
            return body;
        }
        // Si la hoja contiene espacios o caracteres especiales, usar comillas simples
        String sheetPart = c.sheet.contains(" ") || c.sheet.contains("-") || c.sheet.contains(".") || c.sheet.contains("[")
                || c.sheet.contains("]") || c.sheet.contains("!") ? "'" + c.sheet + "'" : c.sheet;
        return sheetPart + "!" + body;
    }

    private static String stripLeadingSheet(String ref) {
        // Elimina «Hoja!» o «'Hoja 1'!» al inicio
        if (ref.indexOf('!') > 0) {
            return ref.substring(ref.indexOf('!') + 1);
        }
        return ref;
    }

    private static class ParsedCell {
        final String sheet;
        final boolean colHasDollar;
        final String col;
        final boolean rowHasDollar;
        final String row;

        ParsedCell(String sheet, boolean colHasDollar, String col, boolean rowHasDollar, String row) {
            this.sheet = sheet;
            this.colHasDollar = colHasDollar;
            this.col = col;
            this.rowHasDollar = rowHasDollar;
            this.row = row;
        }
    }
}
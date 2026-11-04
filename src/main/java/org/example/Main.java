package org.example;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileOutputStream;

public class Main {
    static void main() {
        try {
            Workbook libro = new XSSFWorkbook();
            Sheet hoja = libro.createSheet("EStudiante");
            Row encabezado = hoja.createRow(0);
            encabezado.createCell(0).setCellValue("codigo");
            encabezado.createCell(1).setCellValue("estudiante");
            encabezado.createCell(2).setCellValue("carrera");
            encabezado.createCell(3).setCellValue("nota");

            Row fila1 = hoja.createRow(1);
            fila1.createCell(0).setCellValue("202623315");
            fila1.createCell(1).setCellValue("framk yorvi");
            fila1.createCell(2).setCellValue("sistemas");
            fila1.createCell(3).setCellValue("20");

            Row fila2 = hoja.createRow(2);
            fila2.createCell(0).setCellValue("202624254");
            fila2.createCell(1).setCellValue("maycol");
            fila2.createCell(2).setCellValue("civil");
            fila2.createCell(3).setCellValue("00");

            for (int i = 0; i < 4; i++) {
                hoja.autoSizeColumn(i);
            }

            FileOutputStream archivo = new FileOutputStream("reporte.xlsx");
            libro.write(archivo);
            archivo.close();
            libro.close();
            System.out.println("reporte generado correctamente");
        } catch (Exception e) {
            System.out.println("error " + e.getMessage());
        }
    }
}
/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.archivoss;

/**
 *
 * @author USUARIO
 */
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import static jdk.xml.internal.SecuritySupport.isFileExists;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author nicolas
 */
public class Main {

    public static void main(String args[]) {
        try {
            String rutaArchivoExcel = "C:\\Users\\USUARIO\\Videos\\resultados.xlsx";
            FileInputStream inputStream = new FileInputStream(new File(rutaArchivoExcel));
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator iterator = firstSheet.iterator();

            DataFormatter formatter = new DataFormatter();
            while (iterator.hasNext()) {
                Row nextRow = (Row) iterator.next();
                Iterator cellIterator = nextRow.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = (Cell) cellIterator.next();
                    String contenidoCelda = formatter.formatCellValue(cell);
                    System.out.println(contenidoCelda);

                    String filePath = (contenidoCelda);
                    File file = new File(filePath);

                    if (isFileExists(file)) {
                    } else {
                        System.out.println("el archivo no existe");
                    }

                    File origen = new File(contenidoCelda);
                    File destino = new File("C:/Users/USUARIO/Videos/Resultados/" + file.getName());

                    try {
                        InputStream in = new FileInputStream(origen);
                        OutputStream out = new FileOutputStream(destino);

                        byte[] buf = new byte[1024];
                        int len;

                        while ((len = in.read(buf)) > 0) {
                            out.write(buf, 0, len);
                        }

                        in.close();
                        out.close();
                        
                    } catch (IOException ioe) {
                        ioe.printStackTrace();
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

}
